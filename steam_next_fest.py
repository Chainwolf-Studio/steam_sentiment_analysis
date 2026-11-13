#!/usr/bin/env python3
"""Colectează recenziile din scorul Steam și declarațiile AI pentru Next Fest."""
import argparse
import csv
import os
import sqlite3
import sys
from concurrent.futures import ThreadPoolExecutor
from contextlib import closing
from pathlib import Path

from steam_catalog import Client, Workbook, WriteOnlyCell, disclosure_from_page, now, store_items

ROSTER = Path(__file__).resolve().parent / "data/steam-next-fest-october-2026-app-ids.csv"
COLUMNS = ("appid", "name", "review_count", "reviews_status", "ai_disclosure_status",
           "ai_disclosure", "store_url", "country", "reviews_checked_at_utc",
           "disclosure_checked_at_utc")


def load_appids(path):
    with path.open(encoding="utf-8-sig", newline="") as source:
        rows = csv.DictReader(source)
        if "app_id" not in (rows.fieldnames or []):
            raise ValueError("CSV-ul trebuie să aibă coloana app_id")
        appids = sorted({int(row["app_id"] or "") for row in rows})
    if not appids or appids[0] <= 0:
        raise ValueError("CSV-ul trebuie să conțină AppID-uri pozitive")
    return appids


def collect_reviews(db, client, appids, country):
    for offset in range(0, len(appids), 100):
        batch = appids[offset:offset + 100]
        try:
            items = store_items(client, batch, country)
        except (RuntimeError, KeyError, TypeError, ValueError) as error:
            print(f"Lotul de recenzii de la AppID {batch[0]}: {error}", flush=True)
            items = {}
        checked_at = now()
        with db:
            for appid in batch:
                item = items.get(appid, {})
                summary = item.get("reviews", {}).get("summary_filtered", {})
                count = summary.get("review_count") if item.get("success") == 1 else None
                status = "store_filtered" if type(count) is int and count >= 0 else "unavailable"
                db.execute("""UPDATE next_fest SET name=?, review_count=?, reviews_status=?,
                              reviews_checked_at_utc=? WHERE appid=? AND country=?""",
                           (item.get("name"), count if status == "store_filtered" else None,
                            status, checked_at, appid, country))
        print(f"Recenzii: {offset + len(batch)}/{len(appids)}", flush=True)


def collect_disclosures(db, client, appids, country):
    def inspect(appid):
        status, text = "unavailable", None
        try:
            url, html = client.get(f"https://store.steampowered.com/app/{appid}/",
                                   {"l": "english", "cc": country}, page=True)
            parsed_status, text = disclosure_from_page(appid, url, html)
            if parsed_status == "absent" or (parsed_status == "present" and text):
                status = parsed_status
        except RuntimeError:
            pass
        return status, text or None, now(), appid, country

    workers = ThreadPoolExecutor(max_workers=8)
    try:
        for offset in range(0, len(appids), 100):
            batch = appids[offset:offset + 100]
            for record in workers.map(inspect, batch):
                with db:
                    db.execute("""UPDATE next_fest SET ai_disclosure_status=?, ai_disclosure=?,
                                  disclosure_checked_at_utc=? WHERE appid=? AND country=?""", record)
            print(f"Declarații AI: {offset + len(batch)}/{len(appids)}", flush=True)
    finally:
        workers.shutdown(wait=False, cancel_futures=True)


def export(db, output, appids, country):
    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet("Steam Next Fest")
    sheet.freeze_panes = "B2"
    for column, width in zip("ABCDEFGHIJ", (13, 45, 15, 21, 23, 90, 49, 11, 30, 30)):
        sheet.column_dimensions[column].width = width
    sheet.append(COLUMNS)
    for appid in appids:
        row = db.execute(f"SELECT {', '.join(COLUMNS)} FROM next_fest WHERE appid=? AND country=?",
                         (appid, country)).fetchone()
        cells = []
        for value in row:
            if isinstance(value, str):
                cell = WriteOnlyCell(sheet, value=value)
                cell.data_type = "s"
                cells.append(cell)
            else:
                cells.append(value)
        sheet.append(cells)
    sheet.auto_filter.ref = f"A1:J{len(appids) + 1}"
    temp = output.with_suffix(".tmp.xlsx")
    try:
        workbook.save(temp)
        os.replace(temp, output)
    finally:
        temp.unlink(missing_ok=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--ids", type=Path, default=ROSTER)
    parser.add_argument("--db", type=Path, default=Path("steam_next_fest.sqlite3"))
    parser.add_argument("--output", type=Path, default=Path("steam_next_fest.xlsx"))
    parser.add_argument("--country", default="ro")
    parser.add_argument("--delay", type=float, default=0.25)
    parser.add_argument("--limit", type=int, default=0,
                        help="Limitează verificările noi pentru fiecare sursă")
    parser.add_argument("--export-only", action="store_true")
    args = parser.parse_args()
    if args.delay < 0 or args.limit < 0 or args.output.suffix.lower() != ".xlsx":
        parser.error("Folosește valori nenegative și un fișier de ieșire .xlsx")
    country = args.country.lower()
    appids = load_appids(args.ids)
    with closing(sqlite3.connect(args.db)) as db:
        db.execute("""CREATE TABLE IF NOT EXISTS next_fest (
            appid INTEGER, country TEXT, store_url TEXT, name TEXT, review_count INTEGER,
            reviews_status TEXT NOT NULL DEFAULT 'pending', reviews_checked_at_utc TEXT,
            ai_disclosure_status TEXT NOT NULL DEFAULT 'pending', ai_disclosure TEXT,
            disclosure_checked_at_utc TEXT, PRIMARY KEY (appid, country))""")
        with db:
            db.executemany("INSERT OR IGNORE INTO next_fest (appid, country, store_url) VALUES (?,?,?)",
                           ((appid, country, f"https://store.steampowered.com/app/{appid}/")
                            for appid in appids))
        try:
            if not args.export_only:
                states = {appid: (reviews, disclosure) for appid, reviews, disclosure in db.execute(
                    "SELECT appid, reviews_status, ai_disclosure_status FROM next_fest WHERE country=?",
                    (country,))}
                reviews = [appid for appid in appids if states[appid][0] != "store_filtered"]
                disclosures = [appid for appid in appids if states[appid][1] in ("pending", "unavailable")]
                client = Client(args.delay)
                collect_reviews(db, client, reviews[:args.limit or None], country)
                collect_disclosures(db, client, disclosures[:args.limit or None], country)
        finally:
            export(db, args.output, appids, country)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(130)
    except (RuntimeError, OSError, ValueError, sqlite3.Error) as error:
        sys.exit(f"Eroare: {error}")
