#!/usr/bin/env python3
"""Colectează catalogul Steam și numărul de recenzii. Python 3.10+.

Instalează dependența pentru exportul Excel: python -m pip install openpyxl

Rulare (Linux/macOS):
    export STEAM_API_KEY='cheia_ta'
    python steam_catalog.py --limit 20
    python steam_catalog.py
Windows PowerShell:
    $env:STEAM_API_KEY='cheia_ta'
    python steam_catalog.py --limit 20
    python steam_catalog.py

--limit limitează numărul de jocuri procesate în această rulare, după colectarea
catalogului. Aceeași comandă reia progresul din steam_progress.sqlite3.
Fișierul steam_games.xlsx este regenerat din baza de progres la finalul rulării.
Filtrează o copie pentru etapa de AI disclosure. Folosește --export-only pentru
a genera fișierul din progresul existent, fără cheie sau acces la rețea.
--since 2024-01-01 exportă doar jocurile cu data lansării de la acea zi inclusiv.
Filtrul se aplică la export. --limit limitează jocurile verificate, nu potrivirile.
--fast citește până la 100 de jocuri per cerere și folosește numărul de recenzii
din scorul magazinului (reviews_status=store_filtered). Totalul API per joc are
altă definiție. În modul rapid, numerele pozitive și negative rămân goale.
--refresh-catalog reia catalogul de la început, păstrând metadatele deja colectate.
--retry-errors reîncearcă erorile din rulări anterioare.

Sfera colectării: catalogul disponibil în momentul descărcării și paginile
accesibile pentru --country (implicit ro). Un rezultat valid de appdetails
nu garantează posibilitatea cumpărării în fiecare țară. Sunt exportate jocurile
cu coming_soon=false. Data este cea returnată de magazin, inclusiv pentru Early
Access, nu o reconstrucție a istoricului lansărilor. Lipsa unei date precise
lasă release_date gol și păstrează release_date_raw.

Numărul recenziilor este cel raportat de GetAppReviews pentru toate limbile,
toate achizițiile și cu off-topic inclus, la collected_at_utc. Poate diferi de
scorul afișat pe pagina magazinului. Nu reprezintă descărcarea textelor recenziilor.
Metadatele folosesc /api/appdetails, un endpoint public fără contract documentat
stabil. Dacă structura sa se schimbă, scriptul poate necesita adaptare.

Surse:
https://partner.steamgames.com/doc/webapi/IStoreService
https://partner.steamgames.com/doc/webapi/IUserReviewsService
"""
import argparse
import json
import locale
import os
import re
import sqlite3
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timezone
from concurrent.futures import ThreadPoolExecutor
from email.utils import parsedate_to_datetime
from html import unescape
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
except ModuleNotFoundError as error:
    if error.name != "openpyxl":
        raise
    sys.exit("Instalează openpyxl: python -m pip install openpyxl")

FIELDS = ["appid", "name", "release_date", "release_date_raw", "total_reviews",
          "positive_reviews", "negative_reviews", "store_url", "country",
          "collected_at_utc", "reviews_status"]


def now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


class Client:
    def __init__(self, delay):
        self.delay = delay
        self.next_request = 0.0
        self.request_lock = threading.Lock()

    def get(self, url, params, *, page=False):
        request = urllib.request.Request(
            url + "?" + urllib.parse.urlencode(params),
            headers={"User-Agent": "SteamCatalogResearch/1.0"})
        for attempt in range(5):
            with self.request_lock:
                time.sleep(max(0, self.next_request - time.monotonic()))
                self.next_request = time.monotonic() + self.delay
            wait = 2 ** (attempt + 1)
            try:
                with urllib.request.urlopen(request, timeout=40) as response:
                    if page:
                        return response.geturl(), response.read().decode("utf-8", errors="replace")
                    return json.load(response)
            except urllib.error.HTTPError as error:
                if error.code not in (429, 500, 502, 503, 504):
                    # Do not include the request URL, which may contain the API key.
                    raise RuntimeError(f"HTTP {error.code}") from None
                retry_after = error.headers.get("Retry-After", "")
                try:
                    wait = max(wait, float(retry_after))
                except ValueError:
                    try:
                        wait = max(wait, (parsedate_to_datetime(retry_after)
                                         - datetime.now(timezone.utc)).total_seconds())
                    except (ValueError, TypeError, OverflowError):
                        pass
                if error.code == 429:
                    with self.request_lock:
                        self.next_request = max(self.next_request, time.monotonic() + wait)
            except (urllib.error.URLError, TimeoutError, OSError, ValueError):
                pass
            if attempt < 4:
                print(f"Cerere nereușită. Reîncercare peste {wait:.0f}s.", flush=True)
                time.sleep(wait)
        raise RuntimeError("Cererea a eșuat după 5 încercări")


def open_database(path, country):
    db = sqlite3.connect(path)
    db.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)")
    db.execute("""CREATE TABLE IF NOT EXISTS apps (
        appid INTEGER PRIMARY KEY, name TEXT, active INTEGER DEFAULT 1,
        state TEXT DEFAULT 'pending', payload TEXT, error TEXT)""")
    db.execute("""CREATE TABLE IF NOT EXISTS disclosures (
        appid INTEGER PRIMARY KEY, status TEXT NOT NULL, text TEXT)""")
    saved = db.execute("SELECT value FROM settings WHERE key='country'").fetchone()
    if saved and saved[0] != country:
        raise RuntimeError("Folosește o bază --db separată pentru altă țară")
    with db:
        db.execute("INSERT OR IGNORE INTO settings VALUES ('country', ?)", (country,))
    return db


def catalogue(db, client, key, refresh):
    if refresh:
        with db:
            db.execute("DELETE FROM settings WHERE key IN ('cursor','catalog_complete')")
            db.execute("UPDATE apps SET active=0")
    if db.execute("SELECT 1 FROM settings WHERE key='catalog_complete'").fetchone():
        return
    row = db.execute("SELECT value FROM settings WHERE key='cursor'").fetchone()
    cursor = int(row[0]) if row else 0
    while True:
        result = client.get("https://api.steampowered.com/IStoreService/GetAppList/v1/", {
            "key": key, "input_json": json.dumps({
                "include_games": True, "include_dlc": False, "include_software": False,
                "include_videos": False, "include_hardware": False,
                "max_results": 50000, "last_appid": cursor})})
        apps = result["response"].get("apps", [])
        if not apps:
            with db:
                db.execute("INSERT OR REPLACE INTO settings VALUES ('catalog_complete', ?)", (now(),))
            break
        next_cursor = max(int(app["appid"]) for app in apps)
        if next_cursor <= cursor:
            raise RuntimeError("Paginarea catalogului nu a avansat")
        with db:
            db.executemany("""INSERT INTO apps(appid,name,active) VALUES (?,?,1)
                ON CONFLICT(appid) DO UPDATE SET name=excluded.name, active=1""",
                [(int(app["appid"]), app["name"]) for app in apps])
            db.execute("INSERT OR REPLACE INTO settings VALUES ('cursor', ?)", (str(next_cursor),))
        cursor = next_cursor
        print(f"Catalog salvat până la AppID {cursor}.", flush=True)


def parse_date(value):
    for fmt in ("%d %b, %Y", "%b %d, %Y", "%d %B, %Y", "%B %d, %Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt).date().isoformat()
        except ValueError:
            continue
    return ""


def collect_app(client, appid, country):
    result = client.get("https://store.steampowered.com/api/appdetails", {
        "appids": appid, "l": "english", "cc": country})
    item = result[str(appid)]
    if not item.get("success"):
        raise RuntimeError("Metadatele sunt indisponibile în această țară sau temporar")
    data = item["data"]
    if data["type"] != "game":
        return "excluded", None
    release = data["release_date"]
    if release["coming_soon"]:
        return "excluded", None
    row = dict.fromkeys(FIELDS, "")
    row.update(appid=appid, name=data["name"], release_date_raw=release.get("date", ""),
               release_date=parse_date(release.get("date", "")),
               store_url=f"https://store.steampowered.com/app/{appid}/",
               country=country, collected_at_utc=now())
    try:
        reviews = client.get("https://api.steampowered.com/IUserReviewsService/GetAppReviews/v1/", {
            "input_json": json.dumps({"appid": appid, "languages": ["all"],
                "filter": 1, "review_type": 0, "purchase_type": 1,
                "filter_offtopic_activity": False, "num_per_page": 1, "cursor": "*"})})
        summary = reviews["response"]["query_summary"]
        # Protobuf JSON may omit scalar fields whose value is zero.
        if not isinstance(summary, dict):
            raise ValueError("Invalid query_summary")
        row.update(total_reviews=int(summary.get("total_reviews", 0)),
                   positive_reviews=int(summary.get("total_positive", 0)),
                   negative_reviews=int(summary.get("total_negative", 0)),
                   reviews_status="ok", collected_at_utc=now())
        if row["total_reviews"] != row["positive_reviews"] + row["negative_reviews"]:
            raise ValueError("Totalul recenziilor nu corespunde sumei")
        return "ok", row
    except (RuntimeError, KeyError, TypeError, ValueError):
        row.update(total_reviews="", positive_reviews="", negative_reviews="",
                   reviews_status="error")
        return "review_error", row


def collect_batches(db, client, pending, country):
    for start in range(0, len(pending), 100):
        appids = [appid for (appid,) in pending[start:start + 100]]
        result = client.get("https://api.steampowered.com/IStoreBrowseService/GetItems/v1/", {
            "input_json": json.dumps({
                "ids": [{"appid": appid} for appid in appids],
                "context": {"language": "english", "country_code": country.upper()},
                "data_request": {"include_release": True, "include_reviews": True},
            }, separators=(",", ":"))})
        items = {item["id"]: item for item in result["response"]["store_items"]}
        collected = now()
        updates = []
        for appid in appids:
            item = items.get(appid, {})
            release = item.get("release", {})
            if item.get("success") != 1:
                updates.append(("error", None, "Joc indisponibil în răspunsul Steam", appid))
                continue
            if not item.get("visible") or item.get("type") != 0 or release.get("is_coming_soon"):
                updates.append(("excluded", None, "", appid))
                continue
            timestamp = release.get("steam_release_date")
            released = datetime.fromtimestamp(timestamp, timezone.utc).date() if timestamp else None
            if released and released > datetime.now(timezone.utc).date():
                updates.append(("excluded", None, "", appid))
                continue
            count = item.get("reviews", {}).get("summary_filtered", {}).get("review_count")
            row = dict.fromkeys(FIELDS, "")
            row.update(appid=appid, name=item["name"],
                       release_date=released.isoformat() if released else "",
                       release_date_raw=release.get("custom_release_date_message", ""),
                       total_reviews=count if count is not None else "",
                       store_url=f"https://store.steampowered.com/app/{appid}/",
                       country=country, collected_at_utc=collected,
                       reviews_status="store_filtered" if count is not None else "store_missing")
            state = "store_ok" if count is not None else "review_error"
            updates.append((state, json.dumps(row, ensure_ascii=False), "", appid))
        with db:
            db.executemany("UPDATE apps SET state=?,payload=?,error=? WHERE appid=?", updates)
        print(f"[{start + len(appids)}/{len(pending)}] AppID-uri verificate în loturi.", flush=True)


def disclosure_from_page(appid, url, html):
    path = urllib.parse.urlsplit(url).path
    if not re.match(rf"/app/{appid}(?:/|$)", path) or "apphub_AppName" not in html:
        return "error", None
    section = re.search(
        r'<h2\b[^>]*>\s*AI Generated Content Disclosure\s*</h2>(.*?)</div\s*>',
        html, re.IGNORECASE | re.DOTALL)
    if not section:
        if re.search(r'<h2\b[^>]*>\s*AI Generated Content Disclosure\s*</h2>', html, re.I):
            return "error", None
        return "absent", None
    body = re.search(r'<i\b[^>]*>(.*?)</i\s*>', section[1], re.I | re.S)
    if body:
        text = body[1]
    else:
        text = re.sub(r'<p\b[^>]*>\s*The developers describe[^<]*</p>', '', section[1], flags=re.I)
    text = re.sub(r'<br\s*/?>|</p\s*>', '\n', text, flags=re.I)
    text = unescape(re.sub(r'<[^>]+>', '', text))
    return "present", '\n'.join(line.strip() for line in text.splitlines()).strip()


def collect_disclosures(db, client, country, since, retry_errors):
    pending = db.execute("""SELECT apps.appid FROM apps
        LEFT JOIN disclosures ON disclosures.appid=apps.appid
        WHERE apps.active=1 AND apps.payload IS NOT NULL
          AND json_extract(apps.payload, '$.release_date') >= ?
          AND (disclosures.appid IS NULL OR (? AND disclosures.status='error'))
        ORDER BY apps.appid""", (since.isoformat(), retry_errors)).fetchall()
    print(f"Declarații AI: {len(pending)} pagini de verificat.", flush=True)
    def inspect(appid):
        try:
            url, html = client.get(f"https://store.steampowered.com/app/{appid}/",
                                   {"l": "english", "cc": country}, page=True)
            status, text = disclosure_from_page(appid, url, html)
        except RuntimeError as error:
            # urllib raises an HTTP error when a redirect cannot be followed.
            if str(error) not in ("HTTP 301", "HTTP 302", "HTTP 303", "HTTP 307", "HTTP 308",
                                  "HTTP 403", "HTTP 404", "HTTP 451"):
                raise
            print(f"AppID {appid}: pagină inaccesibilă ({error}).", flush=True)
            status, text = "error", None
        return appid, status, text

    workers = ThreadPoolExecutor(max_workers=4)
    try:
        for start in range(0, len(pending), 100):
            batch = pending[start:start + 100]
            for record in workers.map(inspect, (appid for (appid,) in batch)):
                with db:
                    db.execute("INSERT OR REPLACE INTO disclosures VALUES (?, ?, ?)", record)
            print(f"[{start + len(batch)}/{len(pending)}] Declarații AI verificate.", flush=True)
    finally:
        workers.shutdown(wait=False, cancel_futures=True)
    print("Stări declarații AI:", dict(db.execute(
        "SELECT status, COUNT(*) FROM disclosures GROUP BY status")), flush=True)


def export_xlsx(db, output, since=None, ai_only=False):
    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet("Jocuri Steam")
    sheet.freeze_panes = "A2"
    fields = FIELDS + ["ai_disclosure"] if ai_only else FIELDS
    for column, width in enumerate((12, 42, 16, 22, 18, 19, 19, 47, 11, 30, 19, 90)[:len(fields)], 1):
        sheet.column_dimensions[get_column_letter(column)].width = width

    headers = []
    for field in fields:
        cell = WriteOnlyCell(sheet, value=field)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="17476B")
        headers.append(cell)
    sheet.append(headers)

    last_row = 1
    for payload, disclosure in db.execute("""SELECT apps.payload, disclosures.text FROM apps
            LEFT JOIN disclosures ON disclosures.appid=apps.appid
            WHERE apps.active=1 AND apps.payload IS NOT NULL
              AND (?=0 OR disclosures.status='present') ORDER BY apps.appid""", (ai_only,)):
        row = json.loads(payload)
        row["ai_disclosure"] = disclosure
        release_date = row.get("release_date")
        if since and (not release_date or date.fromisoformat(release_date) < since):
            continue
        cells = []
        for field in fields:
            value = row.get(field)
            if field == "release_date" and value:
                cell = WriteOnlyCell(sheet, value=date.fromisoformat(value))
                cell.number_format = "yyyy-mm-dd"
            elif isinstance(value, str) and value:
                cell = WriteOnlyCell(sheet, value=value)
                cell.data_type = "s"
            else:
                cell = None if value == "" else value
            cells.append(cell)
        sheet.append(cells)
        last_row += 1
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(fields))}{last_row}"

    temp = output.with_suffix(".tmp.xlsx")
    try:
        workbook.save(temp)
        os.replace(temp, output)
    finally:
        temp.unlink(missing_ok=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--db", type=Path, default=Path("steam_progress.sqlite3"))
    parser.add_argument("--output", type=Path, default=Path("steam_games.xlsx"))
    parser.add_argument("--country", default="ro")
    parser.add_argument("--delay", type=float, default=2.0)
    parser.add_argument("--limit", type=int, default=0, help="0 = fără limită")
    parser.add_argument("--fast", action="store_true", help="Metadate și recenzii din magazin în loturi")
    parser.add_argument("--ai-only", action="store_true", help="Verifică și exportă doar jocurile cu declarație AI")
    parser.add_argument("--since", type=date.fromisoformat, metavar="YYYY-MM-DD",
                        help="Data minimă de lansare pentru export (inclusiv)")
    parser.add_argument("--retry-errors", action="store_true")
    parser.add_argument("--refresh-catalog", action="store_true")
    parser.add_argument("--export-only", action="store_true")
    args = parser.parse_args()
    if args.delay < 0 or args.limit < 0:
        parser.error("--delay și --limit trebuie să fie >= 0")
    if args.output.suffix.lower() != ".xlsx":
        parser.error("--output trebuie să aibă extensia .xlsx")
    if args.ai_only and (not args.fast or args.since is None):
        parser.error("--ai-only necesită --fast și --since")
    locale.setlocale(locale.LC_TIME, "C")
    db = open_database(args.db, args.country.lower())
    try:
        if args.export_only:
            return
        key = os.environ.get("STEAM_API_KEY")
        catalog_ready = db.execute("SELECT 1 FROM settings WHERE key='catalog_complete'").fetchone()
        if not key and (args.refresh_catalog or not catalog_ready):
            raise RuntimeError("Setează variabila de mediu STEAM_API_KEY înainte de rulare")
        client = Client(args.delay)
        catalogue(db, client, key, args.refresh_catalog)
        if args.retry_errors:
            with db:
                db.execute("UPDATE apps SET state='pending' WHERE state IN ('error','review_error')")
        pending = db.execute("""SELECT appid FROM apps WHERE active=1
            AND (state='pending' OR (state='store_ok' AND ?=0)) ORDER BY appid""",
            (args.fast,)).fetchall()
        if args.limit:
            pending = pending[:args.limit]
        if args.fast:
            collect_batches(db, client, pending, args.country.lower())
            if args.ai_only:
                collect_disclosures(db, client, args.country.lower(), args.since, args.retry_errors)
            return
        for index, (appid,) in enumerate(pending, 1):
            try:
                state, row = collect_app(client, appid, args.country.lower())
                error = "" if state != "review_error" else "Nu s-au putut obține totalurile recenziilor"
                with db:
                    db.execute("UPDATE apps SET state=?,payload=?,error=? WHERE appid=?",
                               (state, json.dumps(row, ensure_ascii=False) if row else None, error, appid))
            except (RuntimeError, KeyError, TypeError, ValueError) as error:
                state = "error"
                with db:
                    db.execute("UPDATE apps SET state=?,error=? WHERE appid=?", (state, str(error), appid))
            print(f"[{index}/{len(pending)}] AppID {appid}: {state}", flush=True)
    finally:
        try:
            export_xlsx(db, args.output, args.since, args.ai_only)
            print("Stare:", dict(db.execute("SELECT state,COUNT(*) FROM apps WHERE active=1 GROUP BY state")))
            print(f"Fișier Excel salvat în {args.output.resolve()}")
        finally:
            db.close()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("Oprit. Progresul este salvat. Reia aceeași comandă.", file=sys.stderr)
        sys.exit(130)
    except (RuntimeError, OSError, ValueError, KeyError, sqlite3.Error) as error:
        print(f"Eroare: {error}", file=sys.stderr)
        sys.exit(1)
