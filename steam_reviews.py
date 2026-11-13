#!/usr/bin/env python3
"""Colectează recenziile publice și genurile Steam pentru lista filtrată."""
import argparse
import csv
import json
import os
import sqlite3
import sys
from contextlib import closing
from pathlib import Path

from steam_catalog import Client

ROSTER = Path(__file__).resolve().parent / "data/steam-next-fest-filtered-app-ids.csv"
REVIEW_URL = "https://api.steampowered.com/IUserReviewsService/GetAppReviews/v1/"
DETAILS_URL = "https://store.steampowered.com/api/appdetails"


def load_ids(path):
    with path.open(encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        if "app_id" not in (reader.fieldnames or ()):
            raise ValueError("Lista trebuie să aibă coloana app_id")
        appids = list(dict.fromkeys(int(row["app_id"]) for row in reader))
    if not appids or min(appids) <= 0:
        raise ValueError("Lista trebuie să conțină AppID-uri pozitive")
    return appids


def collect_genres(db, client, appid, country):
    response = client.get(DETAILS_URL, {"appids": appid, "l": "english", "cc": country})
    app = response[str(appid)]
    if not app.get("success"):
        raise ValueError("Steam nu a returnat pagina jocului")
    details = app["data"]
    genres = " | ".join(genre["description"] for genre in details.get("genres", []))
    with db:
        db.execute("UPDATE games SET name=?, genres=?, genre_status='ok' WHERE appid=?",
                   (details["name"], genres, appid))


def collect_reviews(db, client, appid, cursor, expected):
    seen = set()
    while True:
        params = {"appid": appid, "filter": 1, "languages": ["all"], "review_type": 0,
                  "purchase_type": 1, "filter_offtopic_activity": False,
                  "num_per_page": 100, "cursor": cursor}
        response = client.get(REVIEW_URL, {"input_json": json.dumps(params)})["response"]
        if cursor == "*":
            expected = response["total_matching"]
            if type(expected) is not int or expected < 0:
                raise ValueError("Număr de recenzii invalid")
        reviews = response.get("reviews", [])
        if not isinstance(reviews, list) or response.get("query_summary", {}).get("num_reviews", len(reviews)) != len(reviews):
            raise ValueError("Pagină de recenzii incompletă")
        if not reviews:
            with db:
                count = db.execute("SELECT COUNT(*) FROM reviews WHERE appid=?", (appid,)).fetchone()[0]
                final_count = response["total_matching"]
                if type(final_count) is not int or final_count < 0:
                    raise ValueError("Număr final de recenzii invalid")
                status = "complete" if count == final_count else "count_mismatch"
                db.execute("""UPDATE games SET review_status=?, expected_reviews=?,
                              final_reported_reviews=? WHERE appid=?""",
                           (status, expected, final_count, appid))
            return status

        next_cursor = response["cursor"]
        if not next_cursor or next_cursor in ("*", cursor) or next_cursor in seen:
            raise ValueError("Cursorul recenziilor nu a avansat")
        records = []
        for review in reviews:
            voted_up = review.get("voted_up", False)  # Protobuf may omit false.
            if type(voted_up) is not bool:
                raise ValueError("Votul recenziei nu este boolean")
            review_id = str(review["recommendationid"])
            if not review_id.isdecimal() or int(review_id) <= 0:
                raise ValueError("ID de recenzie invalid")
            records.append((appid, review_id, "positive" if voted_up else "negative",
                            review.get("language"), review["timestamp_created"], review.get("review", "")))
        with db:
            db.executemany("""INSERT INTO reviews VALUES (?,?,?,?,?,?)
                           ON CONFLICT(appid, review_id) DO UPDATE SET
                           valence=excluded.valence, language=excluded.language,
                           created_at_unix=excluded.created_at_unix, review_text=excluded.review_text""",
                           records)
            db.execute("UPDATE games SET cursor=?, expected_reviews=? WHERE appid=?",
                       (next_cursor, expected, appid))
        seen.add(cursor)
        cursor = next_cursor


def write_csv(path, query):
    temp = path.with_suffix(".tmp")
    try:
        with temp.open("w", encoding="utf-8-sig", newline="") as file:
            writer = csv.writer(file)
            writer.writerow([column[0] for column in query.description])
            for row in query:
                writer.writerow([("'" + value) if isinstance(value, str) and
                                 value.lstrip().startswith(("=", "+", "-", "@")) else value
                                 for value in row])
        os.replace(temp, path)
    finally:
        temp.unlink(missing_ok=True)


def export(db, output):
    write_csv(output / "reviews.csv", db.execute("""
        SELECT r.appid, g.name, g.genres, r.review_id, r.valence, r.language,
               r.created_at_unix, r.review_text
        FROM reviews r JOIN games g USING(appid) ORDER BY r.appid, r.created_at_unix, r.review_id
        """))
    write_csv(output / "games.csv", db.execute("""
        SELECT g.appid, g.name, g.genres, g.genre_status, g.review_status,
               g.expected_reviews, g.final_reported_reviews, COUNT(r.review_id) AS collected_reviews
        FROM games g LEFT JOIN reviews r USING(appid) GROUP BY g.appid ORDER BY g.appid
        """))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--ids", type=Path, default=ROSTER)
    parser.add_argument("--output-dir", type=Path, default=Path("review_output"))
    parser.add_argument("--country", default="ro")
    parser.add_argument("--delay", type=float, default=0.5)
    args = parser.parse_args()
    if args.delay < 0:
        parser.error("--delay trebuie să fie >= 0")
    appids = load_ids(args.ids)
    country = args.country.lower()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    with closing(sqlite3.connect(args.output_dir / "reviews.sqlite3")) as db:
        db.executescript("""
            CREATE TABLE IF NOT EXISTS settings (value TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS games (
                appid INTEGER PRIMARY KEY, name TEXT, genres TEXT,
                genre_status TEXT NOT NULL DEFAULT 'pending',
                review_status TEXT NOT NULL DEFAULT 'pending',
                cursor TEXT NOT NULL DEFAULT '*', expected_reviews INTEGER,
                final_reported_reviews INTEGER);
            CREATE TABLE IF NOT EXISTS reviews (
                appid INTEGER NOT NULL, review_id TEXT NOT NULL,
                valence TEXT NOT NULL, language TEXT, created_at_unix INTEGER NOT NULL,
                review_text TEXT NOT NULL, PRIMARY KEY(appid, review_id));
            """)
        if "final_reported_reviews" not in {column[1] for column in db.execute("PRAGMA table_info(games)")}:
            with db:
                db.execute("ALTER TABLE games ADD COLUMN final_reported_reviews INTEGER")
        config = json.dumps({"appids": sorted(appids), "country": country})
        saved = db.execute("SELECT value FROM settings").fetchone()
        if saved and saved[0] != config:
            raise ValueError("Folosește alt --output-dir pentru o altă listă sau regiune")
        with db:
            if not saved:
                db.execute("INSERT INTO settings VALUES (?)", (config,))
            db.executemany("INSERT OR IGNORE INTO games(appid) VALUES (?)", ((i,) for i in appids))
        client = Client(args.delay)
        errors = (RuntimeError, KeyError, TypeError, ValueError)
        try:
            for index, appid in enumerate(appids, 1):
                genre_status, review_status, cursor, expected = db.execute(
                    "SELECT genre_status, review_status, cursor, expected_reviews FROM games WHERE appid=?",
                    (appid,)).fetchone()
                if genre_status != "ok":
                    try:
                        collect_genres(db, client, appid, country)
                    except errors as error:
                        with db:
                            db.execute("UPDATE games SET genre_status='unavailable' WHERE appid=?", (appid,))
                        print(f"AppID {appid} · genuri: {error}", flush=True)
                if review_status != "complete":
                    if review_status == "count_mismatch":
                        try:
                            review_status = collect_reviews(db, client, appid, cursor, expected)
                        except errors as error:
                            print(f"AppID {appid} · verificare finală: {error}", flush=True)
                    if review_status == "count_mismatch":
                        with db:
                            db.execute("DELETE FROM reviews WHERE appid=?", (appid,))
                            db.execute("""UPDATE games SET cursor='*', expected_reviews=NULL,
                                          final_reported_reviews=NULL, review_status='pending' WHERE appid=?""", (appid,))
                        cursor, expected = "*", None
                    if review_status != "complete":
                        try:
                            review_status = collect_reviews(db, client, appid, cursor, expected)
                        except errors as error:
                            review_status = "error"
                            with db:
                                db.execute("UPDATE games SET review_status='error' WHERE appid=?", (appid,))
                            print(f"AppID {appid} · recenzii: {error}", flush=True)
                print(f"[{index}/{len(appids)}] AppID {appid}: {review_status}", flush=True)
        finally:
            export(db, args.output_dir)
        incomplete = db.execute("""SELECT COUNT(*) FROM games
                                  WHERE review_status!='complete' OR genre_status!='ok'""").fetchone()[0]
        return int(incomplete > 0)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit(130)
    except (OSError, ValueError, sqlite3.Error) as error:
        sys.exit(f"Eroare: {error}")
