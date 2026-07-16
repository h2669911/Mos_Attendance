#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Initialize SQLite database for Mos Attendance.

Usage:
  python scripts/init_db.py
  python scripts/init_db.py --db backend/data/mos_attendance.db
"""

from __future__ import annotations
import argparse
import sqlite3
from pathlib import Path


def run_sql_file(conn: sqlite3.Connection, path: Path) -> None:
    sql = path.read_text(encoding="utf-8")
    conn.executescript(sql)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--db",
        default="backend/data/mos_attendance.db",
        help="SQLite DB file path"
    )
    args = parser.parse_args()

    repo_root = Path(__file__).resolve().parents[1]
    db_path = repo_root / args.db
    db_path.parent.mkdir(parents=True, exist_ok=True)

    schema_path = repo_root / "backend/sql/schema.sql"
    indexes_path = repo_root / "backend/sql/indexes.sql"

    conn = sqlite3.connect(db_path)
    try:
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA synchronous=NORMAL;")
        conn.execute("PRAGMA foreign_keys=ON;")

        run_sql_file(conn, schema_path)
        run_sql_file(conn, indexes_path)

        conn.commit()
        print(f"✅ DB initialized: {db_path}")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
