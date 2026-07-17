# -*- coding: utf-8 -*-
"""
Ensure required SQLite schema for ETL pipeline.

Usage:
  python src/scripts/etl/ensure_schema.py
  python src/scripts/etl/ensure_schema.py --db src/data/mos_attendance.db
"""

from __future__ import annotations

import argparse
import os
import sqlite3
from pathlib import Path


def get_db_path(cli_db: str) -> Path:
    repo_root = Path(__file__).resolve().parents[3]

    if cli_db:
        p = Path(cli_db)
        if p.is_absolute():
            return p
        return (repo_root / p).resolve()

    env = os.getenv("MOS_DB_PATH", "").strip()
    if env:
        pe = Path(env)
        return pe if pe.is_absolute() else (repo_root / pe).resolve()

    return (repo_root / "src/data/mos_attendance.db").resolve()


def ensure_employees(conn: sqlite3.Connection):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS employees(
            emp_no TEXT PRIMARY KEY,
            emp_name TEXT,
            dept_no TEXT,
            foreign_type TEXT,
            shift_cat TEXT,
            shift_raw TEXT,
            active INTEGER DEFAULT 1,
            updated_at TEXT
        )
        """
    )


def ensure_raw_records(conn: sqlite3.Connection):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS raw_records(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_no TEXT,
            emp_name TEXT,
            dept_no TEXT,
            work_date TEXT,
            check_in TEXT,
            check_out TEXT,
            overtime_hours REAL DEFAULT 0,
            support_hours REAL DEFAULT 0,
            leave_hours REAL DEFAULT 0,
            source TEXT,
            raw_json TEXT
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_raw_emp_date ON raw_records(emp_no, work_date)")


def ensure_ot_records(conn: sqlite3.Connection):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS ot_records(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_no TEXT,
            emp_name TEXT,
            dept_no TEXT,
            start_dt TEXT,
            end_dt TEXT,
            apply_item TEXT,
            ot_remark TEXT,
            status TEXT,
            description TEXT,
            hours REAL DEFAULT 0,
            foreign_type TEXT,
            shift_cat TEXT,
            sign_status_raw TEXT,
            source TEXT,
            raw_json TEXT
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_ot_emp_start ON ot_records(emp_no, start_dt)")


def ensure_holiday_balances(conn: sqlite3.Connection):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS holiday_balances(
            emp_no TEXT PRIMARY KEY,
            emp_name TEXT,
            dept_no TEXT,
            shift_raw TEXT,
            import_dt TEXT,
            l_total REAL DEFAULT 0,
            l_initial REAL DEFAULT 0,
            l_deferred REAL DEFAULT 0,
            l_used REAL DEFAULT 0,
            l_adjust REAL DEFAULT 0,
            l_remain REAL DEFAULT 0,
            c_total REAL DEFAULT 0,
            c_deferred REAL DEFAULT 0,
            c_adjust REAL DEFAULT 0,
            c_from_ot REAL DEFAULT 0,
            c_used REAL DEFAULT 0,
            c_def_remain REAL DEFAULT 0,
            c_remain REAL DEFAULT 0,
            total_leave REAL DEFAULT 0,
            source TEXT,
            updated_at TEXT
        )
        """
    )


def ensure_metrics(conn: sqlite3.Connection):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS weekly_metrics(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            week_start TEXT,
            dept_no TEXT,
            foreign_type TEXT,
            shift_cat TEXT,
            headcount INTEGER DEFAULT 0,
            avg_ot_hours REAL DEFAULT 0,
            avg_support_hours REAL DEFAULT 0,
            avg_leave_hours REAL DEFAULT 0,
            source TEXT,
            created_at TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS monthly_metrics(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            month_key TEXT,
            dept_no TEXT,
            foreign_type TEXT,
            shift_cat TEXT,
            headcount INTEGER DEFAULT 0,
            avg_ot_hours REAL DEFAULT 0,
            avg_support_hours REAL DEFAULT 0,
            avg_leave_hours REAL DEFAULT 0,
            source TEXT,
            created_at TEXT
        )
        """
    )


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", default="", help="SQLite DB path override")
    args = ap.parse_args()

    db_path = get_db_path(args.db)
    db_path.parent.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(db_path)
    try:
        conn.execute("PRAGMA foreign_keys=ON")
        ensure_employees(conn)
        ensure_raw_records(conn)
        ensure_ot_records(conn)
        ensure_holiday_balances(conn)
        ensure_metrics(conn)
        conn.commit()
    finally:
        conn.close()

    print(f"[OK] schema ensured: {db_path}")


if __name__ == "__main__":
    main()
