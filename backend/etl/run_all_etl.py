# -*- coding: utf-8 -*-
"""
Run all ETL jobs in order with one command.

Order:
0) ensure_schema.py
1) etl_notes_people.py
2) etl_notes_rawdata.py
3) etl_notes_overtime.py
4) etl_notes_holiday.py
5) aggregate_metrics.py
6) validation summary

Usage:
  python backend/etl/run_all_etl.py
  python backend/etl/run_all_etl.py --db "D:\\Fastapi\\src\\data\\mos_attendance.db" --truncate --replace
  python backend/etl/run_all_etl.py --dept 2531 2537 --include-q --mark-inactive
"""

from __future__ import annotations

import argparse
import datetime as dt
import sqlite3
import subprocess
import sys
import time
from pathlib import Path
from typing import List


def run_step(cmd: List[str], title: str):
    print("\n" + "=" * 70)
    print(f"[RUN] {title}")
    print("=" * 70)
    print("CMD:", " ".join(cmd))
    t0 = time.time()
    p = subprocess.run(cmd)
    sec = time.time() - t0
    if p.returncode != 0:
        raise RuntimeError(f"Step failed: {title} (exit={p.returncode})")
    print(f"[OK] {title} ({sec:.1f}s)")


def query_one(conn: sqlite3.Connection, sql: str, params=()):
    cur = conn.execute(sql, params)
    row = cur.fetchone()
    return row[0] if row else 0


def write_validation_report(db_path: Path, report_path: Path):
    conn = sqlite3.connect(db_path)
    try:
        employees = query_one(conn, "SELECT COUNT(1) FROM employees")
        raw_rows = query_one(conn, "SELECT COUNT(1) FROM raw_records")
        ot_rows = query_one(conn, "SELECT COUNT(1) FROM ot_records")
        holiday_rows = query_one(conn, "SELECT COUNT(1) FROM holiday_balances")
        wk_rows = query_one(conn, "SELECT COUNT(1) FROM weekly_metrics")
        mo_rows = query_one(conn, "SELECT COUNT(1) FROM monthly_metrics")

        bad_emp = query_one(conn, "SELECT COUNT(1) FROM employees WHERE emp_no IS NULL OR TRIM(emp_no)=''")
        bad_raw = query_one(conn, "SELECT COUNT(1) FROM raw_records WHERE emp_no IS NULL OR TRIM(emp_no)=''")
        bad_ot = query_one(conn, "SELECT COUNT(1) FROM ot_records WHERE emp_no IS NULL OR TRIM(emp_no)=''")
        neg_ot = query_one(conn, "SELECT COUNT(1) FROM ot_records WHERE hours < 0")
        huge_ot = query_one(conn, "SELECT COUNT(1) FROM ot_records WHERE hours > 24")

        now = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        lines = [
            "ETL VALIDATION REPORT",
            f"time: {now}",
            f"db: {db_path}",
            "",
            "[Row Counts]",
            f"employees: {employees}",
            f"raw_records: {raw_rows}",
            f"ot_records: {ot_rows}",
            f"holiday_balances: {holiday_rows}",
            f"weekly_metrics: {wk_rows}",
            f"monthly_metrics: {mo_rows}",
            "",
            "[Data Quality]",
            f"employees missing emp_no: {bad_emp}",
            f"raw_records missing emp_no: {bad_raw}",
            f"ot_records missing emp_no: {bad_ot}",
            f"ot_records negative hours: {neg_ot}",
            f"ot_records >24h rows: {huge_ot}",
            "",
            "[Result]",
            "PASS" if (bad_emp + bad_raw + bad_ot + neg_ot == 0) else "CHECK_REQUIRED",
            "",
        ]
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text("\n".join(lines), encoding="utf-8")
    finally:
        conn.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", default="", help="SQLite DB path override")
    ap.add_argument("--dept", nargs="*", default=None, help="Dept override, e.g. --dept 2531 2537")
    ap.add_argument("--truncate", action="store_true", help="truncate rawdata/overtime/holiday before insert")
    ap.add_argument("--replace", action="store_true", help="replace weekly/monthly metrics during aggregate")
    ap.add_argument("--include-q", action="store_true", help="include Q-prefix employees in people ETL")
    ap.add_argument("--mark-inactive", action="store_true", help="mark employees inactive if absent in people ETL")
    ap.add_argument("--no-roster-filter", action="store_true", help="disable roster filter in people ETL")
    args = ap.parse_args()

    repo_root = Path(__file__).resolve().parents[2]
    py = sys.executable

    scripts = {
        "schema": repo_root / "backend/etl/ensure_schema.py",
        "people": repo_root / "backend/etl/etl_notes_people.py",
        "rawdata": repo_root / "backend/etl/etl_notes_rawdata.py",
        "overtime": repo_root / "backend/etl/etl_notes_overtime.py",
        "holiday": repo_root / "backend/etl/etl_notes_holiday.py",
        "agg": repo_root / "backend/etl/aggregate_metrics.py",
    }

    for k, p in scripts.items():
        if not p.exists():
            raise FileNotFoundError(f"Missing script [{k}]: {p}")

    common_db = ["--db", args.db] if args.db else []
    dept_args = ["--dept", *args.dept] if args.dept else []

    run_step([py, str(scripts["schema"]), *common_db], "Ensure Schema")

    cmd_people = [py, str(scripts["people"]), *common_db, *dept_args]
    if args.include_q:
        cmd_people.append("--include-q")
    if args.mark_inactive:
        cmd_people.append("--mark-inactive")
    if args.no_roster_filter:
        cmd_people.append("--no-roster-filter")
    run_step(cmd_people, "ETL People")

    cmd_raw = [py, str(scripts["rawdata"]), *common_db, *dept_args]
    if args.truncate:
        cmd_raw.append("--truncate")
    run_step(cmd_raw, "ETL Rawdata")

    cmd_ot = [py, str(scripts["overtime"]), *common_db, *dept_args]
    if args.truncate:
        cmd_ot.append("--truncate")
    run_step(cmd_ot, "ETL Overtime")

    cmd_holiday = [py, str(scripts["holiday"]), *common_db, *dept_args]
    if args.truncate:
        cmd_holiday.append("--truncate")
    run_step(cmd_holiday, "ETL Holiday")

    cmd_agg = [py, str(scripts["agg"]), *common_db]
    if args.replace:
        cmd_agg.append("--replace")
    run_step(cmd_agg, "Aggregate Metrics")

    db_for_report = Path(args.db) if args.db else Path(r"D:\Fastapi\src\data\mos_attendance.db")
    if not db_for_report.is_absolute():
        db_for_report = (repo_root / db_for_report).resolve()

    report_path = repo_root / "logs" / f"etl_validation_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    write_validation_report(db_for_report, report_path)
    print(f"[OK] validation report: {report_path}")

    print("\n" + "=" * 70)
    print("ALL ETL DONE")
    print("=" * 70)


if __name__ == "__main__":
    main()
