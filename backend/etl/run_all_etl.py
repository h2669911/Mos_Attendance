# -*- coding: utf-8 -*-
"""
Run all ETL jobs in order with one command.

Order:
1) etl_notes_people.py
2) etl_notes_rawdata.py
3) etl_notes_overtime.py
4) etl_notes_holiday.py
5) aggregate_metrics.py

Usage:
  python backend/etl/run_all_etl.py
  python backend/etl/run_all_etl.py --db "D:\\Fastapi\\src\\data\\mos_attendance.db" --truncate --replace
  python backend/etl/run_all_etl.py --dept 2531 2537 --include-q --mark-inactive
"""

from __future__ import annotations

import argparse
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
    dt = time.time() - t0
    if p.returncode != 0:
        raise RuntimeError(f"Step failed: {title} (exit={p.returncode})")
    print(f"[OK] {title} ({dt:.1f}s)")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", default="", help="SQLite DB path override")
    ap.add_argument("--dept", nargs="*", default=None, help="Dept override, e.g. --dept 2531 2537")
    ap.add_argument("--truncate", action="store_true", help="truncate rawdata/overtime/holiday source tables before insert")
    ap.add_argument("--replace", action="store_true", help="replace weekly/monthly metrics during aggregate")
    ap.add_argument("--include-q", action="store_true", help="include Q-prefix employees in people ETL")
    ap.add_argument("--mark-inactive", action="store_true", help="mark employees inactive if not in latest people ETL result")
    ap.add_argument("--no-roster-filter", action="store_true", help="disable roster filter in people ETL")
    args = ap.parse_args()

    repo_root = Path(__file__).resolve().parents[2]
    py = sys.executable

    scripts = {
        "people": repo_root / "backend/etl/etl_notes_people.py",
        "rawdata": repo_root / "backend/etl/etl_notes_rawdata.py",
        "overtime": repo_root / "backend/etl/etl_notes_overtime.py",
        "holiday": repo_root / "backend/etl/etl_notes_holiday.py",
        "agg": repo_root / "backend/etl/aggregate_metrics.py",
    }

    for k, p in scripts.items():
        if not p.exists():
            raise FileNotFoundError(f"Missing script [{k}]: {p}")

    common_db = []
    if args.db:
        common_db = ["--db", args.db]

    dept_args = []
    if args.dept:
        dept_args = ["--dept", *args.dept]

    # 1) people
    cmd_people = [py, str(scripts["people"]), *common_db, *dept_args]
    if args.include_q:
        cmd_people.append("--include-q")
    if args.mark_inactive:
        cmd_people.append("--mark-inactive")
    if args.no_roster_filter:
        cmd_people.append("--no-roster-filter")
    run_step(cmd_people, "ETL People")

    # 2) rawdata
    cmd_raw = [py, str(scripts["rawdata"]), *common_db, *dept_args]
    if args.truncate:
        cmd_raw.append("--truncate")
    run_step(cmd_raw, "ETL Rawdata")

    # 3) overtime
    cmd_ot = [py, str(scripts["overtime"]), *common_db, *dept_args]
    if args.truncate:
        cmd_ot.append("--truncate")
    run_step(cmd_ot, "ETL Overtime")

    # 4) holiday
    cmd_holiday = [py, str(scripts["holiday"]), *common_db, *dept_args]
    if args.truncate:
        cmd_holiday.append("--truncate")
    run_step(cmd_holiday, "ETL Holiday")

    # 5) aggregate
    cmd_agg = [py, str(scripts["agg"]), *common_db]
    if args.replace:
        cmd_agg.append("--replace")
    run_step(cmd_agg, "Aggregate Metrics")

    print("\n" + "=" * 70)
    print("ALL ETL DONE")
    print("=" * 70)


if __name__ == "__main__":
    main()
