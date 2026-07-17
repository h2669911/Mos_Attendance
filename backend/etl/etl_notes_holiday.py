# -*- coding: utf-8 -*-
"""
ETL: Notes holiday balance -> SQLite.holiday_balances (+ employees upsert)

Usage:
  python backend/etl/etl_notes_holiday.py
  python backend/etl/etl_notes_holiday.py --dept 2531 2537 --truncate
  python backend/etl/etl_notes_holiday.py --db "D:\\Fastapi\\src\\data\\mos_attendance.db"
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import re
import sqlite3
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple

import win32com.client


DEFAULT_DEPTS = ["2531", "2537"]
CFG_FILENAME = "部門設定_請先改這裡.ps1"


def get_db_path(cli_db: str) -> Path:
    if cli_db:
        p = Path(cli_db)
        if p.is_absolute():
            return p
        return Path(__file__).resolve().parents[2] / p

    env = os.getenv("MOS_DB_PATH", "").strip()
    if env:
        return Path(env)

    return Path(r"D:\Fastapi\src\data\mos_attendance.db")


def parse_settings() -> Tuple[List[str], str]:
    depts = DEFAULT_DEPTS[:]
    notes_password = ""

    cur = Path(__file__).resolve()
    for p in [cur.parent / CFG_FILENAME, cur.parents[1] / CFG_FILENAME, cur.parents[2] / CFG_FILENAME]:
        if p.exists():
            text = p.read_text(encoding="utf-8", errors="ignore")
            m = re.search(r"\$Depts\s*=\s*@\((.*?)\)", text, re.S)
            if m:
                vals = re.findall(r'"([^"]+)"', m.group(1))
                if vals:
                    depts = vals
            m2 = re.search(r'\$NOTES_PASSWORD\s*=\s*"([^"]*)"', text)
            if m2:
                notes_password = m2.group(1)
            break
    return depts, notes_password


def get_item(doc, name: str):
    try:
        vals = doc.GetItemValue(name)
        if vals and len(vals) > 0:
            return vals[0]
    except Exception:
        pass
    return None


def to_float(v: Any) -> float:
    if v is None:
        return 0.0
    s = str(v).strip().replace(",", "")
    if not s:
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


def iso_now() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def upsert_employee(conn: sqlite3.Connection, emp_no: str, emp_name: str, dept_no: str, shift_raw: str):
    if not emp_no:
        return
    foreign_type = "外勞" if "外勞" in (shift_raw or "") else "本勞"
    shift_cat = "常日" if "常日" in (shift_raw or "") else "輪班"
    conn.execute(
        """
        INSERT INTO employees(emp_no, emp_name, dept_no, foreign_type, shift_cat, shift_raw, active, updated_at)
        VALUES(?, ?, ?, ?, ?, ?, 1, datetime('now'))
        ON CONFLICT(emp_no) DO UPDATE SET
            emp_name=excluded.emp_name,
            dept_no=excluded.dept_no,
            foreign_type=excluded.foreign_type,
            shift_cat=excluded.shift_cat,
            shift_raw=excluded.shift_raw,
            active=1,
            updated_at=datetime('now')
        """,
        (emp_no, emp_name, dept_no, foreign_type, shift_cat, shift_raw),
    )


def fetch_holiday_rows(depts: Iterable[str], notes_password: str) -> List[Dict[str, Any]]:
    print("[1/3] Fetch holiday balances from Notes ...")
    sess = win32com.client.Dispatch("Notes.NotesSession")
    try:
        sess.Initialize(notes_password or "")
    except Exception:
        pass

    db = sess.GetDatabase("CMOS07/CHIPMOS", "ap\\office\\pds0001.nsf")
    if not db.IsOpen:
        db.Open("CMOS07/CHIPMOS", "ap\\office\\pds0001.nsf")

    dt1900 = sess.CreateDateTime("01/01/1900 00:00:00")
    dc = db.Search('SELECT @All & Form != ""', dt1900, 0)
    if dc is None:
        return []

    depts_set = set(str(x) for x in depts)
    rows: List[Dict[str, Any]] = []

    for i in range(1, dc.Count + 1):
        doc = dc.GetNthDocument(i)
        if doc is None:
            continue

        dept_no = str(get_item(doc, "DeptNo") or "").strip()
        if dept_no not in depts_set:
            continue

        emp_no = str(get_item(doc, "EmpNo") or "").strip()
        if not emp_no:
            continue

        emp_name = str(get_item(doc, "EmpName") or "").strip()
        shift_raw = str(get_item(doc, "SHIFT_DESC") or "").strip()

        l_initial = to_float(get_item(doc, "LInitial"))
        l_deferred = to_float(get_item(doc, "LDefer"))
        l_used = to_float(get_item(doc, "LUsed"))
        l_adjust = to_float(get_item(doc, "LAdjust"))
        l_total = to_float(get_item(doc, "LTotal"))
        l_remain = to_float(get_item(doc, "LRemain"))

        c_total = to_float(get_item(doc, "CTotal"))
        c_deferred = to_float(get_item(doc, "CDefer"))
        c_adjust = to_float(get_item(doc, "CAdjust"))
        c_from_ot = to_float(get_item(doc, "CFromOT"))
        c_used = to_float(get_item(doc, "CUsed"))
        c_def_remain = to_float(get_item(doc, "CDefRemain"))
        c_remain = to_float(get_item(doc, "CRemain"))

        total_leave = l_remain + c_remain

        rows.append(
            {
                "emp_no": emp_no,
                "emp_name": emp_name,
                "dept_no": dept_no,
                "shift_raw": shift_raw,
                "import_dt": iso_now(),
                "l_total": l_total,
                "l_initial": l_initial,
                "l_deferred": l_deferred,
                "l_used": l_used,
                "l_adjust": l_adjust,
                "l_remain": l_remain,
                "c_total": c_total,
                "c_deferred": c_deferred,
                "c_adjust": c_adjust,
                "c_from_ot": c_from_ot,
                "c_used": c_used,
                "c_def_remain": c_def_remain,
                "c_remain": c_remain,
                "total_leave": total_leave,
            }
        )

    print(f"    matched rows: {len(rows)}")
    return rows


def write_db(db_path: Path, rows: List[Dict[str, Any]], truncate: bool):
    print("[2/3] Write holiday_balances ...")
    db_path.parent.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(db_path)
    try:
        conn.execute("PRAGMA foreign_keys=ON;")
        if truncate:
            conn.execute("DELETE FROM holiday_balances")

        for r in rows:
            upsert_employee(conn, r["emp_no"], r["emp_name"], r["dept_no"], r["shift_raw"])
            conn.execute(
                """
                INSERT INTO holiday_balances(
                    emp_no, emp_name, dept_no, shift_raw, import_dt,
                    l_total, l_initial, l_deferred, l_used, l_adjust, l_remain,
                    c_total, c_deferred, c_adjust, c_from_ot, c_used, c_def_remain, c_remain,
                    total_leave, source, updated_at
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,datetime('now'))
                ON CONFLICT(emp_no) DO UPDATE SET
                    emp_name=excluded.emp_name,
                    dept_no=excluded.dept_no,
                    shift_raw=excluded.shift_raw,
                    import_dt=excluded.import_dt,
                    l_total=excluded.l_total,
                    l_initial=excluded.l_initial,
                    l_deferred=excluded.l_deferred,
                    l_used=excluded.l_used,
                    l_adjust=excluded.l_adjust,
                    l_remain=excluded.l_remain,
                    c_total=excluded.c_total,
                    c_deferred=excluded.c_deferred,
                    c_adjust=excluded.c_adjust,
                    c_from_ot=excluded.c_from_ot,
                    c_used=excluded.c_used,
                    c_def_remain=excluded.c_def_remain,
                    c_remain=excluded.c_remain,
                    total_leave=excluded.total_leave,
                    source='notes_holiday',
                    updated_at=datetime('now')
                """,
                (
                    r["emp_no"], r["emp_name"], r["dept_no"], r["shift_raw"], r["import_dt"],
                    r["l_total"], r["l_initial"], r["l_deferred"], r["l_used"], r["l_adjust"], r["l_remain"],
                    r["c_total"], r["c_deferred"], r["c_adjust"], r["c_from_ot"], r["c_used"], r["c_def_remain"], r["c_remain"],
                    r["total_leave"], "notes_holiday",
                ),
            )

        conn.commit()
        print(f"    upserted: {len(rows)}")
    finally:
        conn.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", default="", help="SQLite path override")
    ap.add_argument("--dept", nargs="*", default=None)
    ap.add_argument("--truncate", action="store_true")
    args = ap.parse_args()

    cfg_depts, notes_password = parse_settings()
    depts = args.dept if args.dept else cfg_depts
    db_path = get_db_path(args.db)

    print("=== ETL Notes Holiday ===")
    print(f"DB: {db_path}")
    print(f"Depts: {', '.join(depts)}")

    rows = fetch_holiday_rows(depts, notes_password)
    write_db(db_path, rows, args.truncate)
    print("[3/3] Done.")


if __name__ == "__main__":
    main()
