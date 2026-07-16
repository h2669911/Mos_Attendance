# -*- coding: utf-8 -*-
"""
ETL: Notes leave data -> SQLite.leave_records (+ employees upsert)

Usage:
  python backend/etl/etl_notes_rawdata.py
  python backend/etl/etl_notes_rawdata.py --dept 2531 2537 --truncate
  python backend/etl/etl_notes_rawdata.py --db "D:\\Fastapi\\src\\data\\mos_attendance.db"
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import sqlite3
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import win32com.client  # pywin32


DEFAULT_DEPTS = ["2531", "2537"]
CFG_FILENAME = "部門設定_請先改這裡.ps1"


def get_db_path(cli_db: str) -> Path:
    if cli_db:
        p = Path(cli_db)
        if p.is_absolute():
            return p
        repo_root = Path(__file__).resolve().parents[2]
        return repo_root / p

    env = os.getenv("MOS_DB_PATH", "").strip()
    if env:
        return Path(env)

    return Path(r"D:\Fastapi\src\data\mos_attendance.db")


def parse_settings() -> Tuple[List[str], str]:
    depts = DEFAULT_DEPTS[:]
    notes_password = ""

    cur = Path(__file__).resolve()
    candidates = [
        cur.parent / CFG_FILENAME,
        cur.parents[1] / CFG_FILENAME,
        cur.parents[2] / CFG_FILENAME,
    ]

    found = None
    for p in candidates:
        if p.exists():
            found = p
            break

    if not found:
        return depts, notes_password

    text = found.read_text(encoding="utf-8", errors="ignore")

    m = re.search(r"\$Depts\s*=\s*@\((.*?)\)", text, re.S)
    if m:
        vals = re.findall(r'"([^"]+)"', m.group(1))
        if vals:
            depts = vals

    m2 = re.search(r'\$NOTES_PASSWORD\s*=\s*"([^"]*)"', text)
    if m2:
        notes_password = m2.group(1)

    return depts, notes_password


def get_item(doc, name: str):
    try:
        vals = doc.GetItemValue(name)
        if vals and len(vals) > 0:
            return vals[0]
    except Exception:
        pass
    return None


def to_datetime(v: Any) -> Optional[dt.datetime]:
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v
    s = str(v).strip()
    if not s:
        return None
    for fmt in ("%Y/%m/%d %H:%M:%S", "%Y/%m/%d", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y"):
        try:
            return dt.datetime.strptime(s, fmt)
        except Exception:
            pass
    try:
        return dt.datetime.fromisoformat(s)
    except Exception:
        return None


def iso_naive(v: Any) -> Optional[str]:
    d = to_datetime(v)
    if not d:
        return None
    if d.tzinfo is not None:
        d = d.replace(tzinfo=None)
    return d.strftime("%Y-%m-%d %H:%M:%S")


def to_num(v: Any) -> float:
    if v is None:
        return 0.0
    s = str(v).strip()
    if not s:
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


def map_status(sign_status_raw: str) -> str:
    s = (sign_status_raw or "").strip()
    if s in ("Completed", "Approved", "Processing"):
        return s
    if "核准" in s or "完成" in s:
        return "Approved"
    if "處理" in s or "簽核中" in s:
        return "Processing"
    return s or "Unknown"


def upsert_employee(conn: sqlite3.Connection, emp_no: str, emp_name: str, dept_no: str, shift_desc: str):
    if not emp_no:
        return

    foreign_type = "外勞" if ("外勞" in (shift_desc or "")) else "本勞"
    shift_cat = "常日" if ("常日" in (shift_desc or "")) else "輪班"

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
        (emp_no, emp_name, dept_no, foreign_type, shift_cat, shift_desc),
    )


def insert_leave_record(conn: sqlite3.Connection, r: Dict[str, Any]):
    conn.execute(
        """
        INSERT INTO leave_records(
            emp_no, emp_name, dept_no, date_composed, start_dt, end_dt,
            leave_type, status, description, hours, sign_status_raw, source, raw_json
        ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
        """,
        (
            r["emp_no"], r["emp_name"], r["dept_no"], r["date_composed"],
            r["start_dt"], r["end_dt"], r["leave_type"], r["status"],
            r["description"], r["hours"], r["sign_status_raw"], "notes_rawdata",
            json.dumps(r["raw"], ensure_ascii=False),
        ),
    )


def fetch_notes_rows(depts: Iterable[str], notes_password: str) -> List[Dict[str, Any]]:
    print("[1/3] Connect Notes and fetch leave data ...")
    sess = win32com.client.Dispatch("Notes.NotesSession")
    if hasattr(sess, "Initialize"):
        try:
            sess.Initialize(notes_password or "")
        except Exception:
            pass

    db = sess.GetDatabase("CMOS07/CHIPMOS", "ap\\office\\pds0001.nsf")
    if not db.IsOpen:
        db.Open("CMOS07/CHIPMOS", "ap\\office\\pds0001.nsf")

    dt1900 = sess.CreateDateTime("01/01/1900 00:00:00")
    formula = 'SELECT @All & Form != ""'
    dc = db.Search(formula, dt1900, 0)
    if dc is None:
        return []

    result: List[Dict[str, Any]] = []
    total = dc.Count
    depts_set = set(str(x) for x in depts)

    for i in range(1, total + 1):
        doc = dc.GetNthDocument(i)
        if doc is None:
            continue

        dept_no = str(get_item(doc, "DeptNo") or "").strip()
        if dept_no not in depts_set:
            continue

        emp_no = str(get_item(doc, "EmpNo") or "").strip()
        emp_name = str(get_item(doc, "EmpName") or "").strip()
        shift_desc = str(get_item(doc, "SHIFT_DESC") or "").strip()

        if not emp_no:
            continue

        start_raw = get_item(doc, "StartDate")
        end_raw = get_item(doc, "EndDate")
        composed_raw = get_item(doc, "DateComposed")
        leave_type = str(get_item(doc, "LVName") or get_item(doc, "LeaveType") or "").strip()
        desc = str(get_item(doc, "Description") or "").strip()
        sign_raw = str(get_item(doc, "SignStatus") or "").strip()
        hours = to_num(get_item(doc, "LApplyH"))

        row = {
            "emp_no": emp_no,
            "emp_name": emp_name,
            "dept_no": dept_no,
            "shift_desc": shift_desc,
            "date_composed": iso_naive(composed_raw),
            "start_dt": iso_naive(start_raw),
            "end_dt": iso_naive(end_raw),
            "leave_type": leave_type,
            "status": map_status(sign_raw),
            "description": desc,
            "hours": hours,
            "sign_status_raw": sign_raw,
            "raw": {
                "Form": str(get_item(doc, "Form") or ""),
                "UNID": str(getattr(doc, "UniversalID", "") or ""),
            },
        }
        result.append(row)

        if i % 1000 == 0:
            print(f"    scanned: {i}/{total} ... matched: {len(result)}")

    print(f"    done. matched rows: {len(result)}")
    return result


def load_to_db(db_path: Path, rows: List[Dict[str, Any]], truncate: bool):
    print("[2/3] Write to SQLite ...")
    db_path.parent.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(db_path)
    try:
        conn.execute("PRAGMA foreign_keys=ON;")
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA synchronous=NORMAL;")

        if truncate:
            conn.execute("DELETE FROM leave_records WHERE source='notes_rawdata'")

        inserted = 0
        for r in rows:
            upsert_employee(conn, r["emp_no"], r["emp_name"], r["dept_no"], r["shift_desc"])
            insert_leave_record(conn, r)
            inserted += 1

        conn.commit()
        print(f"    inserted: {inserted}")
    finally:
        conn.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", default="", help="SQLite DB path (optional)")
    ap.add_argument("--dept", nargs="*", default=None, help="Dept list override, e.g. --dept 2531 2537")
    ap.add_argument("--truncate", action="store_true", help="Delete old notes_rawdata rows before insert")
    args = ap.parse_args()

    cfg_depts, notes_password = parse_settings()
    depts = args.dept if args.dept else cfg_depts
    db_path = get_db_path(args.db)

    print("===================================================")
    print(" ETL Notes Rawdata -> SQLite")
    print("===================================================")
    print(f"DB     : {db_path}")
    print(f"Depts  : {', '.join(depts)}")
    print(f"Truncate: {args.truncate}")

    rows = fetch_notes_rows(depts, notes_password)
    load_to_db(db_path, rows, args.truncate)

    print("[3/3] Done.")


if __name__ == "__main__":
    main()
