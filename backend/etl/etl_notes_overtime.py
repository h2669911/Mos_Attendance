# -*- coding: utf-8 -*-
"""
ETL: Notes overtime -> SQLite.ot_records (+ employees upsert)

Usage:
  python backend/etl/etl_notes_overtime.py
  python backend/etl/etl_notes_overtime.py --dept 2531 2537 --truncate
  python backend/etl/etl_notes_overtime.py --db "D:\\Fastapi\\src\\data\\mos_attendance.db"
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
    s = str(v).strip().replace(",", "")
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


def fetch_ot_rows(depts: Iterable[str], notes_password: str) -> List[Dict[str, Any]]:
    print("[1/3] Fetch overtime from Notes ...")
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
        shift_desc = str(get_item(doc, "SHIFT_DESC") or "").strip()

        start_raw = get_item(doc, "OTStartDate") or get_item(doc, "StartDate")
        end_raw = get_item(doc, "OTEndDate") or get_item(doc, "EndDate")
        apply_item = str(get_item(doc, "ApplyItem") or "").strip()
        ot_remark = str(get_item(doc, "OTRemark") or "").strip()
        desc = str(get_item(doc, "Description") or "").strip()
        sign_raw = str(get_item(doc, "SignStatus") or "").strip()
        hours = to_num(get_item(doc, "ApplyH") or get_item(doc, "LApplyH"))

        rows.append(
            {
                "emp_no": emp_no,
                "emp_name": emp_name,
                "dept_no": dept_no,
                "shift_desc": shift_desc,
                "start_dt": iso_naive(start_raw),
                "end_dt": iso_naive(end_raw),
                "apply_item": apply_item,
                "ot_remark": ot_remark,
                "status": map_status(sign_raw),
                "description": desc,
                "hours": hours,
                "sign_status_raw": sign_raw,
                "foreign_type": "外勞" if ("外勞" in shift_desc) else "本勞",
                "shift_cat": "常日" if ("常日" in shift_desc) else "輪班",
                "raw": {
                    "Form": str(get_item(doc, "Form") or ""),
                    "UNID": str(getattr(doc, "UniversalID", "") or ""),
                },
            }
        )

    print(f"    matched rows: {len(rows)}")
    return rows


def write_db(db_path: Path, rows: List[Dict[str, Any]], truncate: bool):
    print("[2/3] Write ot_records ...")
    db_path.parent.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(db_path)
    try:
        conn.execute("PRAGMA foreign_keys=ON;")
        if truncate:
            conn.execute("DELETE FROM ot_records WHERE source='notes_overtime'")

        for r in rows:
            upsert_employee(conn, r["emp_no"], r["emp_name"], r["dept_no"], r["shift_desc"])
            conn.execute(
                """
                INSERT INTO ot_records(
                    emp_no, emp_name, dept_no, start_dt, end_dt, apply_item, ot_remark,
                    status, description, hours, foreign_type, shift_cat, sign_status_raw, source, raw_json
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    r["emp_no"], r["emp_name"], r["dept_no"], r["start_dt"], r["end_dt"],
                    r["apply_item"], r["ot_remark"], r["status"], r["description"], r["hours"],
                    r["foreign_type"], r["shift_cat"], r["sign_status_raw"], "notes_overtime",
                    json.dumps(r["raw"], ensure_ascii=False),
                ),
            )

        conn.commit()
        print(f"    inserted: {len(rows)}")
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

    print("=== ETL Notes Overtime ===")
    print(f"DB: {db_path}")
    print(f"Depts: {', '.join(depts)}")

    rows = fetch_ot_rows(depts, notes_password)
    write_db(db_path, rows, args.truncate)
    print("[3/3] Done.")


if __name__ == "__main__":
    main()
