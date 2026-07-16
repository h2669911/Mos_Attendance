from __future__ import annotations

import sqlite3
from pathlib import Path
from fastapi import FastAPI

app = FastAPI(title="Mos Attendance API", version="0.1.0")


def get_db_path() -> Path:
    repo_root = Path(__file__).resolve().parents[2]
    return repo_root / "backend/data/mos_attendance.db"


@app.get("/health")
def health():
    db_path = get_db_path()
    ok = db_path.exists()
    return {"status": "ok", "db_exists": ok, "db_path": str(db_path)}


@app.get("/api/metrics/weekly")
def get_weekly_metrics(limit: int = 26):
    db_path = get_db_path()
    if not db_path.exists():
        return {"items": [], "warning": "database not initialized"}

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            """
            SELECT iso_week, tw_ot, foreign_ot, normal_ot,
                   tw_leave, foreign_leave, normal_leave, updated_at
            FROM weekly_metrics
            ORDER BY iso_week DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()
        return {"items": [dict(r) for r in rows]}
    finally:
        conn.close()
