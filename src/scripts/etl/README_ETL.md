# ETL Operations Guide

This folder contains the ETL pipeline for Notes → SQLite → metrics.

## Scripts

- `ensure_schema.py`: create required SQLite tables/indexes if missing
- `etl_notes_people.py`: employee master sync
- `etl_notes_rawdata.py`: attendance/raw records import
- `etl_notes_overtime.py`: overtime records import
- `etl_notes_holiday.py`: holiday balance import
- `aggregate_metrics.py`: compute weekly/monthly KPI tables
- `run_all_etl.py`: run all steps in order
- `run_all_etl.bat`: one-click Windows launcher

---

## Quick Start

```bash
python src/scripts/etl/run_all_etl.py --db src/data/mos_attendance.db --truncate --replace --no-roster-filter
```

Windows one-click:

- Double click `src/scripts/etl/run_all_etl.bat`

---

## Optional Arguments

### `run_all_etl.py`

- `--db "<path>"` override SQLite path (absolute or relative to repo root)
- `--dept 2531 2537` override departments
- `--truncate` clear source tables (`raw_records`, `ot_records`, `holiday_balances`) before load
- `--replace` replace aggregate metrics output
- `--include-q` include Q-prefix employee IDs in people ETL
- `--mark-inactive` mark employees inactive if absent from current people ETL result
- `--no-roster-filter` disable roster filtering in people ETL

Example:

```bash
python src/scripts/etl/run_all_etl.py --db src/data/mos_attendance.db --dept 2531 2537 --truncate --replace
```

---

## Output / Logs

After `run_all_etl.py`, a validation report is generated under:

- `logs/etl_validation_YYYYMMDD_HHMMSS.txt`

It includes:
- row counts
- key null checks (`emp_no`)
- overtime anomaly checks (negative or >24h)

---

## Recommended Schedule

- Run daily after source Notes data refresh (e.g., 07:00 and 19:00).
- If using Task Scheduler, call:

```bash
python <repo>\src\scripts\etl\run_all_etl.py --db <repo>\src\data\mos_attendance.db --truncate --replace --no-roster-filter
```

---

## Dependencies

Make sure environment includes:

- Python 3.10+
- `pywin32` (for Notes COM access)

Install:

```bash
pip install pywin32
```
