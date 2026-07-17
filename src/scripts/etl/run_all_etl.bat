@echo off
setlocal

REM Run ETL pipeline from repository root
cd /d %~dp0\..\..\..

echo ========================================
echo Running full ETL pipeline...
echo ========================================

python src\scripts\etl\run_all_etl.py --db src/data/mos_attendance.db --truncate --replace --no-roster-filter
if errorlevel 1 (
  echo.
  echo [FAILED] ETL pipeline failed.
  pause
  exit /b 1
)

echo.
echo [OK] ETL pipeline completed.
pause
exit /b 0
