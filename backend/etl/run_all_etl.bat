@echo off
setlocal

REM Run ETL pipeline from repo root
cd /d %~dp0\..\..

echo ========================================
echo Running full ETL pipeline...
echo ========================================

python backend\etl\run_all_etl.py --truncate --replace --no-roster-filter
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
