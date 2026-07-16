-- employees
CREATE INDEX IF NOT EXISTS idx_employees_dept ON employees(dept_no);
CREATE INDEX IF NOT EXISTS idx_employees_shift_cat ON employees(shift_cat);
CREATE INDEX IF NOT EXISTS idx_employees_foreign_type ON employees(foreign_type);

-- leave_records
CREATE INDEX IF NOT EXISTS idx_leave_emp ON leave_records(emp_no);
CREATE INDEX IF NOT EXISTS idx_leave_start_dt ON leave_records(start_dt);
CREATE INDEX IF NOT EXISTS idx_leave_status ON leave_records(status);
CREATE INDEX IF NOT EXISTS idx_leave_dept ON leave_records(dept_no);

-- ot_records
CREATE INDEX IF NOT EXISTS idx_ot_emp ON ot_records(emp_no);
CREATE INDEX IF NOT EXISTS idx_ot_start_dt ON ot_records(start_dt);
CREATE INDEX IF NOT EXISTS idx_ot_status ON ot_records(status);
CREATE INDEX IF NOT EXISTS idx_ot_remark ON ot_records(ot_remark);
CREATE INDEX IF NOT EXISTS idx_ot_dept ON ot_records(dept_no);

-- holiday_balances
CREATE INDEX IF NOT EXISTS idx_holiday_dept ON holiday_balances(dept_no);
CREATE INDEX IF NOT EXISTS idx_holiday_import_dt ON holiday_balances(import_dt);

-- daily/weekly/monthly metrics
CREATE INDEX IF NOT EXISTS idx_daily_updated ON daily_metrics(updated_at);
CREATE INDEX IF NOT EXISTS idx_weekly_updated ON weekly_metrics(updated_at);
CREATE INDEX IF NOT EXISTS idx_monthly_updated ON monthly_metrics(updated_at);
