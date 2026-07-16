PRAGMA foreign_keys = ON;

-- Employees master
CREATE TABLE IF NOT EXISTS employees (
  emp_no TEXT PRIMARY KEY,
  emp_name TEXT NOT NULL,
  dept_no TEXT,
  foreign_type TEXT CHECK (foreign_type IN ('本勞','外勞') OR foreign_type IS NULL),
  shift_cat TEXT CHECK (shift_cat IN ('輪班','常日') OR shift_cat IS NULL),
  shift_raw TEXT,
  active INTEGER NOT NULL DEFAULT 1,
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- Leave records (Raw data)
CREATE TABLE IF NOT EXISTS leave_records (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  emp_no TEXT NOT NULL,
  emp_name TEXT,
  dept_no TEXT,
  date_composed TEXT,
  start_dt TEXT,
  end_dt TEXT,
  leave_type TEXT,
  status TEXT,
  description TEXT,
  hours REAL NOT NULL DEFAULT 0,
  sign_status_raw TEXT,
  source TEXT NOT NULL DEFAULT 'notes_rawdata',
  raw_json TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (emp_no) REFERENCES employees(emp_no)
);

-- Overtime records
CREATE TABLE IF NOT EXISTS ot_records (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  emp_no TEXT NOT NULL,
  emp_name TEXT,
  dept_no TEXT,
  start_dt TEXT,
  end_dt TEXT,
  apply_item TEXT,
  ot_remark TEXT,
  status TEXT,
  description TEXT,
  hours REAL NOT NULL DEFAULT 0,
  foreign_type TEXT,
  shift_cat TEXT,
  sign_status_raw TEXT,
  source TEXT NOT NULL DEFAULT 'notes_overtime',
  raw_json TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (emp_no) REFERENCES employees(emp_no)
);

-- Holiday balances (latest by employee)
CREATE TABLE IF NOT EXISTS holiday_balances (
  emp_no TEXT PRIMARY KEY,
  emp_name TEXT,
  dept_no TEXT,
  shift_raw TEXT,
  import_dt TEXT,

  l_total REAL NOT NULL DEFAULT 0,
  l_initial REAL NOT NULL DEFAULT 0,
  l_deferred REAL NOT NULL DEFAULT 0,
  l_used REAL NOT NULL DEFAULT 0,
  l_adjust REAL NOT NULL DEFAULT 0,
  l_remain REAL NOT NULL DEFAULT 0,

  c_total REAL NOT NULL DEFAULT 0,
  c_deferred REAL NOT NULL DEFAULT 0,
  c_adjust REAL NOT NULL DEFAULT 0,
  c_from_ot REAL NOT NULL DEFAULT 0,
  c_used REAL NOT NULL DEFAULT 0,
  c_def_remain REAL NOT NULL DEFAULT 0,
  c_remain REAL NOT NULL DEFAULT 0,

  total_leave REAL NOT NULL DEFAULT 0,
  source TEXT NOT NULL DEFAULT 'notes_holiday',
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (emp_no) REFERENCES employees(emp_no)
);

-- Daily aggregated metrics (for dashboard)
CREATE TABLE IF NOT EXISTS daily_metrics (
  day TEXT PRIMARY KEY,             -- YYYY-MM-DD
  tw_on REAL NOT NULL DEFAULT 0,
  foreign_on REAL NOT NULL DEFAULT 0,
  normal_on REAL NOT NULL DEFAULT 0,

  tw_ot REAL NOT NULL DEFAULT 0,
  foreign_ot REAL NOT NULL DEFAULT 0,
  normal_ot REAL NOT NULL DEFAULT 0,

  tw_leave REAL NOT NULL DEFAULT 0,
  foreign_leave REAL NOT NULL DEFAULT 0,
  normal_leave REAL NOT NULL DEFAULT 0,

  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- Weekly aggregated metrics
CREATE TABLE IF NOT EXISTS weekly_metrics (
  iso_week TEXT PRIMARY KEY,        -- e.g. 2026-W27
  tw_ot REAL NOT NULL DEFAULT 0,
  foreign_ot REAL NOT NULL DEFAULT 0,
  normal_ot REAL NOT NULL DEFAULT 0,
  tw_leave REAL NOT NULL DEFAULT 0,
  foreign_leave REAL NOT NULL DEFAULT 0,
  normal_leave REAL NOT NULL DEFAULT 0,
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- Monthly aggregated metrics
CREATE TABLE IF NOT EXISTS monthly_metrics (
  ym TEXT PRIMARY KEY,              -- e.g. 2026-07
  tw_ot REAL NOT NULL DEFAULT 0,
  foreign_ot REAL NOT NULL DEFAULT 0,
  normal_ot REAL NOT NULL DEFAULT 0,
  tw_leave REAL NOT NULL DEFAULT 0,
  foreign_leave REAL NOT NULL DEFAULT 0,
  normal_leave REAL NOT NULL DEFAULT 0,
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);
