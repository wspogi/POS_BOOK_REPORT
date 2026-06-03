# POS Book Report Automation

## Purpose

This project automatically generates the monthly POS Book Report from the POS SQL Server database and writes the result to an Excel template.

The automation is designed for POS database reporting where one row represents one OR / receipt / sales invoice.

---

## What This Automation Does

1. Connects to the POS SQL Server database.
2. Extracts sales invoice / receipt data from POS tables.
3. Parses customer information.
4. Parses VAT breakdown.
5. Maps Senior, PWD, and Other discounts.
6. Writes the result to an Excel template.
7. Saves the final Excel file in the output folder.
8. Creates logs in the logs folder.
9. Can be scheduled using Windows Task Scheduler.

---

## Confirmed POS Database Source

| Table      | Purpose                                                                                                 |
| ---------- | ------------------------------------------------------------------------------------------------------- |
| `HISTMAIN` | Main receipt/header table. One row per receipt/OR.                                                      |
| `HISTSUB`  | Transaction detail table. Contains customer details, VAT breakdown, item lines, and discount breakdown. |
| `MDISC`    | Discount master table. Used to identify discount code meanings.                                         |
| `TTRNJRNL` | Receipt printed text/journal. Used only as backup validation/reference.                                 |

---

## Confirmed Field Mapping

| Excel Column  | Source / Logic                                                             |
| ------------- | -------------------------------------------------------------------------- |
| DATE          | `HISTMAIN.trandate`                                                        |
| OR NO         | `HISTMAIN.receipt`                                                         |
| ID NO         | Blank for now                                                              |
| NAME          | Parsed from `HISTSUB.reference` where `type = 'M'` and `code = 'CT'`       |
| ADDRESS       | Parsed from `HISTSUB.reference` where `type = 'M'` and `code = 'CT'`       |
| TIN NO        | Parsed from `HISTSUB.reference` where `type = 'M'` and `code = 'CT'`       |
| TOTAL SALES   | `HISTMAIN.amount`                                                          |
| VATABLE SALES | Parsed from `HISTSUB.reference` where `type = 'U'` and `code = '12'`       |
| VAT EXEMPT    | Default `0.00`                                                             |
| ZERO RATED    | Default `0.00`                                                             |
| VAT 12%       | Parsed from `HISTSUB.reference` where `type = 'U'` and `code = '12'`       |
| SENIOR        | `HISTMAIN.discamt` when `HISTMAIN.disccode = '05'`                         |
| PWD           | `HISTMAIN.discamt` when `HISTMAIN.disccode = 'A1'`                         |
| OTHERS        | `HISTMAIN.discamt` when discount code is not blank, not `05`, and not `A1` |
| NET SALES     | Parsed from VAT row in `HISTSUB.reference`, fallback to VATABLE SALES      |

---

## Confirmed Discount Codes

From table `MDISC`:

| Discount Code | Description  | POS Book Column |
| ------------- | ------------ | --------------- |
| `05`          | SENIOR DISC  | SENIOR          |
| `A1`          | PWD DISCOUNT | PWD             |
| `06`          | SPECIAL      | OTHERS          |
| `01`          | EDIS         | OTHERS          |
| `07`          | OPEN AMT     | OTHERS          |

---

## Folder Structure

```text
POS_BOOK_REPORT/
│
├── main.py
├── test_connection.py
├── requirements.txt
├── .env.example
├── .gitignore
├── README.md
│
├── INSTALL_PYTHON.bat
├── INSTALL_REQUIREMENTS.bat
├── TEST_CONNECTION.bat
├── RUN_POS_BOOK_REPORT.bat
├── CREATE_TASK_SCHEDULER.bat
├── DELETE_TASK_SCHEDULER.bat
│
├── templates/
│   └── APR.xlsx
│
├── output/
│
└── logs/
```

---

## Required Software

Before running the automation, make sure the computer has:

1. Python
2. Microsoft ODBC Driver 17 for SQL Server
3. Network access to the POS SQL Server database
4. Excel template file inside the `templates` folder

---

## Python Installation

If Python is not yet installed, run:

```text
INSTALL_PYTHON.bat
```

This BAT file installs Python using `winget`.

After installation, close and reopen CMD or VS Code Terminal, then check:

```powershell
py --version
```

---

## Install Python Requirements

Run:

```text
INSTALL_REQUIREMENTS.bat
```

Or manually run:

```powershell
py -m pip install -r requirements.txt
```

Required Python packages:

```text
pandas
pyodbc
python-dotenv
openpyxl
```

---

## Environment Setup

Create a local `.env` file in the project root.

Use `.env.example` as guide.

Example:

```env
SQL_SERVER=192.168.xxx.xxx
SQL_DATABASE=VQPBOS
SQL_USERNAME=your_username
SQL_PASSWORD=your_password_here

POS_BRANCH=CAM

DATE_MODE=PREVIOUS_MONTH
# DATE_FROM=2026-04-01
# DATE_TO=2026-04-30

TEMPLATE_FILE=templates/APR.xlsx

OUTPUT_DIR=output
LOG_DIR=logs
```

Important:

Do not commit `.env` to Git because it contains database credentials.

Only `.env.example` should be committed.

---

## Date Mode Options

### Option 1: Automatic Previous Month

Use this for scheduled monthly automation:

```env
DATE_MODE=PREVIOUS_MONTH
# DATE_FROM=2026-04-01
# DATE_TO=2026-04-30
```

Example:

If today is June 2026, the report period will be May 1 to May 31, 2026.

The SQL filter will use:

```sql
HM.trandate >= '2026-05-01'
AND HM.trandate < '2026-06-01'
```

---

### Option 2: Custom Date Range

Use this for manual testing or rerun:

```env
DATE_MODE=CUSTOM
DATE_FROM=2026-04-01
DATE_TO=2026-04-30
```

The `DATE_TO` value in `.env` is inclusive.

The Python script automatically converts it to an exclusive SQL end date.

Example:

```env
DATE_FROM=2026-04-01
DATE_TO=2026-04-30
```

SQL will use:

```sql
HM.trandate >= '2026-04-01'
AND HM.trandate < '2026-05-01'
```

---

## Test SQL Connection

Run:

```text
TEST_CONNECTION.bat
```

Or manually:

```powershell
py test_connection.py
```

Expected successful result:

```text
Connection successful.
Test query result: 1
```

---

## Generate POS Book Report

Run:

```text
RUN_POS_BOOK_REPORT.bat
```

Or manually:

```powershell
py main.py
```

The generated Excel file will be saved in the `output` folder.

Example output:

```text
output/POS_BOOK_REPORT_CAM_2026_04.xlsx
```

---

## Logs

Each run creates a log file inside the `logs` folder.

Example:

```text
logs/POS_BOOK_REPORT_20260603_080000.log
```

Logs are useful when the script is run by Task Scheduler because the CMD window may close after execution.

---

## Task Scheduler Setup

To create a scheduled task, run:

```text
CREATE_TASK_SCHEDULER.bat
```

Default schedule:

```text
Every 1st day of the month at 8:00 AM
```

Recommended `.env` setting for scheduled task:

```env
DATE_MODE=PREVIOUS_MONTH
```

This means every month, the task will automatically generate the previous month’s report.

Example:

If the task runs on June 1, it will generate the May report.

---

## Delete Scheduled Task

If you need to remove the scheduled task, run:

```text
DELETE_TASK_SCHEDULER.bat
```

---

## Git Setup

Initialize Git:

```powershell
git init
```

Add files:

```powershell
git add .gitignore .env.example requirements.txt main.py test_connection.py README.md INSTALL_PYTHON.bat INSTALL_REQUIREMENTS.bat TEST_CONNECTION.bat RUN_POS_BOOK_REPORT.bat CREATE_TASK_SCHEDULER.bat DELETE_TASK_SCHEDULER.bat
```

Check status:

```powershell
git status
```

Make sure these are not included:

```text
.env
output/
logs/
```

Commit:

```powershell
git commit -m "Initial POS book report automation setup"
```

---

## Update Git After Changes

After editing files:

```powershell
git status
git add .
git status
git commit -m "Update POS book report automation"
```

Before committing, always check that `.env`, `output/`, and `logs/` are not included.

---

## Troubleshooting

### Error: Python was not found

Run:

```text
INSTALL_PYTHON.bat
```

Then reopen CMD or VS Code Terminal and run:

```powershell
py --version
```

---

### Error: ODBC Driver not found

Install Microsoft ODBC Driver 17 for SQL Server.

---

### Error: Login timeout expired

Possible causes:

1. Wrong SQL Server IP
2. SQL Server is not reachable from the computer
3. SQL Server TCP/IP is not enabled
4. Firewall is blocking SQL Server port
5. Wrong database instance/server name
6. Computer is not connected to the required network/VPN

---

### Error: Login failed for user

Possible causes:

1. Wrong username
2. Wrong password
3. SQL user has no access to the POS database

---

### Error: Template file not found

Check `.env`:

```env
TEMPLATE_FILE=templates/APR.xlsx
```

Make sure the Excel template exists in the `templates` folder.

---

### Task Scheduler Created But Report Did Not Generate

Check:

1. `logs` folder
2. `.env` file exists
3. Python is installed
4. Requirements are installed
5. SQL Server is reachable
6. The task was created using the correct project folder
7. The task has permission to access the folder

---

## Current Version

Version: POS Book Report Automation v1

Confirmed source query:

* `HISTMAIN`
* `HISTSUB`
* `MDISC`

Current output:

* Monthly POS Book Report Excel file
* One row per OR/receipt
* Branch/month controlled by `.env`
* Supports `DATE_MODE=PREVIOUS_MONTH`
* Supports `DATE_MODE=CUSTOM`
* Supports Windows Task Scheduler
