# POS Book Report Automation

## Purpose

This project generates the monthly POS Book Report from the POS SQL Server database and writes the result to an Excel template.

The automation connects directly to the POS database, extracts sales invoice/receipt data, maps the values to the required POS Book Report columns, and creates a new Excel output file per report month.

---

## Confirmed POS Database Source

The POS Book Report source was identified from the POS database structure.

### Main Tables

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
├── RUN_POS_BOOK_REPORT.bat
├── TEST_CONNECTION.bat
│
├── templates/
│   └── APR.xlsx
│
└── output/
```

---

## Required Software

Install the following before running the automation:

1. Python
2. Microsoft ODBC Driver 17 for SQL Server
3. Access to the POS SQL Server database
4. Excel template file inside the `templates` folder

---

## Python Setup

Open CMD or VS Code Terminal inside the project folder.

Check Python:

```powershell
py --version
```

Install requirements:

```powershell
py -m pip install -r requirements.txt
```

---

## Environment Setup

Create a local `.env` file in the project root.

Use `.env.example` as guide:

```env
SQL_SERVER=192.168.xxx.xxx
SQL_DATABASE=VQPBOS
SQL_USERNAME=your_username
SQL_PASSWORD=your_password_here

POS_BRANCH=CAM
REPORT_YEAR=2026
REPORT_MONTH=04

TEMPLATE_FILE=templates/APR.xlsx
OUTPUT_FOLDER=output
```

Important:

Do not commit `.env` to Git because it contains database credentials.

Only `.env.example` should be committed.

---

## Test SQL Connection

Run:

```powershell
py test_connection.py
```

Or double-click:

```text
TEST_CONNECTION.bat
```

Expected successful result:

```text
Connection successful.
Test query result: 1
```

---

## Generate POS Book Report

Run:

```powershell
py main.py
```

Or double-click:

```text
RUN_POS_BOOK_REPORT.bat
```

The generated Excel file will be saved in the `output` folder.

Example output:

```text
output/POS_BOOK_REPORT_CAM_2026_04.xlsx
```

---

## How to Change Report Month

Edit `.env`:

```env
REPORT_YEAR=2026
REPORT_MONTH=04
```

Example for May 2026:

```env
REPORT_YEAR=2026
REPORT_MONTH=05
```

Then run:

```text
RUN_POS_BOOK_REPORT.bat
```

---

## How to Change Branch or Database

Edit `.env`:

```env
SQL_SERVER=192.168.xxx.xxx
SQL_DATABASE=VQPBOS
POS_BRANCH=CAM
```

Use the correct SQL Server IP, database name, and branch code for the target POS database.

---

## Git Setup

Initialize Git:

```powershell
git init
```

Add files:

```powershell
git add .gitignore .env.example requirements.txt main.py test_connection.py RUN_POS_BOOK_REPORT.bat TEST_CONNECTION.bat README.md
```

Check status:

```powershell
git status
```

Make sure `.env` and `output/` are not included.

Commit:

```powershell
git commit -m "Initial POS book report automation setup"
```

---

## Important Notes

1. The report currently uses `HISTMAIN` and `HISTSUB` as the main source.
2. `TTRNJRNL` is not used in the automation because it contains printed receipt text and is better used only for validation.
3. Senior and PWD discount logic depends on `HISTMAIN.disccode`.
4. Other discounts such as SPECIAL are mapped to the OTHERS column.
5. The Excel template is expected to have headers on row 8 and data starting on row 9.
6. Generated reports are saved in the `output` folder.
7. The `output` folder is ignored by Git because reports are generated files.

---

## Troubleshooting

### Error: Login timeout expired

Possible causes:

* Wrong SQL Server IP
* SQL Server not reachable from the computer
* SQL Server TCP/IP not enabled
* Firewall blocking SQL Server port
* Wrong database instance/server name

### Error: Login failed for user

Possible causes:

* Wrong username
* Wrong password
* User has no access to the POS database

### Error: ODBC Driver not found

Install Microsoft ODBC Driver 17 for SQL Server.

### Error: Template file not found

Check `.env`:

```env
TEMPLATE_FILE=templates/APR.xlsx
```

Make sure the Excel template exists in the `templates` folder.

---

## Current Version

Version: POS Book Report Automation v1

Confirmed working source query:

* `HISTMAIN`
* `HISTSUB`
* `MDISC`

Current output:

* Monthly POS Book Report Excel file
* One row per OR/receipt
* Branch/month controlled by `.env`
