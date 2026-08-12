# OS&D Weekly Missing Report · Automation Pipeline

A weekly Python automation pipeline that downloads OS&D (Over, Short & Damage) reports and IBD Origin files from email, compares them to identify missing items, and distributes plant-level missing reports via email, with no manual steps.

> **Visual overview:** [project-overview.html](project-overview.html)

---

## Overview

Each week, plant teams submit OS&D reports to a shared mailbox, and an IBD Origin file is prepared by the inbound team. This pipeline:

1. **Downloads** this week's OS&D report attachments and IBD Origin file from the POP3 mailbox
2. **Validates** each downloaded OS&D file for required columns and data integrity
3. **Compares** IBD Origin vs OS&D reports per plant: items in IBD but not in OS&D are flagged as **Missing**
4. **Checks** Missing items against the OVG (Over/Good) accumulation file, deducting items found in OVG
5. **Saves** final Missing items as plant-level Excel reports
6. **Updates** the OVG file: new items from OS&D not in IBD are accumulated; records older than 15 weeks are pruned
7. **Distributes** plant-level missing reports and a status summary via SMTP email

---

## Project Structure

```
osd-weekly-report/
├── main.py                        # Core processing: IBD vs OSD comparison, OVG, report generation
├── report_downloader.py           # POP3 downloader entry point + file validation
├── run_report_downloader.bat      # Task Scheduler trigger for downloader
├── requirements.txt
├── email_pipeline/
│   ├── email_setup.py             # All settings (paths, SMTP, plant recipients)
│   ├── downloader.py              # POP3 attachment downloader (OS&D + IBD)
│   └── sender.py                  # SMTP sender (status report, plant report, failure alert)
├── utils/
│   ├── date_utils.py              # Week/date formatting utilities
│   ├── log_utils.py               # Logger setup + log snippet extraction
│   └── pivot_utils.py             # Pivot table refresh + HTML rendering
├── study/                         # Annotated study files for each module
└── brd/                           # Business requirement documents
```

---

## Data Sources

| Source | Description |
|--------|-------------|
| OS&D Reports | One Excel file per plant per week; Summary sheet with O/S/D column |
| IBD Origin | One Excel/xlsb file per week; IB Damage data with Material, Serial number, Quantity |

> If an OS&D file is missing for a plant, that plant is skipped and reported in the status email.

---

## Output

All files are organized by week:

```
Desktop/osd weekly report/
├── data/
│   ├── osd report/                # Downloaded OS&D report files (kept 5 weeks)
│   ├── IB Damage/                 # Downloaded IBD Origin files (kept 5 weeks)
│   └── OVG/
│       └── OSD_OVG.xlsx           # Accumulated OVG data (15 weeks rolling)
├── report/
│   └── W21/
│       ├── W21 S356 Missing items from OSD Report.xlsx
│       ├── W21 S377 Missing items from OSD Report.xlsx
│       └── ...
└── logs/
    └── main_20260528.log
```

---

## How It Works

### 1. Download (`report_downloader.py`)
Connects to the POP3 mailbox and scans emails newest-first, stopping at emails older than today. Matches OS&D reports by subject tokens (`OS`, `Report`, week number, plant code) and IBD Origin by subject tokens (`IB Damage Origin`, week number). Saves attachments to the configured folders.

After downloading, each OS&D file is validated:
- Summary sheet must exist
- O/S/D column must be present in the header row
- At least one `D` value must exist in the O/S/D column
- Rows without a Serial Number must have a Material/Model column and a Qty column

### 2. IBD vs OSD Comparison (`main.py → find_new_missing`)
For each plant, IBD rows are compared against OS&D D-rows using `_match_and_consume`:

| Condition | Matching Method |
|-----------|----------------|
| IBD has Serial Number | 1:1 match against OSD `Serial Number(s)`; matched OSD row is consumed |
| IBD has no Serial Number | Match by `Material` vs OSD `Model`; OSD qty is summed and deducted |

IBD rows with remaining qty after matching → **Missing candidates**
OSD D-rows not consumed by any IBD row → **OVG new** (OS&D has it, IBD doesn't)

### 3. OVG Check (`main.py → check_missing_against_ovg`)
Missing candidates are checked against the existing OVG file. Items found in OVG are deducted from Missing using the same `_match_and_consume` logic. What remains after deduction is the **final Missing**.

### 4. OVG Update (`main.py → save_ovg`)
- Remaining OVG (not matched by this week's Missing) is kept
- New OVG rows from this week's comparison are appended
- Records older than 15 weeks are pruned automatically
- OSD columns are preserved exactly: `Plant, Week, Date, Container, SN, Live/Drop, Seal in Tact, O/S/D, Model, SN QTY, Qty, Serial Number(s), WH Note`

### 5. Distribution (`sender.py`)
- **Status summary email**: sent to the configured list with per-plant Success/Fail status and OVG report notice
- **Plant-level report emails**: each plant's missing report is attached and sent to plant-specific recipients
- **Failure alert**: sent if an unexpected error occurs during processing, with log snippet attached

---

## Configuration

All settings are managed in `email_pipeline/email_setup.py`:

| Setting | Description |
|---------|-------------|
| `OSD_DIR` / `IBD_DIR` / `OVG_DIR` | Paths to data folders |
| `PLANT_CODES` | List of all plant codes to process |
| `OSD_SUBJECT_TOKENS` / `IBD_SUBJECT_TOKENS` | Email subject match keywords |
| `SMTP_SERVER` / `SMTP_PORT` | SMTP server settings |
| `STATUS_TO` / `STATUS_CC` | Status summary email recipients |
| `PLANT_RECIPIENTS` | Per-plant report email recipients `{code: {to, cc, bcc}}` |
| `MAIN_FAILURE_TO` | Failure alert recipients |

Credentials (POP3/SMTP username and password) are stored in a `.env` file on the network drive and loaded at runtime.

---

## Warehouse → Plant Mapping

| Warehouse Code | Plant Code |
|---------------|-----------|
| WH1 | PLANT1 |
| WH2 | PLANT2 |
| WH3 | PLANT3 |
| WH4 | PLANT4 |
| WH5 | PLANT5 |
| WH6 | PLANT6 |
| WH7 | PLANT7 |
| WH8 | PLANT8 |
| WH9 | PLANT9 |

---

## Requirements

```
pandas==3.0.3
openpyxl==3.1.5
pyxlsb==1.0.10
pytz==2026.2
python-dotenv==1.2.2
```

---

## Running

```bash
# Download OS&D reports and IBD Origin from email
python report_downloader.py
# or double-click run_report_downloader.bat

# Process reports and generate missing item reports
python main.py
```

Intended to be triggered weekly (after OS&D reports are received) via Windows Task Scheduler.

---

## Changelog

### v1.0 · 2026-05-28 · Initial Release
- POP3 downloader for OS&D reports + IBD Origin
- IBD vs OSD comparison with Serial Number and Material+Qty matching
- OVG accumulation file (15-week rolling, OSD column schema)
- Plant-level missing report generation and SMTP distribution
- File validation on download (Summary sheet, O/S/D column, D values)
- Automatic cleanup of files older than 5 weeks
