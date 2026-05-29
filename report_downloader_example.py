"""
report_downloader_example.py

Reference implementation of the weekly report downloader.
Runs Mon–Thu to collect OS&D reports from email and validate each file.
main.py runs Friday using the validated files on disk.

Sensitive values (paths, email addresses, plant codes) are replaced with generic placeholders.
"""

import re
import pandas as pd
from utils.log_utils import setup_logger, notify_failure
from email_pipeline_example import load_downloader, load_sender, DownloadResult
from pathlib import Path
from datetime import datetime

LOG_DIR     = Path.home() / 'Desktop' / 'your-project' / 'logs'
MISSING_DIR = Path.home() / 'Desktop' / 'your-project' / 'report'
OSD_DIR     = Path.home() / 'Desktop' / 'your-project' / 'data' / 'osd report'
IBD_DIR     = Path.home() / 'Desktop' / 'your-project' / 'data' / 'IB Damage'

logger, log_file = setup_logger(LOG_DIR, 'report_downloader')

PLANT_CODES = ['PLANT1', 'PLANT2', 'PLANT3']
KEEP_WEEKS  = 5

FAILURE_SUBJECT  = '[ERROR] OS&D Weekly Report'
NOTIFICATION_TO  = ['oncall@example.com']
NOTIFICATION_CC  = []

REPORT_SUBJECT = 'OS&D Weekly Report - Download Status'
REPORT_TO  = ['recipient@example.com']
REPORT_CC  = []
REPORT_BCC = []


def _current_iso_week() -> int:
    return datetime.today().isocalendar()[1]


def _cleanup(directory: Path, pattern: str, keep: int) -> None:
    """Delete files whose week number is more than `keep` weeks behind current week."""
    if not directory.exists():
        return
    current_week = _current_iso_week()
    for f in directory.iterdir():
        if not f.is_file():
            continue
        m = re.search(pattern, f.name, re.IGNORECASE)
        if not m:
            continue
        file_week = int(m.group(1))
        if (current_week - file_week) % 52 > keep:
            f.unlink()
            logger.info(f"Cleaned up: {f.name}")


def _validate_file(path: Path) -> list[str]:
    """
    Validate a downloaded OS&D report file.
    Checks: Summary sheet exists, O/S/D column present, at least one D value,
    and Material/Qty columns present for rows without a Serial Number.
    Returns a list of issue strings (empty = valid).
    """
    try:
        try:
            raw = pd.read_excel(path, sheet_name='Summary', header=None, nrows=10)
        except Exception as e:
            return [f"Failed to open 'Summary' sheet: {e}"]

        if 'O/S/D' not in raw.iloc[0].values:
            return ["'O/S/D' not found in first row"]

        df = pd.read_excel(path, sheet_name='Summary')
        osd_col = next((c for c in df.columns if str(c).strip() == 'O/S/D'), None)

        values = df[osd_col].dropna().astype(str).str.strip().str.upper()
        if 'D' not in values.values:
            return [f"No 'D' value found in O/S/D column (values: {sorted(values.unique())})"]

        sn_col  = next((c for c in df.columns if 'serial' in str(c).strip().lower()), None)
        mat_col = next((c for c in df.columns if 'model' in str(c).strip().lower()
                        or 'material' in str(c).strip().lower()), None)
        qty_col = next((c for c in df.columns if 'qty' in str(c).strip().lower()), None)

        no_sn_count = int((df[sn_col].isna() | (df[sn_col].astype(str).str.strip() == '')).sum()) \
            if sn_col else len(df)
        if no_sn_count > 0 and mat_col is None:
            return [f"Material/Model column missing ({no_sn_count} row(s) without Serial Number)"]
        if mat_col and qty_col is None:
            return ["Qty column missing (required for Material-based matching)"]

        return []

    except Exception as e:
        return [str(e)]


def _send_download_report(sender, result: DownloadResult, week: str) -> bool:
    """Send a status email listing received files, pending files, and any validation issues."""
    lines = ["Dear All,", ""]

    received_today = []
    for plant in result.downloaded_osd:
        received_today.append(
            f"&nbsp;&nbsp;ㆍ {plant} (Received: {result.received_times.get(plant, '')})"
        )
    if result.ibd_downloaded:
        received_today.append(
            f"&nbsp;&nbsp;ㆍ IB Damage Origin (Received: {result.ibd_received_time or ''})"
        )

    lines += ["<b>Files received today:</b>"]
    lines += received_today if received_today else ["&nbsp;&nbsp;ㆍ No files received today"]
    lines += [""]

    pending = [p for p in PLANT_CODES if p not in result.downloaded_osd]
    if not result.ibd_downloaded:
        pending.append("IB Damage Origin")
    if pending:
        lines += ["<b>Pending (not yet received):</b>"]
        lines += [f"&nbsp;&nbsp;ㆍ {p}" for p in pending]
        lines += [""]

    if result.file_issues:
        lines += ["<b>File Issues:</b>"]
        for fname, issues in result.file_issues.items():
            lines += [f"&nbsp;&nbsp;ㆍ <i>{fname}</i>"]
            lines += [f"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- {issue}" for issue in issues]
        lines += [""]

    lines += ["Thank you"]
    body = '<BODY style="font-size:11pt;font-family:Calibri">' + "<br>".join(lines) + "</BODY>"

    return sender.send(
        subject=f"{week} {REPORT_SUBJECT}",
        body=body,
        to=REPORT_TO,
        cc=REPORT_CC,
        bcc=REPORT_BCC,
    )


def run():
    # 1. Download OS&D reports and IBD Origin from email
    downloader = load_downloader()
    if not downloader:
        logger.error("Failed to load downloader, exiting")
        return

    result = downloader.download()
    logger.info(f"Downloaded OSD: {result.downloaded_osd}")
    logger.info(f"Missing OSD:    {result.missing_osd}")
    logger.info(f"IBD downloaded: {result.ibd_downloaded}")
    if result.error:
        logger.error(f"Download error: {result.error}")

    # 2. Validate each downloaded OSD file
    from utils.date_utils import previous_week
    week = previous_week()
    for f in OSD_DIR.iterdir():
        if not f.is_file():
            continue
        plant = next((p for p in PLANT_CODES if p.upper() in f.name.upper()), None)
        if plant not in result.downloaded_osd:
            continue
        issues = _validate_file(f)
        if issues:
            result.file_issues[f.name] = issues
            logger.warning(f"File issues [{f.name}]: {issues}")

    # 3. Send download status report
    sender = load_sender()
    if not sender:
        logger.error("Failed to load sender, exiting")
        return

    if result.error:
        notify_failure(sender, FAILURE_SUBJECT, NOTIFICATION_TO, NOTIFICATION_CC, log_file)
    else:
        _send_download_report(sender, result, week)

    # 4. Cleanup old files (keep last 5 weeks)
    _cleanup(OSD_DIR,  r'W(\d+)',  KEEP_WEEKS)
    _cleanup(IBD_DIR,  r'WK(\d+)', KEEP_WEEKS)
    _cleanup(MISSING_DIR, r'W(\d+)', KEEP_WEEKS)


if __name__ == '__main__':
    run()
