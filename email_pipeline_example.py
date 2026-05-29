"""
email_pipeline_example.py

Reference implementation of the email pipeline.
Copy and split into email_setup.py / downloader.py / sender.py for production use.

Sections:
  1. Configuration  (email_setup.py)
  2. Downloader     (downloader.py)
  3. Sender         (sender.py)
"""

# ============================================================
# 1. CONFIGURATION  (email_setup.py)
# ============================================================
from pathlib import Path

DESKTOP_PATH = Path.home() / 'Desktop'
_BASE        = DESKTOP_PATH / 'your-project' / 'data'

OSD_DIR = _BASE / 'osd report'
IBD_DIR = _BASE / 'IB Damage'
OVG_DIR = _BASE / 'OVG'

# .env file path — holds POP3/SMTP credentials
ENV_PATH = Path(r'\\your-server\path\to\.env')

# POP3 credentials (.env keys)
POP3_USER_ENV = 'POP3_USERNAME'
POP3_PASS_ENV = 'POP3_PASSWORD'

# POP3
POP3_SERVER = 'pop3.your-mail-server.com'
POP3_PORT   = 995

# Plant codes
PLANT_CODES = ['PLANT1', 'PLANT2', 'PLANT3']

# Email subject match keywords
OSD_SUBJECT_TOKENS = ['OS', 'Report']
IBD_SUBJECT_TOKENS = ['IB Damage Origin']

# SMTP
SMTP_SERVER       = 'smtp.your-mail-server.com'
SMTP_PORT         = 25
SMTP_USER_ENV     = 'SMTP_USERNAME'
SMTP_PASS_ENV     = 'SMTP_PASSWORD'
NOTIFICATION_FROM = 'sender@example.com'

# Download status email
REPORT_SUBJECT = 'OS&D Weekly Report - Download Status'
REPORT_TO  = ['recipient@example.com']
REPORT_CC  = []
REPORT_BCC = []

# Status summary email
STATUS_SUBJECT = 'OS&D Weekly Missing Report Status'
STATUS_TO  = ['recipient@example.com']
STATUS_CC  = []
STATUS_BCC = []

# Failure alert
FAILURE_SUBJECT  = '[ERROR] OS&D Weekly Report'
NOTIFICATION_TO  = ['oncall@example.com']
NOTIFICATION_CC  = []
NOTIFICATION_BCC = []

MAIN_FAILURE_SUBJECT = '[ERROR] OS&D Weekly Missing Report - Processing'
MAIN_FAILURE_TO  = ['oncall@example.com']
MAIN_FAILURE_CC  = []
MAIN_FAILURE_BCC = []

# Per-plant report emails
PLANT_REPORT_SUBJECT = '{week} OS&D Weekly Missing Report - {code}'
PLANT_RECIPIENTS = {
    'PLANT1': {'to': ['plant1@example.com'], 'cc': [], 'bcc': []},
    'PLANT2': {'to': ['plant2@example.com'], 'cc': [], 'bcc': []},
    'PLANT3': {'to': ['plant3@example.com'], 'cc': [], 'bcc': []},
}


# ============================================================
# 2. DOWNLOADER  (downloader.py)
# ============================================================
import os
import re
import poplib
import logging

from dataclasses import dataclass, field
from datetime import datetime, timedelta
from email.parser import Parser
from email.utils import parsedate_to_datetime
from typing import List, Optional
from dotenv import load_dotenv

import pytz

logger = logging.getLogger(__name__)
_tz = pytz.timezone('America/Chicago')

_REPLY_PREFIXES = ('re:', 'fw:', 'fwd:', 'reply:', 'aw:')


@dataclass
class DownloadResult:
    downloaded_osd:    List[str] = field(default_factory=list)
    already_have_osd:  List[str] = field(default_factory=list)
    missing_osd:       List[str] = field(default_factory=list)
    received_times:    dict      = field(default_factory=dict)
    ibd_downloaded:    bool = False
    ibd_received_time: Optional[str] = None
    file_issues:       dict = field(default_factory=dict)
    error:             Optional[str] = None


class EmailDownloader:
    """
    Downloads OS&D report attachments and IBD Origin file from a POP3 mailbox.

    Usage:
        downloader = EmailDownloader(user, password)
        result = downloader.download()
    """

    def __init__(self, user: str, password: str):
        self.user = user
        self.password = password
        self._connection: Optional[poplib.POP3_SSL] = None

    def download(self) -> DownloadResult:
        OSD_DIR.mkdir(parents=True, exist_ok=True)
        IBD_DIR.mkdir(parents=True, exist_ok=True)

        from utils.date_utils import previous_week, previous_week_ibd
        week     = previous_week()
        week_ibd = previous_week_ibd()

        try:
            self._connect()
            return self._scan(week, week_ibd)
        except Exception as e:
            logger.error(f"Download failed: {e}", exc_info=True)
            return DownloadResult(error=str(e))
        finally:
            self._close()

    def _connect(self) -> None:
        logger.info(f"Connecting to {POP3_SERVER}:{POP3_PORT}")
        self._connection = poplib.POP3_SSL(POP3_SERVER, POP3_PORT)
        self._connection.user(self.user)
        self._connection.pass_(self.password)
        logger.info("POP3 login successful")

    def _close(self) -> None:
        if self._connection:
            try:
                self._connection.quit()
                logger.info("POP3 connection closed")
            except Exception as e:
                logger.warning(f"Error during quit: {e}")
            finally:
                self._connection = None

    def _scan(self, week: str, week_ibd: str) -> DownloadResult:
        _, mails_list, _ = self._connection.list()
        all_mails = list(reversed(mails_list))  # newest first

        cutoff = datetime.now(_tz).date()

        downloaded_osd:   List[str] = []
        received_times:   dict = {}
        ibd_downloaded    = False
        ibd_received_time = None
        found_plants: set = set()

        for mail in all_mails:
            idx     = self._parse_index(mail)
            headers = self._fetch_headers(idx)

            try:
                mail_dt   = parsedate_to_datetime(headers.get('Date', '')).astimezone(_tz)
                mail_date = mail_dt.date()
            except Exception:
                continue

            if mail_date < cutoff:
                logger.info(f"Reached mail older than cutoff at index {idx}, stopping")
                break

            raw_subject = headers.get('Subject', '')
            if raw_subject.strip().lower().startswith(_REPLY_PREFIXES):
                continue

            clean_subject = re.sub(r'[&\-–\[\] ]', ' ', raw_subject).lower()
            time_str = mail_dt.strftime('%m/%d/%Y %H:%M CST')

            # Match OS&D reports
            for plant in PLANT_CODES:
                if plant not in found_plants:
                    tokens = OSD_SUBJECT_TOKENS + [week, plant]
                    if all(t.lower() in clean_subject for t in tokens):
                        msg   = self._fetch_email(idx)
                        saved = self._save_attachments(msg, OSD_DIR)
                        if saved:
                            downloaded_osd.append(plant)
                            received_times[plant] = time_str
                            found_plants.add(plant)
                            logger.info(f"OSD saved: {plant} | {headers.get('Subject', '')}")
                        break

            # Match IBD Origin
            if not ibd_downloaded:
                tokens = IBD_SUBJECT_TOKENS + [week_ibd]
                if all(t.lower() in clean_subject for t in tokens):
                    msg   = self._fetch_email(idx)
                    saved = self._save_attachments(msg, IBD_DIR)
                    if saved:
                        ibd_downloaded = True
                        ibd_received_time = time_str
                        logger.info(f"IBD saved | {headers.get('Subject', '')}")

        return DownloadResult(
            downloaded_osd=downloaded_osd,
            missing_osd=[p for p in PLANT_CODES if p not in found_plants],
            received_times=received_times,
            ibd_downloaded=ibd_downloaded,
            ibd_received_time=ibd_received_time,
        )

    @staticmethod
    def _parse_index(mail) -> str:
        return mail.split()[0].decode() if isinstance(mail, bytes) else str(mail)[2:].split()[0].strip()

    def _fetch_headers(self, idx: str):
        _, lines, _ = self._connection.top(idx, 0)
        return Parser().parsestr(b'\r\n'.join(lines).decode('utf-8', 'replace'))

    def _fetch_email(self, idx: str):
        _, lines, _ = self._connection.retr(idx)
        return Parser().parsestr(b'\r\n'.join(lines).decode('utf-8', 'replace'))

    def _save_attachments(self, msg, dest: Path) -> List[str]:
        content_type = msg.get_content_type().lower()
        if content_type.startswith('multipart'):
            results = []
            for part in msg.get_payload():
                results += self._save_attachments(part, dest)
            return results
        if not (content_type.startswith('application') or content_type == 'text/csv'):
            return []
        filename = msg.get_filename()
        if not filename:
            return []
        filename = filename.replace('\r\n', '').replace('\t', ' ').strip()
        if Path(filename).suffix.lower() not in {'.xlsx', '.xlsb', '.xls', '.csv'}:
            return []
        path = dest / filename
        path.write_bytes(msg.get_payload(decode=True))
        logger.info(f"Saved: {path}")
        return [filename]


def load_downloader() -> Optional[EmailDownloader]:
    """Load POP3 credentials from .env and return an EmailDownloader instance."""
    load_dotenv(ENV_PATH)
    user = os.environ.get(POP3_USER_ENV)
    pw   = os.environ.get(POP3_PASS_ENV)
    if not user or not pw:
        logger.error("POP3 credentials missing in .env")
        return None
    return EmailDownloader(user, pw)


# ============================================================
# 3. SENDER  (sender.py)
# ============================================================
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders


class EmailSender:
    def __init__(self, user: str, password: str):
        self.user = user
        self.password = password

    def send(
        self,
        subject: str,
        body: str,
        to: List[str],
        cc: List[str] = None,
        bcc: List[str] = None,
        attachments: List[Path] = None,
    ) -> bool:
        cc          = cc          or []
        bcc         = bcc         or []
        attachments = attachments or []

        root = MIMEMultipart('mixed')
        root['From']    = NOTIFICATION_FROM
        root['To']      = ', '.join(to)
        root['Cc']      = ', '.join(cc)
        root['Date']    = formatdate(localtime=True)
        root['Subject'] = subject
        root.attach(MIMEText(body, 'html', 'utf-8'))

        for path in attachments:
            with open(path, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment', filename=path.name)
            root.attach(part)

        recipients = to + cc + bcc
        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
                smtp.login(self.user, self.password)
                smtp.sendmail(NOTIFICATION_FROM, recipients, root.as_string())
            logger.info(f"Mail sent → {recipients}")
            return True
        except Exception as e:
            logger.error(f"Failed to send mail: {e}", exc_info=True)
            return False


def load_sender() -> Optional[EmailSender]:
    """Load SMTP credentials from .env and return an EmailSender instance."""
    load_dotenv(ENV_PATH)
    user = os.environ.get(SMTP_USER_ENV)
    pw   = os.environ.get(SMTP_PASS_ENV)
    if not user or not pw:
        logger.error("SMTP credentials missing in .env")
        return None
    return EmailSender(user, pw)


def send_status_report(sender: EmailSender, result, week: str) -> bool:
    lines = ["Dear All,", ""]
    if result.success:
        lines += ["<b>Success:</b>"]
        for code, has_report in result.success:
            label = "Missing report attached" if has_report else "No missing items found"
            lines += [f"&nbsp;&nbsp;ㆍ {code} — {label}"]
        lines += [""]
    if result.file_missing:
        lines += ["<b>File Missing:</b>"]
        lines += [f"&nbsp;&nbsp;ㆍ {code}" for code in result.file_missing]
        lines += [""]
    if result.fail:
        lines += ["<b>Fail:</b>"]
        lines += [f"&nbsp;&nbsp;ㆍ {code} — {reason}" for code, reason in result.fail]
        lines += [""]
    lines += [f"Please find attached the OSD OVG Report for {week}.", ""]
    lines += ["Thank you"]
    body = '<BODY style="font-size:11pt;font-family:Calibri">' + "<br>".join(lines) + "</BODY>"
    return sender.send(
        subject=f"{week} {STATUS_SUBJECT}",
        body=body,
        to=STATUS_TO,
        cc=STATUS_CC,
        bcc=STATUS_BCC,
    )


def send_plant_report(sender: EmailSender, code: str, report_path: Path, week: str) -> bool:
    recipients = PLANT_RECIPIENTS.get(code, {})
    if not recipients:
        logger.warning(f"No recipients configured for {code}")
        return False
    lines = ["Dear All,", "",
             f"Please find attached the OSD Missing Items Report for {code} ({week}).",
             "", "Thank you"]
    body = '<BODY style="font-size:11pt;font-family:Calibri">' + "<br>".join(lines) + "</BODY>"
    return sender.send(
        subject=PLANT_REPORT_SUBJECT.format(code=code, week=week),
        body=body,
        to=recipients['to'],
        cc=recipients.get('cc', []),
        bcc=recipients.get('bcc', []),
        attachments=[report_path],
    )


def send_main_failure(sender: EmailSender, log_file: Path, week: str) -> bool:
    from utils.log_utils import get_log_snippet
    snippet = get_log_snippet(log_file)
    lines = ["Dear All,", "",
             "An unexpected error occurred while processing the OS&D Weekly Report.", ""]
    if snippet:
        lines += ["<b>Error log:</b>", f"<pre>{snippet}</pre>", ""]
    lines += ["Thank you"]
    body = '<BODY style="font-size:11pt;font-family:Calibri">' + "<br>".join(lines) + "</BODY>"
    return sender.send(
        subject=f"{week} {MAIN_FAILURE_SUBJECT}",
        body=body,
        to=MAIN_FAILURE_TO,
        cc=MAIN_FAILURE_CC,
        bcc=MAIN_FAILURE_BCC,
    )
