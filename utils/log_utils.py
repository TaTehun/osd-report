import logging
from datetime import date
from pathlib import Path


def setup_logger(log_dir: Path, name: str = 'main'):
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{name}_{date.today().strftime('%Y%m%d')}.log"

    fmt = logging.Formatter('%(asctime)s | %(levelname)-7s | %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setFormatter(fmt)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(fmt)

    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(file_handler)
    root.addHandler(console_handler)

    return logging.getLogger(name), log_file


def get_log_snippet(log_file: Path) -> str:
    if not log_file.exists():
        return '(log file not found)'
    lines = [l.rstrip() for l in log_file.read_text(encoding='utf-8').splitlines()
             if '| ERROR   |' in l or '| WARNING |' in l]
    return '<br>'.join(lines) if lines else '(no warning/error lines found)'


def notify_failure(notifier, subject: str, to: list, cc: list, log_file: Path):
    snippet = get_log_snippet(log_file)
    body = (
        '<BODY style="font-size:11pt;font-family:Calibri">'
        'An error occurred. Please see the log below.<br><br>'
        '<b>Log (WARNING / ERROR):</b><br>'
        f'<pre style="font-size:10pt;font-family:Consolas">{snippet}</pre>'
        '</BODY>'
    )
    notifier.send(subject=subject, body=body, to=to, cc=cc)
