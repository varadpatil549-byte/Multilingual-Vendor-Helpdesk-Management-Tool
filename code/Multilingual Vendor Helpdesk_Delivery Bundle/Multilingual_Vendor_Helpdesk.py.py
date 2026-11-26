import sys
import csv
import uuid
import time
import threading
import queue
import re
import hashlib
from pathlib import Path
from datetime import datetime

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont, QColor, QPalette
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QTextEdit, QComboBox, QTableWidget,
    QTableWidgetItem, QMessageBox, QSplitter, QGroupBox, QHeaderView,
    QAbstractItemView, QAbstractScrollArea, QScrollArea, QStatusBar, QFrame
)

import ollama
import win32com.client
from langdetect import detect, LangDetectException


APP_TITLE = "Multilingual Vendor Helpdesk Management Tool"
# Deployment V1: Outlook folder/shared-mailbox configuration remains unchanged.
# Subject triage is a separate compact translation pass. Integrity percentage
# is a deterministic preservation check, not AI/model confidence.
MODEL = "translategemma:4b"
OLLAMA_KEEP_ALIVE = -1
OLLAMA_OPTIONS = {"temperature": 0.0, "num_ctx": 1024, "num_predict": 384}
OUTLOOK_FOLDER_NAME = "Multilingual"
LOG_FILE = Path(__file__).resolve().parent / "multilingual_helpdesk_log.csv"
CACHE_FILE = Path(__file__).resolve().parent / "multilingual_translation_cache.csv"

LOG_FIELDS = [
    "session_id",
    "outlook_entry_id",
    "source",
    "sender",
    "received_time",
    "source_language",
    "response_language",
    "original_subject",
    "translated_subject",
    "priority",
    "translation_integrity_percent",
    "integrity_status",
    "query_word_count",
    "query_translation_seconds",
    "response_translation_seconds",
    "total_elapsed_seconds",
    "status",
    "last_updated",
]

LANGUAGES = [
    ("Auto Detect", ""),
    ("German", "German"),
    ("French", "French"),
    ("Spanish", "Spanish"),
    ("Italian", "Italian"),
    ("Dutch", "Dutch"),
    ("Portuguese", "Portuguese"),
    ("Polish", "Polish"),
    ("Swedish", "Swedish"),
    ("Danish", "Danish"),
    ("Norwegian", "Norwegian"),
    ("Finnish", "Finnish"),
    ("Czech", "Czech"),
    ("Romanian", "Romanian"),
    ("Hungarian", "Hungarian"),
    ("Turkish", "Turkish"),
    ("English", "English"),
]

LANG_CODE_TO_NAME = {
    "de": "German",
    "fr": "French",
    "es": "Spanish",
    "it": "Italian",
    "nl": "Dutch",
    "pt": "Portuguese",
    "pl": "Polish",
    "sv": "Swedish",
    "da": "Danish",
    "no": "Norwegian",
    "fi": "Finnish",
    "cs": "Czech",
    "ro": "Romanian",
    "hu": "Hungarian",
    "tr": "Turkish",
    "en": "English",
}


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

def now_string():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def format_received_time(value):
    """Convert Outlook COM ReceivedTime into a compact, reliable display."""
    raw = str(value or "").strip()
    if not raw:
        return "—"

    candidates = [
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%d-%m-%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M:%S%z",
        "%Y-%m-%d %H:%M:%S.%f",
        "%Y-%m-%d %H:%M:%S.%f%z",
    ]

    normalized = raw.replace("Z", "+00:00")
    try:
        parsed = datetime.fromisoformat(normalized)
        return parsed.strftime("%a %H:%M")
    except Exception:
        pass

    for fmt in candidates:
        try:
            parsed = datetime.strptime(raw, fmt)
            return parsed.strftime("%a %H:%M")
        except Exception:
            continue

    # Last-resort extraction for locale-specific Outlook strings.
    match = re.search(r"(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?", raw, re.I)
    if match:
        hour = int(match.group(1))
        minute = match.group(2)
        ampm = match.group(3)
        if ampm:
            return f"{hour:02d}:{minute} {ampm.upper()}"
        return f"{hour:02d}:{minute}"

    return raw[-16:] if len(raw) > 16 else raw


def word_count(text):
    return len((text or "").split())


def detect_language_code(text):
    text = (text or "").strip()
    if len(text) < 10:
        return ""
    try:
        return detect(text)
    except LangDetectException:
        return ""
    except Exception:
        return ""


def language_name(code):
    return LANG_CODE_TO_NAME.get(code, code.upper() if code else "Unknown")


def clean_model_text(text):
    return (text or "").replace("```text", "").replace("```", "").strip()


def parse_translation(raw_text):
    """Parse TranslateGemma's requested SUBJECT/BODY response."""
    text = clean_model_text(raw_text)
    if not text:
        return "", ""

    lines = text.splitlines()
    subject = ""
    body = text
    body_start = None

    for index, line in enumerate(lines):
        stripped = line.strip()
        upper = stripped.upper()
        if upper.startswith("SUBJECT:") and not subject:
            subject = stripped.split(":", 1)[1].strip()
        elif upper in ("BODY:", "BODY") and body_start is None:
            body_start = index + 1

    if body_start is not None:
        body = "\n".join(lines[body_start:]).strip()
    elif subject:
        subject_lines = []
        skipped_subject = False
        for line in lines:
            if not skipped_subject and line.strip().upper().startswith("SUBJECT:"):
                skipped_subject = True
                continue
            subject_lines.append(line)
        body = "\n".join(subject_lines).strip()

    return clean_model_text(subject), clean_model_text(body)



# -----------------------------------------------------------------------------
# Queue priority + deterministic translation integrity
# -----------------------------------------------------------------------------

PRIORITY_WORDS = {
    "urgent", "urgently", "important", "priority", "high", "asap",
    "immediately", "critical", "emergency", "dringend", "wichtig",
    "sofort", "eilig", "priorität", "prioritaet", "importante", "urgente",
    "prioridad", "prioritaire", "priorité", "belangrijk", "spoed",
}

INTEGRITY_PATTERNS = [
    r'\b[A-Z]{2,8}[-/]?\d{3,}(?:[-/][A-Z0-9]+)*\b',
    r'\b\d{1,4}[./-]\d{1,2}[./-]\d{2,4}\b',
    r'\b\d{4}[./-]\d{1,2}[./-]\d{1,2}\b',
    r'\b(?:EUR|USD|GBP|CHF|SEK|NOK|DKK|PLN|CZK|HUF|TRY|€|\$|£)\s*[\d.,]+\b',
    r'\b[\d.,]+\s*(?:EUR|USD|GBP|CHF|SEK|NOK|DKK|PLN|CZK|HUF|TRY)\b',
    r'\b[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}\b',
    r'https?://[^\s<>"\']+',
]

def critical_tokens(text):
    """Extract business-critical values that should survive translation."""
    found = []
    for pattern in INTEGRITY_PATTERNS:
        found.extend(re.findall(pattern, text or "", flags=re.IGNORECASE))
    result = []
    seen = set()
    for token in found:
        normalized = token.strip().rstrip(".,;:)")
        key = normalized.lower()
        if key and key not in seen:
            seen.add(key)
            result.append(normalized)
    return result


def translation_integrity(original, translated):
    """
    Deterministic preservation score. This is NOT model confidence and does
    not claim to measure linguistic quality.
    """
    tokens = critical_tokens(original)
    if not tokens:
        return 100, [], "NO_CRITICAL_VALUES"
    translated_lower = (translated or "").lower()
    missing = [token for token in tokens if token.lower() not in translated_lower]
    score = round((len(tokens) - len(missing)) / len(tokens) * 100)
    return score, missing, "VERIFIED" if not missing else "REVIEW_REQUIRED"


def priority_level(original_subject, translated_subject=""):
    """Check original + English subject for urgency/importance wording."""
    combined = f"{original_subject or ''} {translated_subject or ''}".lower()
    compact = re.sub(r"[^a-zA-ZÀ-ÿ0-9]+", " ", combined)
    words = set(compact.split())
    phrases = (
        "high priority", "top priority", "asap", "action required",
        "sofort bearbeiten", "très urgent", "muy urgente",
    )
    return (
        "URGENT / IMPORTANT"
        if words & PRIORITY_WORDS or any(p in compact for p in phrases)
        else "NORMAL"
    )


def translate_subject_only(subject, target_language="English"):
    """Compact subject-only translation used for early queue triage."""
    if not (subject or "").strip():
        return ""
    prompt = f"""Translate this email subject to {target_language}.
Preserve invoice numbers, PO numbers, dates, amounts, names, company names,
IDs and reference values exactly. Do not explain. Return only the translation.

SUBJECT:
{subject}"""
    result = ollama.chat(
        model=MODEL,
        messages=[{"role": "user", "content": prompt}],
        keep_alive=OLLAMA_KEEP_ALIVE,
        options={"temperature": 0.0, "num_ctx": 384, "num_predict": 64},
    )
    return clean_model_text(result.get("message", {}).get("content", ""))


def translation_options(subject, body):
    """Keep context/output budgets as small as the current email allows."""
    chars = len(subject or "") + len(body or "")
    if chars <= 2500:
        return {"temperature": 0.0, "num_ctx": 768, "num_predict": 256}
    if chars <= 5000:
        return {"temperature": 0.0, "num_ctx": 1024, "num_predict": 384}
    if chars <= 9000:
        return {"temperature": 0.0, "num_ctx": 1536, "num_predict": 512}
    return {"temperature": 0.0, "num_ctx": 2048, "num_predict": 640}


def translate_text(subject, body, target_language):
    """Fast local translation using only the current message and a dynamic budget."""
    prompt = f"""Translate this business email to {target_language}.
Preserve invoice/PO/reference IDs, dates, amounts/currencies, names, email addresses and URLs exactly.
Do not answer or explain. Return only the translated SUBJECT and BODY.

SUBJECT: {subject}
BODY:
{body}"""

    result = ollama.chat(
        model=MODEL,
        messages=[{"role": "user", "content": prompt}],
        keep_alive=OLLAMA_KEEP_ALIVE,
        options=translation_options(subject, body),
    )
    content = result.get("message", {}).get("content", "")
    return parse_translation(content)


def warmup_model():
    """Load TranslateGemma once so the first real translation avoids model-load latency."""
    try:
        ollama.chat(
            model=MODEL,
            messages=[{"role": "user", "content": "Translate Hello to German. Return only the translation."}],
            keep_alive=OLLAMA_KEEP_ALIVE,
            options={"temperature": 0.0, "num_ctx": 256, "num_predict": 8},
        )
        return True, ""
    except Exception as exc:
        return False, str(exc)


# -----------------------------------------------------------------------------
# CSV recovery log
# -----------------------------------------------------------------------------

class LogStore:
    def __init__(self, path):
        self.path = Path(path)
        self.rows = []
        self.load()

    def load(self):
        if not self.path.exists():
            self.rows = []
            return
        try:
            with self.path.open("r", newline="", encoding="utf-8-sig") as handle:
                self.rows = list(csv.DictReader(handle))
        except Exception:
            self.rows = []

    def save(self):
        temp_path = self.path.with_suffix(".tmp")
        with temp_path.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=LOG_FIELDS)
            writer.writeheader()
            writer.writerows(self.rows)
        temp_path.replace(self.path)

    def find(self, session_id="", entry_id=""):
        for row in reversed(self.rows):
            if session_id and row.get("session_id") == session_id:
                return row
            if entry_id and row.get("outlook_entry_id") == entry_id:
                return row
        return None

    def upsert(self, data):
        session_id = data.get("session_id", "")
        entry_id = data.get("outlook_entry_id", "")
        row = self.find(session_id=session_id, entry_id=entry_id)

        if row is None:
            row = {field: "" for field in LOG_FIELDS}
            row["session_id"] = session_id or str(uuid.uuid4())
            self.rows.append(row)

        for key, value in data.items():
            if key in LOG_FIELDS:
                row[key] = str(value)

        row["last_updated"] = now_string()
        self.save()
        return row


# -----------------------------------------------------------------------------
# Translation cache + latest-message extraction
# -----------------------------------------------------------------------------

CACHE_FIELDS = [
    "outlook_entry_id",
    "source_language",
    "original_subject",
    "translated_subject",
    "query_hash",
    "original_query",
    "translated_query",
    "integrity_percent",
    "integrity_status",
    "translated_at",
]


def text_hash(subject, body):
    payload = f"{subject.strip()}\n{body.strip()}".encode("utf-8", errors="ignore")
    return hashlib.sha256(payload).hexdigest()


def extract_latest_message(body):
    """Return only the newest message and stop before common quoted-mail trails."""
    body = (body or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    if not body:
        return ""

    markers = [
        r"^-----Original Message-----\s*$",
        r"^[-_]{8,}\s*$",
        r"^On .+ wrote:\s*$",
        r"^From:\s+.+$",
        r"^Sent:\s+.+$",
        r"^To:\s+.+$",
        r"^Subject:\s+.+$",
        r"^Von:\s+.+$",
        r"^Gesendet:\s+.+$",
        r"^An:\s+.+$",
        r"^Betreff:\s+.+$",
        r"^De\s*:\s+.+$",
        r"^Envoyé\s*:\s+.+$",
        r"^À\s*:\s+.+$",
        r"^Objet\s*:\s+.+$",
        r"^Enviado\s*:\s+.+$",
        r"^Para\s*:\s+.+$",
        r"^Asunto\s*:\s+.+$",
        r"^Inviato\s*:\s+.+$",
        r"^Oggetto\s*:\s+.+$",
        r"^Verzonden\s*:\s+.+$",
        r"^Onderwerp\s*:\s+.+$",
    ]
    pattern = re.compile("|".join(f"(?:{m})" for m in markers), re.IGNORECASE | re.MULTILINE)
    match = pattern.search(body)
    if match and match.start() > 0:
        candidate = body[:match.start()].strip()
        if len(candidate) >= 20:
            return candidate

    return body


class TranslationCache:
    """Persistent CSV cache for subject translations and full latest-message translations."""
    def __init__(self, path):
        self.path = Path(path)
        self.rows = {}
        self.load()

    def load(self):
        self.rows = {}
        if not self.path.exists():
            return
        try:
            with self.path.open("r", newline="", encoding="utf-8-sig") as handle:
                for row in csv.DictReader(handle):
                    entry_id = row.get("outlook_entry_id", "")
                    if entry_id:
                        self.rows[entry_id] = row
        except Exception:
            self.rows = {}

    def save(self):
        temp = self.path.with_suffix(".tmp")
        with temp.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=CACHE_FIELDS)
            writer.writeheader()
            writer.writerows(self.rows.values())
        temp.replace(self.path)

    def get(self, entry_id):
        return self.rows.get(entry_id)

    def has_subject(self, entry_id, original_subject):
        row = self.get(entry_id)
        return bool(row and row.get("original_subject", "") == original_subject)

    def get_subject(self, entry_id, original_subject):
        row = self.get(entry_id)
        if row and row.get("original_subject", "") == original_subject:
            return row.get("translated_subject", "")
        return ""

    def get_query(self, entry_id, query_hash):
        row = self.get(entry_id)
        if row and row.get("query_hash", "") == query_hash:
            translated = row.get("translated_query", "")
            if translated:
                return row
        return None

    def upsert(self, data):
        entry_id = data.get("outlook_entry_id", "")
        if not entry_id:
            return
        row = self.rows.get(entry_id, {field: "" for field in CACHE_FIELDS})
        for key, value in data.items():
            if key in CACHE_FIELDS:
                row[key] = str(value)
        self.rows[entry_id] = row
        self.save()


# -----------------------------------------------------------------------------
# Outlook connector
# -----------------------------------------------------------------------------

class OutlookService:
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.folder = None

    def connect(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.folder = self.find_folder_recursive(
            self.namespace.Folders, OUTLOOK_FOLDER_NAME
        )
        if self.folder is None:
            raise RuntimeError(
                f'Outlook folder "{OUTLOOK_FOLDER_NAME}" was not found.\n\n'
                "Create a folder named Multilingual in Classic Outlook and "
                "put the test vendor emails inside it."
            )

    def find_folder_recursive(self, folders, target_name):
        for index in range(1, folders.Count + 1):
            try:
                folder = folders.Item(index)
                if str(folder.Name).strip().lower() == target_name.lower():
                    return folder
                found = self.find_folder_recursive(folder.Folders, target_name)
                if found is not None:
                    return found
            except Exception:
                continue
        return None

    @staticmethod
    def read_item(item):
        return {
            "entry_id": str(getattr(item, "EntryID", "")),
            "sender": str(getattr(item, "SenderName", "")),
            "sender_email": str(getattr(item, "SenderEmailAddress", "")),
            "subject": str(getattr(item, "Subject", "")),
            "received": str(getattr(item, "ReceivedTime", "")),
            "body": str(getattr(item, "Body", "")),
            "latest_body": extract_latest_message(str(getattr(item, "Body", ""))),
        }

    def list_messages(self):
        if self.folder is None:
            self.connect()

        items = self.folder.Items
        items.Sort("[ReceivedTime]", True)
        messages = []

        for index in range(1, items.Count + 1):
            try:
                item = items.Item(index)
                # Outlook MailItem class = 43
                if getattr(item, "Class", None) != 43:
                    continue
                message = self.read_item(item)
                message["item_index"] = index
                messages.append(message)
            except Exception:
                continue

        return messages


# -----------------------------------------------------------------------------
# Main window
# -----------------------------------------------------------------------------

class HelpdeskWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1540, 980)
        self.setMinimumSize(1180, 760)

        self.outlook = OutlookService()
        self.log = LogStore(LOG_FILE)
        self.cache = TranslationCache(CACHE_FILE)
        self.current_message = None
        self.session_id = str(uuid.uuid4())
        self.session_started = time.time()
        self.query_seconds = 0.0
        self.response_seconds = 0.0
        self.warmup_done = False
        self.warmup_error = ""
        self.subject_translation_results = queue.Queue()
        self.subject_translation_running = False

        # Translation runs off the Qt GUI thread. The worker threads only call
        # Ollama and put plain Python results into these queues; all widgets,
        # CSV logging and cache writes stay on the GUI thread.
        self.query_translation_results = queue.Queue()
        self.response_translation_results = queue.Queue()
        self.query_translation_running = False
        self.response_translation_running = False
        self.query_job_token = None
        self.response_job_token = None

        self.queue_messages = []
        self.queue_translated_subjects = {}
        self.current_integrity_percent = None
        self.current_integrity_status = ""
        self.current_missing_tokens = []

        self.build_ui()
        self.apply_palette()
        self.apply_stylesheet()
        self.refresh_outlook()

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_footer)
        self.timer.timeout.connect(self._process_subject_translation_results)
        self.timer.timeout.connect(self._process_query_translation_results)
        self.timer.timeout.connect(self._process_response_translation_results)
        self.timer.start(250)

        # Warm the model in the background. This avoids blocking the UI while
        # keeping TranslateGemma resident for faster subsequent translations.
        threading.Thread(target=self._warmup_background, daemon=True).start()

    def _warmup_background(self):
        ok, error = warmup_model()
        self.warmup_done = ok
        self.warmup_error = error

    # ------------------------------------------------------------------ UI
    def build_ui(self):
        root = QWidget()
        root.setObjectName("Root")
        main = QVBoxLayout(root)
        main.setContentsMargins(14, 12, 14, 12)
        main.setSpacing(10)

        # Header
        header = QHBoxLayout()
        title_box = QVBoxLayout()
        title = QLabel(APP_TITLE)
        title.setObjectName("Title")
        subtitle = QLabel(
            "Outlook queue  •  Read the latest vendor message  •  Translate  •  "
            "Write  •  Translate back  •  Send manually"
        )
        subtitle.setObjectName("Subtitle")
        title_box.addWidget(title)
        title_box.addWidget(subtitle)
        header.addLayout(title_box)
        header.addStretch(1)

        header.addWidget(QLabel("Source Language"))
        self.language_combo = QComboBox()
        for label, _ in LANGUAGES:
            self.language_combo.addItem(label)
        self.language_combo.setMinimumWidth(145)
        header.addWidget(self.language_combo)

        self.refresh_button = QPushButton("↻  Refresh Outlook")
        self.refresh_button.clicked.connect(self.refresh_outlook)
        header.addWidget(self.refresh_button)
        main.addLayout(header)

        # Outlook-style two-pane workspace:
        # LEFT  = complete queue
        # RIGHT = vertically scrollable reading/translation pane
        workspace = QSplitter(Qt.Horizontal)
        workspace.setChildrenCollapsible(False)
        workspace.setHandleWidth(5)

        # LEFT — OUTLOOK QUEUE
        queue_group = QGroupBox("OUTLOOK QUEUE  •  PRIORITY TRIAGE")
        queue_group.setObjectName("QueuePanel")
        queue_layout = QVBoxLayout(queue_group)
        queue_layout.setContentsMargins(10, 12, 10, 10)
        queue_layout.setSpacing(7)

        queue_header = QHBoxLayout()
        queue_title = QLabel("INBOX  •  OUTLOOK VIEW")
        queue_title.setObjectName("PaneTitle")
        queue_header.addWidget(queue_title)
        queue_header.addStretch(1)
        queue_layout.addLayout(queue_header)

        queue_metrics = QHBoxLayout()
        queue_metrics.setSpacing(6)
        self.total_chip = QLabel("  ●  TOTAL  —")
        self.total_chip.setObjectName("QueueChip")
        self.priority_chip = QLabel("  ●  PRIORITY  —")
        self.priority_chip.setObjectName("PriorityChip")
        self.ready_chip = QLabel("  ●  SUBJECTS  —")
        self.ready_chip.setObjectName("ReadyChip")
        queue_metrics.addWidget(self.total_chip)
        queue_metrics.addWidget(self.priority_chip)
        queue_metrics.addWidget(self.ready_chip)
        queue_metrics.addStretch(1)
        queue_layout.addLayout(queue_metrics)

        self.email_table = QTableWidget(0, 4)
        self.email_table.setHorizontalHeaderLabels([
            "Mail", "Language", "Priority", "Status"
        ])
        self.email_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.email_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.email_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.email_table.setAlternatingRowColors(True)
        self.email_table.setWordWrap(True)
        self.email_table.setShowGrid(False)
        self.email_table.verticalHeader().setVisible(True)
        self.email_table.verticalHeader().setDefaultSectionSize(82)
        self.email_table.verticalHeader().setMinimumSectionSize(82)
        self.email_table.horizontalHeader().setDefaultAlignment(
            Qt.AlignLeft | Qt.AlignVCenter
        )
        self.email_table.setSizeAdjustPolicy(QAbstractScrollArea.AdjustIgnored)
        self.email_table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.email_table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.email_table.verticalScrollBar().setSingleStep(12)
        header_view = self.email_table.horizontalHeader()
        header_view.setSectionResizeMode(0, QHeaderView.Stretch)
        header_view.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header_view.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header_view.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.email_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.email_table.cellDoubleClicked.connect(
            lambda *_: self.load_selected_email()
        )
        queue_layout.addWidget(self.email_table, 1)

        queue_actions = QHBoxLayout()
        queue_actions.setContentsMargins(0, 4, 0, 0)
        queue_actions.setSpacing(7)
        self.load_button = QPushButton("↳  Load Selected")
        self.load_button.clicked.connect(self.load_selected_email)
        queue_actions.addWidget(self.load_button)
        self.manual_button = QPushButton("+  New Manual")
        self.manual_button.clicked.connect(self.new_manual_session)
        queue_actions.addWidget(self.manual_button)
        queue_layout.addLayout(queue_actions)

        self.queue_info = QLabel("Loading Outlook queue…")
        self.queue_info.setObjectName("QueueInfo")
        self.queue_info.setWordWrap(True)
        queue_layout.addWidget(self.queue_info)
        workspace.addWidget(queue_group)

        # RIGHT — OUTLOOK-STYLE READING / TRANSLATION PANE
        reading_scroll = QScrollArea()
        reading_scroll.setObjectName("ReadingScroll")
        reading_scroll.setWidgetResizable(True)
        reading_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        reading_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        reading_root = QWidget()
        reading_root.setObjectName("ReadingRoot")
        reading_layout = QVBoxLayout(reading_root)
        reading_layout.setContentsMargins(4, 4, 12, 4)
        reading_layout.setSpacing(10)

        reading_header = QFrame()
        reading_header.setObjectName("ReadingHeader")
        reading_header_layout = QVBoxLayout(reading_header)
        reading_header_layout.setContentsMargins(14, 10, 14, 10)
        reading_title = QLabel("READING & TRANSLATION")
        reading_title.setObjectName("ReadingTitle")
        reading_header_layout.addWidget(reading_title)
        reading_subtitle = QLabel(
            "Select an email from the queue. The current/latest vendor message "
            "is shown here; older Outlook trail text is not translated automatically."
        )
        reading_subtitle.setObjectName("ReadingSubtitle")
        reading_subtitle.setWordWrap(True)
        reading_header_layout.addWidget(reading_subtitle)
        reading_layout.addWidget(reading_header)

        # 1 — QUERY RECEIVED
        original_group = QGroupBox("1  |  QUERY RECEIVED")
        original_group.setObjectName("ReadingCard")
        original_layout = QVBoxLayout(original_group)
        original_layout.setContentsMargins(12, 12, 12, 12)
        original_layout.setSpacing(7)
        original_layout.addWidget(QLabel("Original Subject"))
        self.original_subject = QLineEdit()
        self.original_subject.setPlaceholderText(
            "Select an Outlook email or paste/type the subject manually."
        )
        original_layout.addWidget(self.original_subject)
        original_layout.addWidget(QLabel("Query Received"))
        self.query_text = QTextEdit()
        self.query_text.setMinimumHeight(190)
        self.query_text.setPlaceholderText(
            "Current/newest vendor message only. Older email trail is intentionally not translated."
        )
        original_layout.addWidget(self.query_text)
        latest_note = QLabel(
            "⚡ Latest message only  •  older trail is excluded automatically  •  "
            "paste older text manually if needed"
        )
        latest_note.setObjectName("LatestNote")
        latest_note.setWordWrap(True)
        original_layout.addWidget(latest_note)
        self.translate_query_button = QPushButton("Translate Query  →  English")
        self.translate_query_button.clicked.connect(self.translate_query)
        original_layout.addWidget(self.translate_query_button)
        reading_layout.addWidget(original_group)

        # 2 — ENGLISH TRANSLATION
        english_group = QGroupBox("2  |  ENGLISH TRANSLATION")
        english_group.setObjectName("ReadingCard")
        english_layout = QVBoxLayout(english_group)
        english_layout.setContentsMargins(12, 12, 12, 12)
        english_layout.setSpacing(7)
        english_layout.addWidget(QLabel("Translated Subject"))
        self.english_subject = QLineEdit()
        self.english_subject.setReadOnly(True)
        english_layout.addWidget(self.english_subject)
        english_layout.addWidget(QLabel("English Query"))
        self.english_query = QTextEdit()
        self.english_query.setReadOnly(True)
        self.english_query.setMinimumHeight(190)
        english_layout.addWidget(self.english_query)
        self.query_time_label = QLabel("Translation time: —")
        english_layout.addWidget(self.query_time_label)
        reading_layout.addWidget(english_group)

        # 3 — USER RESPONSE
        response_group = QGroupBox("3  |  YOUR RESPONSE — MANUAL ONLY")
        response_group.setObjectName("ReadingCard")
        response_layout = QVBoxLayout(response_group)
        response_layout.setContentsMargins(12, 12, 12, 12)
        response_layout.setSpacing(7)
        response_layout.addWidget(QLabel("Response Subject (English)"))
        self.response_subject = QLineEdit()
        self.response_subject.setPlaceholderText("Edit if required.")
        response_layout.addWidget(self.response_subject)
        response_layout.addWidget(QLabel("Your Response"))
        self.response_text = QTextEdit()
        self.response_text.setMinimumHeight(190)
        self.response_text.setPlaceholderText(
            "Type/paste the approved English response here. No AI reply generation is used."
        )
        response_layout.addWidget(self.response_text)
        self.translate_response_button = QPushButton(
            "Translate Response  →  Vendor Language"
        )
        self.translate_response_button.clicked.connect(self.translate_response)
        response_layout.addWidget(self.translate_response_button)
        reading_layout.addWidget(response_group)

        # 4 — FINAL RESPONSE
        final_group = QGroupBox("4  |  FINAL RESPONSE — ORIGINAL VENDOR LANGUAGE")
        final_group.setObjectName("ReadingCard")
        final_layout = QVBoxLayout(final_group)
        final_layout.setContentsMargins(12, 12, 12, 12)
        final_layout.setSpacing(7)
        final_layout.addWidget(QLabel("Final Subject"))
        self.final_subject = QLineEdit()
        self.final_subject.setReadOnly(True)
        final_layout.addWidget(self.final_subject)
        final_layout.addWidget(QLabel("Final Response"))
        self.final_response = QTextEdit()
        self.final_response.setReadOnly(True)
        self.final_response.setMinimumHeight(210)
        final_layout.addWidget(self.final_response)

        final_actions = QHBoxLayout()
        self.copy_subject_button = QPushButton("Copy Subject")
        self.copy_subject_button.clicked.connect(self.copy_subject)
        self.copy_response_button = QPushButton("Copy Response")
        self.copy_response_button.clicked.connect(self.copy_response)
        self.mark_complete_button = QPushButton("✓  Mark Completed")
        self.mark_complete_button.clicked.connect(self.mark_completed)
        final_actions.addWidget(self.copy_subject_button)
        final_actions.addWidget(self.copy_response_button)
        final_actions.addStretch(1)
        final_actions.addWidget(self.mark_complete_button)
        final_layout.addLayout(final_actions)

        self.response_time_label = QLabel("Translation time: —")
        final_layout.addWidget(self.response_time_label)
        self.integrity_label = QLabel("Translation Integrity: —")
        self.integrity_label.setObjectName("IntegrityLabel")
        self.integrity_label.setToolTip(
            "Deterministic preservation check for invoice/PO/reference IDs, "
            "dates, amounts/currencies, email addresses and URLs. "
            "Not model confidence."
        )
        final_layout.addWidget(self.integrity_label)

        reading_layout.addWidget(final_group)
        reading_layout.addStretch(1)

        reading_scroll.setWidget(reading_root)
        workspace.addWidget(reading_scroll)

        workspace.setSizes([600, 900])
        workspace.setStretchFactor(0, 4)
        workspace.setStretchFactor(1, 6)
        main.addWidget(workspace, 1)

        self.setCentralWidget(root)
        self.setMinimumSize(1250, 820)
        self.resize(1500, 950)
        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage(
            f"Ready • Model: {MODEL} • Log: {LOG_FILE.name}"
        )

    def apply_palette(self):
        # Explicit palette prevents Windows/Qt theme combinations from making
        # white text appear on white backgrounds.
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor("#F4F6F8"))
        palette.setColor(QPalette.WindowText, QColor("#18212B"))
        palette.setColor(QPalette.Base, QColor("#FFFFFF"))
        palette.setColor(QPalette.AlternateBase, QColor("#F7F9FB"))
        palette.setColor(QPalette.Text, QColor("#18212B"))
        palette.setColor(QPalette.Button, QColor("#E9EEF3"))
        palette.setColor(QPalette.ButtonText, QColor("#18212B"))
        palette.setColor(QPalette.Highlight, QColor("#2D6CDF"))
        palette.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
        palette.setColor(QPalette.PlaceholderText, QColor("#6B7785"))
        palette.setColor(QPalette.ToolTipBase, QColor("#FFFFFF"))
        palette.setColor(QPalette.ToolTipText, QColor("#18212B"))
        QApplication.instance().setPalette(palette)

    def apply_stylesheet(self):
        self.setStyleSheet("""
            QWidget#Root {
                background: #F3F5F7;
                color: #17212B;
            }
            QMainWindow {
                background: #F3F5F7;
            }
            QScrollArea#ReadingScroll {
                background: #F3F5F7;
                border: none;
            }
            QWidget#ReadingRoot {
                background: #F3F5F7;
            }
            QFrame#ReadingHeader {
                background: #FFFFFF;
                border: 1px solid #D8A3A3;
                border-radius: 11px;
            }
            QLabel#ReadingTitle {
                color: #8B0000;
                font-size: 14px;
                font-weight: 850;
            }
            QLabel#ReadingSubtitle {
                color: #66717D;
                font-size: 10px;
            }
            QLabel#PaneTitle {
                color: #8B0000;
                font-size: 12px;
                font-weight: 850;
            }
            QGroupBox#ReadingCard {
                border: 1px solid #D8A3A3;
                border-radius: 11px;
                background: #FFFFFF;
            }
            QGroupBox#ReadingCard::title {
                color: #8B0000;
            }
            QSplitter::handle:horizontal {
                background: #B22222;
                width: 5px;
                margin: 3px 0;
                border-radius: 2px;
            }

            QLabel {
                color: #18212B;
            }
            QLabel#Title {
                color: #8B0000;
                font-size: 24px;
                font-weight: 800;
                padding: 2px 0 3px 0;
                letter-spacing: 0.2px;
            }
            QLabel#Subtitle {
                color: #66717D;
                font-size: 10px;
                padding-top: 1px;
            }
            QLabel#QueueChip, QLabel#PriorityChip, QLabel#ReadyChip {
                background: #FFFFFF;
                border: 1px solid #E3B6BA;
                border-radius: 10px;
                color: #6E1820;
                padding: 6px 10px;
                font-size: 10px;
                font-weight: 800;
            }
            QLabel#PriorityChip {
                background: #FFF1F2;
                border-color: #C93B45;
                color: #8B0000;
            }
            QLabel#ReadyChip {
                background: #F5FAF7;
                border-color: #B9D8C3;
                color: #27643A;
            }
            QLabel#QueueInfo {
                color: #44515F;
                font-weight: 600;
            }
            QGroupBox {
                background: #FFFFFF;
                color: #18212B;
                border: 1px solid #D6DDE4;
                border-radius: 11px;
                margin-top: 10px;
                padding: 10px;
                font-weight: 750;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 14px;
                padding: 0 8px;
                color: #8B0000;
                background: #FFFFFF;
            }
            QLineEdit, QTextEdit, QComboBox {
                background: #FFFFFF;
                color: #18212B;
                border: 1px solid #C5CFDA;
                border-radius: 5px;
                selection-background-color: #2D6CDF;
                selection-color: #FFFFFF;
            }
            QLineEdit {
                padding: 7px;
            }
            QTextEdit {
                padding: 6px;
            }
            QComboBox {
                padding: 6px 8px;
            }
            QComboBox QAbstractItemView {
                background: #FFFFFF;
                color: #18212B;
                selection-background-color: #2D6CDF;
                selection-color: #FFFFFF;
            }
            QPushButton {
                background: #FFFFFF;
                color: #252E38;
                border: 1px solid #C9D1DA;
                border-radius: 8px;
                padding: 9px 14px;
                font-weight: 700;
            }
            QPushButton:hover {
                background: #FFF4F4;
                border-color: #A7192E;
                color: #7F0712;
            }
            QPushButton:pressed {
                background: #F9DDE0;
            }
            QPushButton:disabled {
                background: #E3E7EB;
                color: #7B8794;
                border-color: #D2D8DE;
            }
            QTableWidget {
                background: #FFFFFF;
                color: #18212B;
                alternate-background-color: #FAFBFC;
                border: 1px solid #D4DCE4;
                border-radius: 8px;
                gridline-color: transparent;
                selection-background-color: #FBE1E3;
                selection-color: #18212B;
            }
            QTableWidget::item {
                padding: 7px 8px;
                color: #18212B;
            }
            QTableWidget::item:selected {
                background: #FBE1E3;
                color: #18212B;
            }
            QTableWidget::item:focus {
                outline: none;
            }
            QHeaderView::section {
                background: #E9EEF3;
                color: #18212B;
                border: 0;
                border-bottom: 1px solid #C5CFDA;
                padding: 7px;
                font-weight: 700;
            }
            QStatusBar {
                background: #E9EEF3;
                color: #18212B;
                border-top: 1px solid #C5CFDA;
            }
            QSplitter::handle {
                background: #C5CFDA;
            }
            QLabel#Title {
                color: #8B0000;
            }
            QLabel#LatestNote {
                color: #8B0000;
                background: #FFF5F5;
                border: 1px solid #E8B8BC;
                border-radius: 6px;
                padding: 5px 8px;
                font-size: 9px;
                font-weight: 700;
            }
            QLabel#IntegrityLabel {
                color: #8B0000;
                font-size: 11px;
                font-weight: 800;
                padding: 4px 0;
            }
            QGroupBox {
                border: 1px solid #D8A3A3;
            }
            QGroupBox::title {
                color: #8B0000;
            }
            QPushButton {
                border: 1px solid #B22222;
            }
            QPushButton:hover {
                background: #FCE7E7;
            }
            QTableWidget {
                border: 1px solid #B22222;
            }
            QHeaderView::section {
                border-bottom: 2px solid #B22222;
            }
            QScrollBar:vertical {
                background: #F7E6E6;
                width: 16px;
                margin: 1px;
            }
            QScrollBar::handle:vertical {
                background: #9B1C1C;
                min-height: 45px;
                border-radius: 7px;
            }
            QScrollBar::handle:vertical:hover {
                background: #6F0000;
            }
        """)

    # --------------------------------------------------------------- Outlook
    def refresh_outlook(self):
        self.refresh_button.setEnabled(False)
        self.queue_info.setText("Connecting to Classic Outlook…")
        QApplication.processEvents()
        try:
            messages = self.outlook.list_messages()
        except Exception as exc:
            self.email_table.setRowCount(0)
            self.queue_info.setText("Outlook connection failed.")
            QMessageBox.warning(self, "Outlook Connection", str(exc))
            return
        finally:
            self.refresh_button.setEnabled(True)

        self.queue_messages = list(messages)
        self.queue_translated_subjects = {}
        self.email_table.setRowCount(0)

        # Restore cached subject translations first. No model call is made for
        # previously seen messages.
        new_messages = []
        for message in self.queue_messages:
            entry_id = message.get("entry_id", "")
            subject = message.get("subject", "")
            cached_subject = self.cache.get_subject(entry_id, subject)
            if cached_subject:
                self.queue_translated_subjects[entry_id] = cached_subject
            else:
                new_messages.append(message)

        self.queue_messages.sort(
            key=lambda m: 0 if priority_level(
                m.get("subject", ""),
                self.queue_translated_subjects.get(m.get("entry_id", ""), "")
            ) != "NORMAL" else 1
        )

        for message in self.queue_messages:
            has_cached = message.get("entry_id", "") in self.queue_translated_subjects
            self._append_queue_row(message, subject_pending=not has_cached)

        total = len(messages)
        cached_count = total - len(new_messages)
        priority_count = sum(
            priority_level(
                m.get("subject", ""),
                self.queue_translated_subjects.get(m.get("entry_id", ""), "")
            ) != "NORMAL"
            for m in messages
        )
        self.total_chip.setText(f"  ●  TOTAL  {total:02d}")
        self.priority_chip.setText(f"  ●  PRIORITY  {priority_count:02d}")
        self.ready_chip.setText(f"  ●  CACHED  {cached_count:02d}/{total:02d}")

        if not new_messages:
            self.queue_info.setText(
                f"{total} email(s) loaded • cached subjects reused • no retranslation needed"
            )
            self.statusBar().showMessage(
                "Queue refreshed. Previously translated subjects were loaded from CSV; no duplicate AI translation was run."
            )
            return

        # Explicit user confirmation: never silently translate the whole inbox.
        answer = QMessageBox.question(
            self,
            "New Emails Detected",
            f"Outlook contains {total} email(s).\n\n"
            f"{len(new_messages)} email(s) are new and do not have a cached subject translation.\n"
            f"{cached_count} previously translated email(s) will be reused from CSV.\n\n"
            "Translate SUBJECTS for the NEW emails only?\n"
            "Older emails will NOT be translated again.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )

        if answer == QMessageBox.StandardButton.Yes:
            self.queue_info.setText(
                f"{len(new_messages)} new email(s) • translating NEW subjects only…"
            )
            self.statusBar().showMessage(
                f"Translating subjects for {len(new_messages)} new email(s) only."
            )
            self._start_subject_translation(new_messages)
        else:
            self.queue_info.setText(
                f"{total} email(s) loaded • {cached_count} cached • new subjects skipped"
            )
            self.statusBar().showMessage(
                "New subject translation skipped. Refresh later to translate new emails."
            )

    def _append_queue_row(self, message, subject_pending=False):
        row = self.email_table.rowCount()
        self.email_table.insertRow(row)

        detected_code = detect_language_code(
            message.get("latest_body", "") or message.get("body", "")
        )
        detected_name = language_name(detected_code) if detected_code else "Unknown"
        log_row = self.log.find(entry_id=message["entry_id"])
        status = log_row.get("status", "NEW") if log_row else "NEW"

        original = message.get("subject", "") or "(No subject)"
        translated = self.queue_translated_subjects.get(
            message.get("entry_id", ""), ""
        )
        priority = priority_level(original, translated)

        # Outlook-style compact mail preview:
        #   Line 1 = sender + received time
        #   Line 2 = original subject + English subject
        #   Line 3 = first lines of the latest message
        sender = (
            message.get("sender", "")
            or message.get("sender_email", "")
            or "(Unknown)"
        )
        received = str(message.get("received", "") or "")
        received_display = format_received_time(received)

        latest = (
            message.get("latest_body", "")
            or extract_latest_message(message.get("body", ""))
            or ""
        )
        preview_lines = [
            re.sub(r"\s+", " ", line).strip()
            for line in latest.splitlines()
            if line.strip()
        ]
        preview = " ".join(preview_lines)[:155]
        if len(" ".join(preview_lines)) > 155:
            preview += "…"

        if translated:
            subject_line = f"{original}  •  EN: {translated}"
        elif subject_pending:
            subject_line = f"{original}  •  EN: Translating…"
        else:
            subject_line = original

        mail_text = f"{sender}   •   {received_display}\n{subject_line}"
        if preview:
            mail_text += f"\n{preview}"

        mail_item = QTableWidgetItem(mail_text)
        mail_item.setToolTip(
            f"From: {sender}\n"
            f"Received: {received}\n"
            f"Original: {original}\n"
            f"English: {translated or 'Not translated yet'}\n\n"
            f"Preview: {preview}"
        )
        mail_item.setData(Qt.UserRole, message)
        mail_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.email_table.setItem(row, 0, mail_item)

        language_item = QTableWidgetItem(detected_name)
        language_item.setTextAlignment(Qt.AlignCenter)
        self.email_table.setItem(row, 1, language_item)

        priority_text = (
            "🔴 URGENT / IMPORTANT" if priority != "NORMAL" else "Normal"
        )
        priority_item = QTableWidgetItem(priority_text)
        priority_item.setTextAlignment(Qt.AlignCenter)
        self.email_table.setItem(row, 2, priority_item)

        status_item = QTableWidgetItem(status)
        status_item.setTextAlignment(Qt.AlignCenter)
        self.email_table.setItem(row, 3, status_item)

        if priority != "NORMAL":
            priority_item.setForeground(QColor("#8B0000"))
            priority_item.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
            for column in range(self.email_table.columnCount()):
                item = self.email_table.item(row, column)
                if item:
                    item.setBackground(QColor("#FFF0F0"))

        if status == "COMPLETED":
            status_item.setForeground(QColor("#1B7F3A"))
        elif status == "RESPONSE_READY":
            status_item.setForeground(QColor("#B05A00"))
        elif status == "FAILED":
            status_item.setForeground(QColor("#B42318"))
        else:
            status_item.setForeground(QColor("#334155"))

    def _start_subject_translation(self, messages=None):
        if self.subject_translation_running:
            return
        messages = list(messages or [])
        if not messages:
            return
        self.subject_translation_running = True

        def worker():
            for message in messages:
                entry_id = message.get("entry_id", "")
                subject = message.get("subject", "")
                try:
                    body_language_code = detect_language_code(message.get("latest_body", "") or message.get("body", ""))
                    body_language_name = (
                        language_name(body_language_code)
                        if body_language_code else "English"
                    )
                    if body_language_name == "English" or not subject.strip():
                        translated = subject
                    else:
                        translated = translate_subject_only(subject, "English")
                    self.cache.upsert({
                        "outlook_entry_id": entry_id,
                        "source_language": body_language_name,
                        "original_subject": subject,
                        "translated_subject": translated or subject,
                        "translated_at": now_string(),
                    })
                    self.subject_translation_results.put(
                        ("OK", entry_id, subject, translated)
                    )
                except Exception as exc:
                    self.subject_translation_results.put(
                        ("ERR", entry_id, subject, "", str(exc))
                    )
            self.subject_translation_results.put(("DONE",))

        threading.Thread(target=worker, daemon=True).start()

    def _process_subject_translation_results(self):
        changed = False
        done = False

        while True:
            try:
                result = self.subject_translation_results.get_nowait()
            except queue.Empty:
                break

            if result[0] == "DONE":
                done = True
                continue

            status, entry_id, original, translated = result[:4]
            if status != "OK":
                continue

            self.queue_translated_subjects[entry_id] = translated or original
            changed = True

            for row in range(self.email_table.rowCount()):
                row_item = self.email_table.item(row, 0)
                message = row_item.data(Qt.UserRole) if row_item else None
                if not message or message.get("entry_id") != entry_id:
                    continue

                priority = priority_level(original, translated)
                subject_item = self.email_table.item(row, 0)
                if subject_item:
                    message = subject_item.data(Qt.UserRole) or {}
                    sender = (
                        message.get("sender", "")
                        or message.get("sender_email", "")
                        or "(Unknown)"
                    )
                    received = str(message.get("received", "") or "")
                    received_display = format_received_time(received)
                    latest = (
                        message.get("latest_body", "")
                        or extract_latest_message(message.get("body", ""))
                        or ""
                    )
                    preview_lines = [
                        re.sub(r"\s+", " ", line).strip()
                        for line in latest.splitlines()
                        if line.strip()
                    ]
                    preview = " ".join(preview_lines)[:155]
                    if len(" ".join(preview_lines)) > 155:
                        preview += "…"
                    subject_item.setText(
                        f"{sender}   •   {received_display}\n"
                        f"{original}  •  EN: {translated or '(unavailable)'}"
                        + (f"\n{preview}" if preview else "")
                    )
                    subject_item.setToolTip(
                        f"Original: {original}\nEnglish: {translated or '(unavailable)'}"
                    )

                priority_item = self.email_table.item(row, 2)
                if priority != "NORMAL":
                    priority_item.setText("🔴 URGENT / IMPORTANT")
                    priority_item.setForeground(QColor("#8B0000"))
                    priority_item.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
                    for column in range(self.email_table.columnCount()):
                        self.email_table.item(row, column).setBackground(QColor("#FFF0F0"))
                else:
                    priority_item.setText("Normal")
                    priority_item.setForeground(QColor("#334155"))
                break

        if changed:
            self._reorder_queue_by_priority()

        if done:
            self.subject_translation_running = False
            priority_count = sum(
                priority_level(
                    m.get("subject", ""),
                    self.queue_translated_subjects.get(m.get("entry_id", ""), "")
                ) != "NORMAL"
                for m in self.queue_messages
            )
            self.total_chip.setText(f"  ●  TOTAL  {len(self.queue_messages):02d}")
            self.priority_chip.setText(f"  ●  PRIORITY  {priority_count:02d}")
            self.ready_chip.setText(
                f"  ●  SUBJECTS  {len(self.queue_translated_subjects):02d}/{len(self.queue_messages):02d}"
            )
            self.queue_info.setText(
                f"{len(self.queue_messages)} email(s) • subject triage complete • "
                f"urgent/important items highlighted first"
            )
            self.statusBar().showMessage("Subject triage completed.")

    def _reorder_queue_by_priority(self):
        ordered = []
        for row in range(self.email_table.rowCount()):
            item = self.email_table.item(row, 0)
            message = item.data(Qt.UserRole) if item else None
            if not message:
                continue
            translated = self.queue_translated_subjects.get(message.get("entry_id", ""), "")
            priority = priority_level(message.get("subject", ""), translated)
            ordered.append((0 if priority != "NORMAL" else 1, message))

        ordered.sort(key=lambda x: x[0])
        current_ids = []
        for row in range(self.email_table.rowCount()):
            item = self.email_table.item(row, 0)
            msg = item.data(Qt.UserRole) if item else None
            if msg:
                current_ids.append(msg.get("entry_id", ""))

        new_ids = [message.get("entry_id", "") for _, message in ordered]
        if current_ids == new_ids:
            return

        selected_entry = ""
        current = self.email_table.currentRow()
        if current >= 0 and self.email_table.item(current, 0):
            selected = self.email_table.item(current, 0).data(Qt.UserRole)
            if selected:
                selected_entry = selected.get("entry_id", "")

        self.email_table.setUpdatesEnabled(False)
        self.email_table.setRowCount(0)
        for _, message in ordered:
            self._append_queue_row(message, subject_pending=False)
        self.email_table.setUpdatesEnabled(True)

        if selected_entry:
            for row in range(self.email_table.rowCount()):
                item = self.email_table.item(row, 0)
                msg = item.data(Qt.UserRole) if item else None
                if msg and msg.get("entry_id") == selected_entry:
                    self.email_table.selectRow(row)
                    break

    def load_selected_email(self):
        row = self.email_table.currentRow()
        if row < 0:
            QMessageBox.information(
                self,
                "Select Email",
                "Select an email from the Outlook queue first.",
            )
            return

        item = self.email_table.item(row, 0)
        message = item.data(Qt.UserRole) if item else None
        if not message:
            QMessageBox.warning(self, "Load Error", "Could not read the selected Outlook message.")
            return

        self.current_message = message
        self.start_new_session()

        latest_body = message.get("latest_body") or extract_latest_message(message.get("body", ""))
        message["latest_body"] = latest_body
        self.original_subject.setText(message["subject"])
        self.query_text.setPlainText(latest_body)
        self.clear_translation_and_response()

        # Reuse a prior full translation if this exact latest-message content
        # was already translated. This avoids another 30–40s model call.
        detected_code = detect_language_code(latest_body)
        detected_name_for_cache = language_name(detected_code) if detected_code else ""
        cached_query = self.cache.get_query(
            message.get("entry_id", ""),
            text_hash(message.get("subject", ""), latest_body),
        )
        cache_hit = bool(cached_query)
        if cached_query:
            cached_subject = cached_query.get("translated_subject", "") or message["subject"]
            cached_body = cached_query.get("translated_query", "")
            self.english_subject.setText(cached_subject)
            self.english_query.setPlainText(cached_body)
            self.response_subject.setText(cached_subject)
            try:
                self.current_integrity_percent = int(float(cached_query.get("integrity_percent", "")))
            except Exception:
                self.current_integrity_percent = None
            self.current_integrity_status = cached_query.get("integrity_status", "")
            self.integrity_label.setText(
                f"Translation Integrity: {self.current_integrity_percent}%  •  CACHED"
                if self.current_integrity_percent is not None
                else "Translation Integrity: Cached translation"
            )
            self.integrity_label.setStyleSheet("color: #27643A; font-weight: 800;")
            self.query_time_label.setText("Translation time: 0.00 seconds • cached")
            self.queue_info.setText(
                "Loaded latest message only • previous full translation restored from CSV cache"
            )
            self.statusBar().showMessage(
                "Cached translation restored instantly. Older email trail was not translated."
            )

        detected = detect_language_code(latest_body)
        detected_name = language_name(detected) if detected else ""
        index = self.language_combo.findText(detected_name)
        self.language_combo.setCurrentIndex(index if index >= 0 else 0)

        if cache_hit:
            self.queue_info.setText(
                "Loaded latest message only • cached translation restored • no AI call needed"
            )
            self.statusBar().showMessage(
                "Latest message loaded. Cached translation restored; older trail was not translated."
            )
        else:
            self.queue_info.setText(
                f"Loaded latest message: {message['subject'] or '(No subject)'}"
            )
            self.statusBar().showMessage(
                "Latest message loaded. Review/edit the Query Received box, then translate."
            )

    # ------------------------------------------------------------- Sessions
    def start_new_session(self):
        # Changing email/session invalidates any result still running in a
        # background worker. The worker may finish, but its result is ignored.
        self.session_id = str(uuid.uuid4())
        self.session_started = time.time()
        self.query_seconds = 0.0
        self.response_seconds = 0.0
        self.current_integrity_percent = None
        self.current_integrity_status = ""
        self.current_missing_tokens = []
        self.query_job_token = None
        self.response_job_token = None
        self.query_translation_running = False
        self.response_translation_running = False
        if hasattr(self, "translate_query_button"):
            self.translate_query_button.setEnabled(True)
        if hasattr(self, "translate_response_button"):
            self.translate_response_button.setEnabled(True)

    def clear_translation_and_response(self):
        self.english_subject.clear()
        self.english_query.clear()
        self.response_subject.clear()
        self.response_text.clear()
        self.final_subject.clear()
        self.final_response.clear()
        self.query_time_label.setText("Translation time: —")
        self.response_time_label.setText("Translation time: —")
        self.integrity_label.setText("Translation Integrity: —")
        self.integrity_label.setStyleSheet("color: #8B0000; font-weight: 800;")
        self.integrity_label.setToolTip(
            "Deterministic preservation check for critical business values. Not model confidence."
        )

    def new_manual_session(self):
        self.current_message = None
        self.start_new_session()
        self.original_subject.clear()
        self.query_text.clear()
        self.clear_translation_and_response()
        self.language_combo.setCurrentIndex(0)
        self.queue_info.setText("Manual session started — paste/type the vendor query.")
        self.statusBar().showMessage("Manual session ready.")

    # ------------------------------------------------------------ Languages
    def selected_source_language(self):
        selected = self.language_combo.currentText()
        if selected != "Auto Detect":
            return selected

        code = detect_language_code(self.query_text.toPlainText())
        return language_name(code) if code else ""

    # --------------------------------------------------------------- Logging
    def base_log_data(self):
        message = self.current_message or {}
        entry_id = message.get("entry_id", "")
        source_language = self.selected_source_language()
        return {
            "session_id": self.session_id,
            "outlook_entry_id": entry_id,
            "source": "Outlook" if entry_id else "Manual",
            "sender": message.get("sender", ""),
            "received_time": message.get("received", ""),
            "source_language": source_language,
            "response_language": source_language,
            "original_subject": self.original_subject.text().strip(),
            "translated_subject": self.english_subject.text().strip(),
            "priority": priority_level(
                self.original_subject.text().strip(),
                self.english_subject.text().strip(),
            ),
            "translation_integrity_percent": (
                str(self.current_integrity_percent)
                if self.current_integrity_percent is not None else ""
            ),
            "integrity_status": self.current_integrity_status,
            "query_word_count": word_count(self.query_text.toPlainText()),
            "total_elapsed_seconds": f"{time.time() - self.session_started:.2f}",
        }

    # ---------------------------------------------------------- Translation
    def translate_query(self):
        subject = self.original_subject.text().strip()
        body = extract_latest_message(self.query_text.toPlainText().strip())
        self.query_text.setPlainText(body)

        if not body:
            QMessageBox.warning(
                self,
                "Missing Query",
                "Enter or paste the vendor query first.",
            )
            return

        source_language = self.selected_source_language()
        if not source_language:
            QMessageBox.warning(
                self,
                "Language Detection",
                "Could not detect the source language. Select it manually from the Source Language box.",
            )
            return

        # Exact-content cache check: reopening/retranslating the same email should
        # not invoke TranslateGemma a second time.
        cached_query = self.cache.get_query(
            (self.current_message or {}).get("entry_id", ""),
            text_hash(subject, body),
        )
        if cached_query and cached_query.get("translated_query", ""):
            final_subject = cached_query.get("translated_subject", "") or subject
            final_body = cached_query.get("translated_query", "")
            self.english_subject.setText(final_subject)
            self.english_query.setPlainText(final_body)
            self.response_subject.setText(final_subject)
            try:
                score = int(float(cached_query.get("integrity_percent", "")))
            except Exception:
                score = None
            self.current_integrity_percent = score
            self.current_integrity_status = cached_query.get("integrity_status", "")
            self.query_seconds = 0.0
            self.query_time_label.setText("Translation time: 0.00 seconds • cached")
            self.integrity_label.setText(
                f"Translation Integrity: {score}%  •  CACHED"
                if score is not None else "Translation Integrity: Cached translation"
            )
            self.integrity_label.setStyleSheet("color: #27643A; font-weight: 800;")
            self.statusBar().showMessage(
                "Cached translation reused — no AI call was required."
            )
            return

        # Invalidate any older query job and start a new worker. The GUI thread
        # never waits on Ollama, so Windows can continue repainting/switching apps.
        token = str(uuid.uuid4())
        session_token = self.session_id
        self.query_job_token = token
        self.query_translation_running = True
        self.translate_query_button.setEnabled(False)
        self.query_time_label.setText("Translation running in background…")
        self.statusBar().showMessage(
            "Translating in background — the helpdesk window remains responsive."
        )
        started = time.perf_counter()

        def worker():
            try:
                if source_language.lower() == "english":
                    translated_subject = subject
                    translated_body = body
                else:
                    translated_subject, translated_body = translate_text(
                        subject, body, "English"
                    )

                elapsed = time.perf_counter() - started
                self.query_translation_results.put((
                    "OK",
                    token,
                    session_token,
                    subject,
                    body,
                    source_language,
                    translated_subject or subject,
                    translated_body or body,
                    elapsed,
                ))
            except Exception as exc:
                self.query_translation_results.put((
                    "ERR",
                    token,
                    session_token,
                    subject,
                    body,
                    source_language,
                    str(exc),
                    "",
                    time.perf_counter() - started,
                ))

        threading.Thread(
            target=worker,
            name="AxaltaQueryTranslation",
            daemon=True,
        ).start()

    def _process_query_translation_results(self):
        while True:
            try:
                result = self.query_translation_results.get_nowait()
            except queue.Empty:
                return

            status, token, session_token, subject, body, source_language, value, translated_body, elapsed = result

            # Ignore a completed worker if the user has already moved to another
            # email/manual session. This prevents stale translations overwriting
            # the new reading pane.
            if token != self.query_job_token or session_token != self.session_id:
                continue

            self.query_translation_running = False
            self.translate_query_button.setEnabled(True)

            if status == "ERR":
                self.query_time_label.setText(
                    f"Translation failed after {elapsed:.2f} seconds"
                )
                self.log.upsert({**self.base_log_data(), "status": "FAILED"})
                QMessageBox.critical(
                    self,
                    "Query Translation Error",
                    "Could not translate the query.\n\n"
                    f"{value}\n\n"
                    f"Check that Ollama is running and '{MODEL}' is installed.",
                )
                return

            translated_subject = value
            final_body = translated_body
            self.query_seconds = elapsed
            self.english_subject.setText(translated_subject)
            self.english_query.setPlainText(final_body)
            self.response_subject.setText(translated_subject)

            score, missing, integrity_status = translation_integrity(
                f"{subject}\n{body}",
                f"{translated_subject}\n{final_body}",
            )
            self.current_integrity_percent = score
            self.current_integrity_status = integrity_status
            self.current_missing_tokens = missing

            if integrity_status == "VERIFIED":
                self.integrity_label.setText(
                    f"Translation Integrity: {score}%  •  VERIFIED"
                )
                self.integrity_label.setStyleSheet(
                    "color: #137333; font-weight: 800;"
                )
            else:
                self.integrity_label.setText(
                    f"Translation Integrity: {score}%  •  REVIEW REQUIRED"
                )
                self.integrity_label.setStyleSheet(
                    "color: #8B0000; font-weight: 800;"
                )
                self.integrity_label.setToolTip(
                    "Critical values not found in translation: " + ", ".join(missing)
                )

            self.query_time_label.setText(
                f"Translation time: {elapsed:.2f} seconds • background"
            )

            entry_id = (self.current_message or {}).get("entry_id", "")
            if entry_id:
                self.cache.upsert({
                    "outlook_entry_id": entry_id,
                    "source_language": source_language,
                    "original_subject": subject,
                    "translated_subject": translated_subject,
                    "query_hash": text_hash(subject, body),
                    "original_query": body,
                    "translated_query": final_body,
                    "integrity_percent": score,
                    "integrity_status": integrity_status,
                    "translated_at": now_string(),
                })

            self.log.upsert({
                **self.base_log_data(),
                "source_language": source_language,
                "response_language": source_language,
                "query_translation_seconds": f"{elapsed:.2f}",
                "status": "QUERY_TRANSLATED",
            })

            self.statusBar().showMessage(
                f"Query translated successfully in {elapsed:.2f} seconds."
            )

    def translate_response(self):
        response_subject = self.response_subject.text().strip()
        response_body = self.response_text.toPlainText().strip()

        if not response_body:
            QMessageBox.warning(
                self,
                "Missing Response",
                "Type or paste your approved English response first.",
            )
            return

        source_language = self.selected_source_language()
        if not source_language:
            QMessageBox.warning(
                self,
                "Language",
                "Translate the query first or select the source language manually.",
            )
            return

        token = str(uuid.uuid4())
        session_token = self.session_id
        self.response_job_token = token
        self.response_translation_running = True
        self.translate_response_button.setEnabled(False)
        self.response_time_label.setText("Translation running in background…")
        self.statusBar().showMessage(
            "Translating response in background — the helpdesk window remains responsive."
        )
        started = time.perf_counter()

        def worker():
            try:
                if source_language.lower() == "english":
                    translated_subject = response_subject
                    translated_body = response_body
                else:
                    translated_subject, translated_body = translate_text(
                        response_subject,
                        response_body,
                        source_language,
                    )

                elapsed = time.perf_counter() - started
                self.response_translation_results.put((
                    "OK",
                    token,
                    session_token,
                    response_subject,
                    response_body,
                    source_language,
                    translated_subject or response_subject,
                    translated_body or response_body,
                    elapsed,
                ))
            except Exception as exc:
                self.response_translation_results.put((
                    "ERR",
                    token,
                    session_token,
                    response_subject,
                    response_body,
                    source_language,
                    str(exc),
                    "",
                    time.perf_counter() - started,
                ))

        threading.Thread(
            target=worker,
            name="AxaltaResponseTranslation",
            daemon=True,
        ).start()

    def _process_response_translation_results(self):
        while True:
            try:
                result = self.response_translation_results.get_nowait()
            except queue.Empty:
                return

            status, token, session_token, response_subject, response_body, source_language, value, translated_body, elapsed = result

            if token != self.response_job_token or session_token != self.session_id:
                continue

            self.response_translation_running = False
            self.translate_response_button.setEnabled(True)

            if status == "ERR":
                self.response_time_label.setText(
                    f"Translation failed after {elapsed:.2f} seconds"
                )
                self.log.upsert({**self.base_log_data(), "status": "FAILED"})
                QMessageBox.critical(
                    self,
                    "Response Translation Error",
                    f"Could not translate the response.\n\n{value}",
                )
                return

            final_subject = value
            final_body = translated_body
            self.response_seconds = elapsed
            self.final_subject.setText(final_subject)
            self.final_response.setPlainText(final_body)

            score, missing, integrity_status = translation_integrity(
                f"{response_subject}\n{response_body}",
                f"{final_subject}\n{final_body}",
            )
            self.current_integrity_percent = score
            self.current_integrity_status = integrity_status
            self.current_missing_tokens = missing

            if integrity_status == "VERIFIED":
                self.integrity_label.setText(
                    f"Translation Integrity: {score}%  •  VERIFIED"
                )
                self.integrity_label.setStyleSheet(
                    "color: #137333; font-weight: 800;"
                )
            else:
                self.integrity_label.setText(
                    f"Translation Integrity: {score}%  •  REVIEW REQUIRED"
                )
                self.integrity_label.setStyleSheet(
                    "color: #8B0000; font-weight: 800;"
                )
                self.integrity_label.setToolTip(
                    "Critical values not found in translation: " + ", ".join(missing)
                )

            self.response_time_label.setText(
                f"Translation time: {elapsed:.2f} seconds • background"
            )

            self.log.upsert({
                **self.base_log_data(),
                "source_language": source_language,
                "response_language": source_language,
                "response_translation_seconds": f"{elapsed:.2f}",
                "total_elapsed_seconds": f"{time.time() - self.session_started:.2f}",
                "status": "RESPONSE_READY",
            })

            self.statusBar().showMessage(
                f"Response translated successfully in {elapsed:.2f} seconds."
            )

    # --------------------------------------------------------------- Copying
    def copy_subject(self):
        text = self.final_subject.text().strip()
        if not text:
            QMessageBox.information(
                self, "Nothing to Copy", "Translate the response first."
            )
            return
        QApplication.clipboard().setText(text)
        self.statusBar().showMessage("Final subject copied to clipboard.")

    def copy_response(self):
        text = self.final_response.toPlainText().strip()
        if not text:
            QMessageBox.information(
                self, "Nothing to Copy", "Translate the response first."
            )
            return
        QApplication.clipboard().setText(text)
        self.statusBar().showMessage("Final response copied to clipboard.")

    # ------------------------------------------------------------- Completion
    def mark_completed(self):
        if not self.final_response.toPlainText().strip():
            QMessageBox.information(
                self,
                "Response Not Ready",
                "Translate the response before marking the item completed.",
            )
            return

        if self.current_integrity_status == "REVIEW_REQUIRED":
            answer = QMessageBox.warning(
                self,
                "Translation Integrity Review",
                "Critical values were not found in the translated response.\n\n"
                + ", ".join(self.current_missing_tokens)
                + "\n\nReview the translation before completing the item.",
                QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Ok,
            )
            if answer != QMessageBox.StandardButton.Ok:
                return

        self.log.upsert({
            **self.base_log_data(),
            "response_translation_seconds": f"{self.response_seconds:.2f}",
            "total_elapsed_seconds": f"{time.time() - self.session_started:.2f}",
            "status": "COMPLETED",
        })

        self.refresh_outlook()
        self.statusBar().showMessage(
            "Marked COMPLETED. The actual email is still sent manually from Outlook."
        )

    # --------------------------------------------------------------- Footer
    def update_footer(self):
        elapsed = int(time.time() - self.session_started)
        minutes, seconds = divmod(elapsed, 60)
        self.statusBar().showMessage(
            f"Session AHT: {minutes:02d}:{seconds:02d}  |  "
            f"Query: {self.query_seconds:.2f}s  |  "
            f"Response: {self.response_seconds:.2f}s  |  "
            f"Model: {MODEL}  |  Cache: {CACHE_FILE.name}"
        )

    # ------------------------------------------------------------- Shutdown
    def closeEvent(self, event):
        try:
            body_exists = bool(self.query_text.toPlainText().strip())
            translation_exists = bool(self.english_query.toPlainText().strip())
            if self.current_message and body_exists:
                self.log.upsert({
                    **self.base_log_data(),
                    "query_translation_seconds": f"{self.query_seconds:.2f}",
                    "response_translation_seconds": f"{self.response_seconds:.2f}",
                    "total_elapsed_seconds": f"{time.time() - self.session_started:.2f}",
                    "status": "QUERY_TRANSLATED" if translation_exists else "NEW",
                })
        except Exception:
            pass
        event.accept()


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setStyle("Fusion")
    app.setFont(QFont("Segoe UI", 10))

    window = HelpdeskWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
