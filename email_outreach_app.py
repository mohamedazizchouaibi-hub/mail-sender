#!/usr/bin/env python3
"""
Email Outreach Automation
PyQt6 desktop app for LLM-powered personalized email campaigns.
"""

import sys
import os
import csv
import json
import smtplib
import time
import re
import threading
from datetime import datetime
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from typing import Optional, List, Dict, Any, Tuple

import requests

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QTextEdit, QProgressBar,
    QFileDialog, QFrame, QCheckBox, QSpinBox, QGroupBox,
    QMessageBox, QDialog, QDialogButtonBox, QSplitter, QComboBox,
    QScrollArea, QFormLayout, QSizePolicy,
    QTableWidget, QTableWidgetItem, QHeaderView,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QMutex, QWaitCondition
from PyQt6.QtGui import QFont, QTextCursor, QPalette, QColor

# ── Optional deps ──────────────────────────────────────────────────────────────
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    from PyPDF2 import PdfReader
    HAS_PDF = True
except ImportError:
    try:
        from pypdf import PdfReader  # type: ignore
        HAS_PDF = True
    except ImportError:
        HAS_PDF = False

try:
    from docx import Document as DocxDoc  # type: ignore
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ── Constants ──────────────────────────────────────────────────────────────────
APP_NAME    = "قل الحمد لله"
APP_VERSION = "1.0.0"
LOG_FILE      = Path("outreach_log.json")
SETTINGS_FILE = Path("outreach_settings.json")
MAX_CV_CHARS  = 3000
EMAIL_RE = re.compile(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$")

# ── Palette (Catppuccin Mocha) ─────────────────────────────────────────────────
C = {
    "base":    "#1e1e2e",
    "mantle":  "#181825",
    "crust":   "#11111b",
    "surface0":"#313244",
    "surface1":"#45475a",
    "surface2":"#585b70",
    "overlay0":"#6c7086",
    "overlay2":"#9399b2",
    "text":    "#cdd6f4",
    "subtext": "#bac2de",
    "blue":    "#89b4fa",
    "lavender":"#b4befe",
    "mauve":   "#cba6f7",
    "pink":    "#f5c2e7",
    "red":     "#f38ba8",
    "maroon":  "#eba0ac",
    "peach":   "#fab387",
    "yellow":  "#f9e2af",
    "green":   "#a6e3a1",
    "teal":    "#94e2d5",
    "sky":     "#89dceb",
    "sapphire":"#74c7ec",
}

DARK_STYLE = f"""
QWidget {{
    background-color: {C['base']};
    color: {C['text']};
    font-family: 'Segoe UI', 'SF Pro Display', 'Ubuntu', Arial, sans-serif;
    font-size: 13px;
}}
QMainWindow {{ background-color: {C['base']}; }}

/* ── Buttons ── */
QPushButton {{
    background-color: {C['surface0']}; color: {C['text']};
    border: 1px solid {C['surface1']}; border-radius: 7px;
    padding: 8px 18px; font-weight: 500; min-height: 28px;
}}
QPushButton:hover   {{ background-color: {C['surface1']}; border-color: {C['blue']}; color: #ffffff; }}
QPushButton:pressed {{ background-color: {C['surface2']}; }}
QPushButton:disabled {{ color: {C['surface1']}; border-color: {C['surface0']}; background-color: {C['mantle']}; }}

QPushButton#btnStart {{
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 {C['green']},stop:1 {C['teal']});
    color: {C['base']}; border: none; font-weight: bold; border-radius: 7px;
}}
QPushButton#btnStart:hover    {{ background: {C['teal']}; color: {C['base']}; }}
QPushButton#btnStart:disabled {{ background: {C['surface0']}; color: {C['surface2']}; }}

QPushButton#btnPause {{
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 {C['yellow']},stop:1 {C['peach']});
    color: {C['base']}; border: none; font-weight: bold; border-radius: 7px;
}}
QPushButton#btnPause:hover {{ background: {C['peach']}; color: {C['base']}; }}

QPushButton#btnStop {{
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 {C['red']},stop:1 {C['maroon']});
    color: {C['base']}; border: none; font-weight: bold; border-radius: 7px;
}}
QPushButton#btnStop:hover {{ background: {C['maroon']}; color: {C['base']}; }}

QPushButton#btnBrowse {{
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 {C['blue']},stop:1 {C['sapphire']});
    color: {C['base']}; border: none; padding: 5px 14px; font-weight: 600; border-radius: 6px;
}}
QPushButton#btnBrowse:hover {{ background: {C['sapphire']}; color: {C['base']}; }}

QPushButton#btnTest {{
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 {C['mauve']},stop:1 {C['pink']});
    color: {C['base']}; border: none; font-weight: 600; border-radius: 7px;
}}
QPushButton#btnTest:hover {{ background: {C['pink']}; color: {C['base']}; }}

QPushButton#btnGenerate {{
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 {C['sky']},stop:1 {C['blue']});
    color: {C['base']}; border: none; font-weight: 600; border-radius: 7px;
}}
QPushButton#btnGenerate:hover {{ background: {C['blue']}; color: {C['base']}; }}

/* ── Inputs ── */
QLineEdit, QSpinBox, QComboBox {{
    background-color: {C['mantle']}; border: 1px solid {C['surface1']};
    border-radius: 6px; padding: 7px 11px;
    color: {C['text']}; selection-background-color: {C['blue']}; min-height: 22px;
}}
QLineEdit:focus, QSpinBox:focus, QComboBox:focus {{
    border-color: {C['blue']}; background-color: {C['surface0']};
}}
QLineEdit:hover, QSpinBox:hover, QComboBox:hover {{ border-color: {C['overlay0']}; }}

QComboBox::drop-down {{ border: none; width: 22px; }}
QComboBox QAbstractItemView {{
    background-color: {C['surface0']}; border: 1px solid {C['surface1']};
    color: {C['text']}; selection-background-color: {C['blue']};
    selection-color: {C['base']}; outline: none;
}}
QSpinBox::up-button, QSpinBox::down-button {{
    width: 20px; background: {C['surface1']}; border: none; border-radius: 3px;
}}
QSpinBox::up-button:hover, QSpinBox::down-button:hover {{ background: {C['surface2']}; }}

/* ── Text edit / log ── */
QTextEdit {{
    background-color: {C['crust']}; border: 1px solid {C['surface0']};
    border-radius: 6px; padding: 6px; color: {C['text']};
    font-family: 'Consolas', 'JetBrains Mono', 'Fira Code', 'Courier New', monospace;
    font-size: 12px; line-height: 1.4;
}}

/* ── Progress bar ── */
QProgressBar {{
    border: 1px solid {C['surface1']}; border-radius: 6px;
    background-color: {C['mantle']}; text-align: center;
    color: {C['text']}; min-height: 24px; font-weight: 500;
}}
QProgressBar::chunk {{
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
        stop:0 {C['blue']}, stop:0.5 {C['mauve']}, stop:1 {C['pink']});
    border-radius: 5px;
}}

/* ── Group boxes ── */
QGroupBox {{
    border: 1px solid {C['surface0']}; border-radius: 10px;
    margin-top: 1.4em; padding-top: 0.8em; padding-bottom: 4px;
    font-weight: 600; color: {C['blue']};
    background-color: rgba(17,17,27,0.4);
}}
QGroupBox::title {{
    subcontrol-origin: margin; left: 14px; padding: 2px 6px;
    color: {C['blue']}; background-color: {C['base']}; border-radius: 4px;
}}

/* ── Checkboxes ── */
QCheckBox {{ color: {C['text']}; spacing: 8px; }}
QCheckBox::indicator {{
    width: 17px; height: 17px;
    border: 1px solid {C['surface1']}; border-radius: 4px;
    background-color: {C['mantle']};
}}
QCheckBox::indicator:hover {{ border-color: {C['blue']}; }}
QCheckBox::indicator:checked {{
    background-color: {C['blue']}; border-color: {C['blue']};
    image: none;
}}

/* ── Scrollbars ── */
QScrollBar:vertical {{ background: {C['base']}; width: 8px; border-radius: 4px; }}
QScrollBar::handle:vertical {{
    background: {C['surface1']}; border-radius: 4px; min-height: 30px;
}}
QScrollBar::handle:vertical:hover {{ background: {C['surface2']}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar:horizontal {{ background: {C['base']}; height: 8px; border-radius: 4px; }}
QScrollBar::handle:horizontal {{ background: {C['surface1']}; border-radius: 4px; }}

/* ── Labels ── */
QLabel#lblTitle    {{ font-size: 22px; font-weight: bold; color: {C['blue']}; }}
QLabel#lblSubtitle {{ font-size: 11px; color: {C['overlay0']}; letter-spacing: 0.3px; }}
QLabel#lblStatus   {{ color: {C['green']}; font-style: italic; }}
QLabel#lblBadge    {{
    background-color: {C['surface0']}; border: 1px solid {C['surface1']};
    border-radius: 10px; padding: 2px 10px; color: {C['text']}; font-size: 11px;
}}
QLabel#lblBadgeGood {{
    background-color: rgba(166,227,161,0.15); border: 1px solid {C['green']};
    border-radius: 10px; padding: 2px 10px; color: {C['green']}; font-size: 11px;
}}

/* ── Dialog ── */
QDialog {{ background-color: {C['base']}; }}
QDialogButtonBox QPushButton {{ min-width: 90px; }}

/* ── Splitter ── */
QSplitter::handle {{ background-color: {C['surface0']}; }}
QSplitter::handle:horizontal {{ width: 1px; }}

/* ── Scroll area ── */
QScrollArea {{ border: none; background-color: transparent; }}
"""

# ── Helpers ────────────────────────────────────────────────────────────────────

def valid_email(addr: str) -> bool:
    return bool(EMAIL_RE.match(addr.strip()))


def parse_cv(path: str) -> str:
    suffix = Path(path).suffix.lower()
    text = ""
    if suffix == ".pdf":
        if not HAS_PDF:
            return "[PDF parsing unavailable — install PyPDF2 or pypdf]"
        try:
            reader = PdfReader(path)
            text = " ".join(p.extract_text() or "" for p in reader.pages)
        except Exception as e:
            return f"[PDF error: {e}]"
    elif suffix in (".docx", ".doc"):
        if not HAS_DOCX:
            return "[DOCX parsing unavailable — install python-docx]"
        try:
            doc = DocxDoc(path)
            text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        except Exception as e:
            return f"[DOCX error: {e}]"
    else:
        return "[Unsupported format — use PDF or DOCX]"
    text = re.sub(r"\s+", " ", text).strip()
    return text[:MAX_CV_CHARS] + ("…" if len(text) > MAX_CV_CHARS else "")


def load_contacts(path: str) -> List[Dict[str, str]]:
    if HAS_PANDAS:
        df = pd.read_csv(path)
        df.columns = [c.strip().lower() for c in df.columns]
        return df.where(pd.notna(df), "").astype(str).to_dict(orient="records")
    rows: List[Dict[str, str]] = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append({k.strip().lower(): v.strip() for k, v in row.items()})
    return rows


def get_field(row: Dict[str, str], *keys: str, default: str = "") -> str:
    for k in keys:
        v = row.get(k, "").strip()
        if v:
            return v
    return default


def load_template(path: str) -> str:
    suffix = Path(path).suffix.lower()
    if suffix == ".txt":
        try:
            return Path(path).read_text(encoding="utf-8").strip()
        except Exception as e:
            return f"[Template error: {e}]"
    if suffix == ".pdf":
        if not HAS_PDF:
            return "[PDF parsing unavailable — install PyPDF2 or pypdf]"
        try:
            reader = PdfReader(path)
            return " ".join(p.extract_text() or "" for p in reader.pages).strip()
        except Exception as e:
            return f"[PDF error: {e}]"
    if suffix in (".docx", ".doc"):
        if not HAS_DOCX:
            return "[DOCX parsing unavailable — install python-docx]"
        try:
            doc = DocxDoc(path)
            return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        except Exception as e:
            return f"[DOCX error: {e}]"
    return "[Unsupported format — use .txt, .pdf, or .docx]"


def build_prompt(cv_text: str, contact: Dict[str, str], goal: str,
                 company_context: str = "",
                 sender_name: str = "", sender_email: str = "",
                 sender_phone: str = "", sender_extra: str = "",
                 template_text: str = "") -> str:
    name    = get_field(contact, "name", "full_name", "first_name", "contact", default="there")
    company = get_field(contact, "company", "organization", "employer", "firm",
                        default="your organization")
    role    = get_field(contact, "role", "position", "title", "job_title")

    skip = {"name", "full_name", "first_name", "email", "e-mail", "email_address",
            "company", "organization", "employer", "firm", "role", "position",
            "title", "job_title", "contact"}
    extras = "\n".join(
        f"- {k.replace('_', ' ').title()}: {v}"
        for k, v in contact.items()
        if k not in skip and v.strip()
    )
    role_line = f"- Role: {role}\n" if role else ""
    context_block = (
        f'\nCompany Research (use this to personalize the email):\n"""\n{company_context}\n"""\n'
        if company_context else ""
    )

    sender_label = sender_name or "the sender"
    contact_lines = []
    if sender_phone:
        contact_lines.append(sender_phone)
    if sender_email:
        contact_lines.append(sender_email)
    contact_line = " | ".join(contact_lines)
    contact_info = f"- Contact: {contact_line}\n" if contact_line else ""
    extra_info = f"- Additional info: {sender_extra}\n" if sender_extra.strip() else ""

    sig_parts = ["Best regards,", sender_label]
    if sender_phone:
        sig_parts.append(sender_phone)
    if sender_email:
        sig_parts.append(sender_email)
    signature = " / ".join(sig_parts)

    template_block = (
        f'\nEmail Template (follow this structure and tone closely, '
        f'but personalize all placeholder content for this specific recipient):\n"""\n{template_text}\n"""\n'
        if template_text else ""
    )
    template_rule = (
        "\n8. Follow the structure and tone of the provided email template above."
        if template_text else ""
    )

    return f"""You are an expert email copywriter helping {sender_label} reach out for: {goal}

About the sender ({sender_label}):
{contact_info}{extra_info}
The full CV / Resume (use this as the authoritative source for the sender's background, skills, experience, and education):
\"\"\"
{cv_text}
\"\"\"

Recipient:
- Name: {name}
- Company: {company}
{role_line}{extras}{context_block}{template_block}
Task: Write a highly personalized, professional, concise outreach email from {sender_label}.

Rules:
1. Derive the sender's background, skills, and experience entirely from the CV text above — do NOT invent details.
2. If the recipient name is known, address them directly (e.g. "Dear [Name],"); otherwise use "Dear Hiring Manager,".
3. Reference 1-2 specific aspects of the company/role that make it a strong fit for the sender's background.
4. Keep the body to 3-4 short paragraphs — no fluff, no generic opener.
5. End with a soft call to action ("I'd welcome a quick discussion") before the signature.
6. Always close with this exact signature block on its own lines: {signature}
7. Never use "I hope this email finds you well" or similar clichés.{template_rule}

Respond with ONLY valid JSON (no markdown, no extra text):
{{"subject": "...", "body": "..."}}"""


# ── Company Researcher ─────────────────────────────────────────────────────────

class CompanyResearcher:
    _UA = "Mozilla/5.0 (compatible; outreach-research/1.0)"

    def __init__(self):
        self._cache: Dict[str, str] = {}

    def research(self, company: str, website: str = "") -> str:
        key = company.strip().lower()
        if key in self._cache:
            return self._cache[key]
        ctx = self._scrape(website) if website else ""
        if not ctx:
            ctx = self._duckduckgo(company)
        self._cache[key] = ctx
        return ctx

    def _scrape(self, url: str) -> str:
        if not url.startswith("http"):
            url = "https://" + url
        try:
            r = requests.get(url, timeout=10, headers={"User-Agent": self._UA})
            r.raise_for_status()
            html = r.text
            for pat in [
                r'<meta[^>]+name=["\']description["\'][^>]+content=["\'](.*?)["\']',
                r'<meta[^>]+content=["\'](.*?)["\'][^>]+name=["\']description["\']',
                r'<meta[^>]+property=["\']og:description["\'][^>]+content=["\'](.*?)["\']',
                r'<meta[^>]+content=["\'](.*?)["\'][^>]+property=["\']og:description["\']',
            ]:
                m = re.search(pat, html, re.IGNORECASE | re.DOTALL)
                if m:
                    desc = m.group(1).strip()
                    if len(desc) > 40:
                        return desc[:600]
            text = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", html,
                          flags=re.DOTALL | re.IGNORECASE)
            text = re.sub(r"<[^>]+>", " ", text)
            text = re.sub(r"\s+", " ", text).strip()
            return text[:600]
        except Exception:
            return ""

    def _duckduckgo(self, company: str) -> str:
        try:
            r = requests.get(
                "https://api.duckduckgo.com/",
                params={"q": company, "format": "json", "no_html": "1", "skip_disambig": "1"},
                timeout=12,
                headers={"User-Agent": self._UA},
            )
            data = r.json()
            abstract = data.get("AbstractText", "").strip()
            if abstract:
                return abstract[:600]
            topics = data.get("RelatedTopics", [])
            if topics and isinstance(topics[0], dict):
                return topics[0].get("Text", "")[:400]
        except Exception:
            pass
        return ""


def extract_sender_info_from_cv(cv_text: str) -> Dict[str, str]:
    info: Dict[str, str] = {"name": "", "phone": "", "email": "", "extra": ""}
    m = re.search(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", cv_text)
    if m:
        info["email"] = m.group(0)
    m = re.search(r"(?:\+?\d[\d\s\-().]{6,}\d)", cv_text)
    if m:
        info["phone"] = m.group(0).strip()
    for line in cv_text.split("\n"):
        line = line.strip()
        if not line:
            continue
        words = line.split()
        if (2 <= len(words) <= 4
                and all(w[0].isupper() for w in words if w)
                and not any(ch.isdigit() for ch in line)
                and len(line) < 60):
            info["name"] = line
            break
    return info


def parse_llm_json(text: str) -> Tuple[str, str]:
    cleaned = re.sub(r"```(?:json)?\s*|\s*```", "", text).strip()
    try:
        data = json.loads(cleaned)
        return data.get("subject", ""), data.get("body", "")
    except json.JSONDecodeError:
        ms = re.search(r'"subject"\s*:\s*"((?:[^"\\]|\\.)*)"', cleaned)
        mb = re.search(r'"body"\s*:\s*"((?:[^"\\]|\\.)*)"', cleaned, re.DOTALL)
        subject = ms.group(1) if ms else "Outreach"
        body    = mb.group(1).replace("\\n", "\n") if mb else cleaned
        return subject, body


# ── LLM Client ─────────────────────────────────────────────────────────────────

class LLMClient:
    _REGISTRY: Dict[str, Tuple[str, str]] = {
        "OpenAI":           ("https://api.openai.com/v1/chat/completions",
                             "gpt-4o-mini"),
        "Gemini":           ("https://generativelanguage.googleapis.com/v1beta/models",
                             "gemini-2.0-flash"),
        "Groq":             ("https://api.groq.com/openai/v1/chat/completions",
                             "llama-3.3-70b-versatile"),
        "OpenRouter":       ("https://openrouter.ai/api/v1/chat/completions",
                             "google/gemini-2.0-flash-exp:free"),
        "Anthropic":        ("https://api.anthropic.com/v1/messages",
                             "claude-3-5-haiku-20241022"),
        "Mistral":          ("https://api.mistral.ai/v1/chat/completions",
                             "mistral-small-latest"),
        "Mistral Codestral":("https://codestral.mistral.ai/v1/chat/completions",
                             "codestral-latest"),
        "Together AI":      ("https://api.together.xyz/v1/chat/completions",
                             "meta-llama/Llama-3.3-70B-Instruct-Turbo"),
        "Cohere":           ("https://api.cohere.com/v2/chat",
                             "command-r-plus-08-2024"),
        "Cerebras":         ("https://api.cerebras.ai/v1/chat/completions",
                             "llama-3.3-70b"),
        "NVIDIA NIM":       ("https://integrate.api.nvidia.com/v1/chat/completions",
                             "meta/llama-3.3-70b-instruct"),
        "GitHub Models":    ("https://models.inference.ai.azure.com/chat/completions",
                             "gpt-4o-mini"),
        "Fireworks":        ("https://api.fireworks.ai/inference/v1/chat/completions",
                             "accounts/fireworks/models/llama-v3p3-70b-instruct"),
        "Nebius":           ("https://api.studio.nebius.ai/v1/chat/completions",
                             "meta-llama/Llama-3.3-70B-Instruct"),
        "Novita":           ("https://api.novita.ai/v3/openai/chat/completions",
                             "meta-llama/llama-3.3-70b-instruct"),
        "Hyperbolic":       ("https://api.hyperbolic.xyz/v1/chat/completions",
                             "meta-llama/Llama-3.3-70B-Instruct"),
        "SambaNova":        ("https://api.sambanova.ai/v1/chat/completions",
                             "Meta-Llama-3.3-70B-Instruct"),
        "Scaleway":         ("https://api.scaleway.ai/v1/chat/completions",
                             "llama-3.3-70b-instruct"),
        "Inference.net":    ("https://api.inference.net/v1/chat/completions",
                             "meta-llama/llama-3.3-70b-instruct/fp-8"),
    }
    MAX_RETRIES  = 4
    BASE_BACKOFF = 10

    def __init__(self, provider: str, api_key: str, model: str = ""):
        self.provider  = provider
        self.api_key   = api_key
        base, default  = self._REGISTRY.get(provider, ("", ""))
        self._base_url = base
        self.model     = model.strip() or default

    def generate(self, prompt: str) -> str:
        dispatch = {
            "Gemini":    self._gemini,
            "Anthropic": self._anthropic,
            "Cohere":    self._cohere,
        }
        fn = dispatch.get(self.provider, self._openai_compat)
        for attempt in range(self.MAX_RETRIES):
            try:
                return fn(prompt)
            except requests.HTTPError as e:
                if e.response is not None and e.response.status_code == 429:
                    wait = self.BASE_BACKOFF * (2 ** attempt)
                    ra = e.response.headers.get("Retry-After")
                    if ra:
                        try:
                            wait = max(wait, int(ra))
                        except ValueError:
                            pass
                    if attempt < self.MAX_RETRIES - 1:
                        time.sleep(wait)
                        continue
                raise
        raise RuntimeError("LLM retries exhausted")

    def _gemini(self, prompt: str) -> str:
        url = f"{self._base_url}/{self.model}:generateContent?key={self.api_key}"
        r = requests.post(
            url,
            json={
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {"temperature": 0.7, "maxOutputTokens": 1024},
            },
            timeout=45,
        )
        r.raise_for_status()
        return r.json()["candidates"][0]["content"]["parts"][0]["text"]

    def _openai_compat(self, prompt: str) -> str:
        headers: Dict[str, str] = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type":  "application/json",
        }
        if self.provider == "OpenRouter":
            headers["HTTP-Referer"] = "https://email-outreach-app"
            headers["X-Title"]      = "Email Outreach Automation"
        r = requests.post(
            self._base_url,
            headers=headers,
            json={
                "model":       self.model,
                "messages":    [{"role": "user", "content": prompt}],
                "temperature": 0.7,
                "max_tokens":  1024,
            },
            timeout=45,
        )
        r.raise_for_status()
        return r.json()["choices"][0]["message"]["content"]

    def _anthropic(self, prompt: str) -> str:
        r = requests.post(
            self._base_url,
            headers={
                "x-api-key":         self.api_key,
                "anthropic-version": "2023-06-01",
                "Content-Type":      "application/json",
            },
            json={
                "model":      self.model,
                "max_tokens": 1024,
                "messages":   [{"role": "user", "content": prompt}],
            },
            timeout=45,
        )
        r.raise_for_status()
        return r.json()["content"][0]["text"]

    def _cohere(self, prompt: str) -> str:
        r = requests.post(
            self._base_url,
            headers={
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type":  "application/json",
            },
            json={
                "model":    self.model,
                "messages": [{"role": "user", "content": prompt}],
            },
            timeout=45,
        )
        r.raise_for_status()
        return r.json()["message"]["content"][0]["text"]


# ── Email Sender ───────────────────────────────────────────────────────────────

class EmailSender:
    def __init__(self, sender: str, password: str,
                 host: str = "smtp.gmail.com", port: int = 587):
        self.sender   = sender
        self.password = password
        self.host     = host
        self.port     = port

    def send(self, to: str, subject: str, body: str,
             cv_path: Optional[str] = None) -> None:
        msg = MIMEMultipart("mixed")
        msg["From"]    = self.sender
        msg["To"]      = to
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain", "utf-8"))

        if cv_path and os.path.isfile(cv_path):
            with open(cv_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f'attachment; filename="{Path(cv_path).name}"',
            )
            msg.attach(part)

        with smtplib.SMTP(self.host, self.port, timeout=30) as srv:
            srv.ehlo()
            srv.starttls()
            srv.login(self.sender, self.password)
            srv.sendmail(self.sender, to, msg.as_string())


# ── Preview Dialog ─────────────────────────────────────────────────────────────

class PreviewDialog(QDialog):
    def __init__(self, name: str, email: str, subject: str, body: str,
                 parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Preview — {name} <{email}>")
        self.resize(660, 540)
        lay = QVBoxLayout(self)
        lay.setSpacing(12)
        lay.setContentsMargins(16, 16, 16, 16)

        to_lbl = QLabel(f"<b>To:</b> {name} &lt;{email}&gt;")
        to_lbl.setStyleSheet(f"color: {C['subtext']}; font-size: 13px;")
        lay.addWidget(to_lbl)

        lay.addWidget(QLabel("Subject:"))
        self.inp_subject = QLineEdit(subject)
        lay.addWidget(self.inp_subject)

        lay.addWidget(QLabel("Body:"))
        self.txt_body = QTextEdit()
        self.txt_body.setPlainText(body)
        self.txt_body.setMinimumHeight(320)
        lay.addWidget(self.txt_body)

        btns = QDialogButtonBox()
        btns.addButton("Send", QDialogButtonBox.ButtonRole.AcceptRole)
        btns.addButton("Skip", QDialogButtonBox.ButtonRole.RejectRole)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    @property
    def subject(self) -> str:
        return self.inp_subject.text().strip()

    @property
    def body(self) -> str:
        return self.txt_body.toPlainText().strip()


# ── Draft Email ────────────────────────────────────────────────────────────────

class DraftEmail:
    __slots__ = ("name", "email", "company", "subject", "body",
                 "contact", "include", "company_context")

    def __init__(self, name: str, email: str, company: str,
                 subject: str, body: str, contact: Dict,
                 company_context: str = ""):
        self.name            = name
        self.email           = email
        self.company         = company
        self.subject         = subject
        self.body            = body
        self.contact         = contact
        self.include         = True
        self.company_context = company_context


# ── CV Info Extraction Thread ──────────────────────────────────────────────────

class CVInfoThread(QThread):
    sig_done = pyqtSignal(dict)

    def __init__(self, cv_text: str,
                 provider: str = "", api_key: str = "", model: str = "",
                 parent=None):
        super().__init__(parent)
        self.cv_text  = cv_text
        self.provider = provider
        self.api_key  = api_key
        self.model    = model

    def run(self):
        info = extract_sender_info_from_cv(self.cv_text)
        if self.api_key and self.provider:
            try:
                llm = LLMClient(self.provider, self.api_key, self.model)
                prompt = (
                    "Extract the following fields from this CV text and return ONLY valid JSON "
                    "(no markdown, no extra text):\n"
                    '{"name": "full name of the CV owner", '
                    '"phone": "phone number or empty string", '
                    '"email": "email address or empty string", '
                    '"extra": "LinkedIn URL or portfolio URL if present, else empty string"}\n\n'
                    f'CV:\n"""\n{self.cv_text}\n"""'
                )
                raw = llm.generate(prompt)
                cleaned = re.sub(r"```(?:json)?\s*|\s*```", "", raw).strip()
                parsed = json.loads(cleaned)
                for k in ("name", "phone", "email", "extra"):
                    if parsed.get(k, "").strip():
                        info[k] = parsed[k].strip()
            except Exception:
                pass
        self.sig_done.emit(info)


# ── Generation Thread ──────────────────────────────────────────────────────────

class GenerationThread(QThread):
    sig_log      = pyqtSignal(str, str)
    sig_progress = pyqtSignal(int, int)
    sig_status   = pyqtSignal(str)
    sig_done     = pyqtSignal(list)

    def __init__(self, cfg: Dict[str, Any], contacts: List[Dict],
                 cv_text: str, parent=None):
        super().__init__(parent)
        self.cfg           = cfg
        self.contacts      = contacts
        self.cv_text       = cv_text
        self._stopped      = False
        self._researcher   = CompanyResearcher()

    def stop(self):
        self._stopped = True

    def run(self):
        cfg   = self.cfg
        llm   = LLMClient(cfg["provider"], cfg["api_key"], cfg.get("model", ""))
        total = len(self.contacts)
        drafts: List[DraftEmail] = []

        self.sig_log.emit(f"Generating {total} draft(s)…", "info")
        self.sig_progress.emit(0, total)

        for i, contact in enumerate(self.contacts):
            if self._stopped:
                break

            email   = get_field(contact, "email", "e-mail", "email_address", "mail")
            name    = get_field(contact, "name", "full_name", "first_name",
                                default=f"Contact {i + 1}")
            company = get_field(contact, "company", "organization", "employer", "firm",
                                default="")

            if not email or not valid_email(email):
                self.sig_log.emit(
                    f"[{i+1}/{total}] Invalid/missing email for '{name}' — skipped", "warn"
                )
                self.sig_progress.emit(i + 1, total)
                continue

            company_context = ""
            if cfg.get("research") and company:
                self.sig_status.emit(f"Researching {company}…")
                website = get_field(contact, "website", "url", "homepage", "site")
                company_context = self._researcher.research(company, website)
                if company_context:
                    self.sig_log.emit(
                        f"[{i+1}/{total}] Context: {company} ({len(company_context)} chars)",
                        "info",
                    )
                else:
                    self.sig_log.emit(
                        f"[{i+1}/{total}] No context found for '{company}'", "warn"
                    )

            self.sig_status.emit(f"Generating draft for {name}…")
            self.sig_log.emit(f"[{i+1}/{total}] Generating → {name} <{email}>", "info")
            try:
                raw = llm.generate(
                    build_prompt(
                        self.cv_text, contact, cfg["goal"], company_context,
                        sender_name=cfg.get("sender_name", ""),
                        sender_email=cfg.get("sender_email", ""),
                        sender_phone=cfg.get("sender_phone", ""),
                        sender_extra=cfg.get("sender_extra", ""),
                        template_text=cfg.get("template_text", ""),
                    )
                )
                subject, body = parse_llm_json(raw)
                if not body.strip():
                    raise ValueError("LLM returned empty body")
                subject = subject or cfg["goal"]
            except Exception as exc:
                self.sig_log.emit(f"[{i+1}/{total}] LLM error: {exc}", "error")
                self.sig_progress.emit(i + 1, total)
                continue

            drafts.append(
                DraftEmail(name, email, company, subject, body, contact, company_context)
            )
            self.sig_log.emit(
                f"[{i+1}/{total}] ✓ Draft ready: {name} — {subject}", "success"
            )
            self.sig_progress.emit(i + 1, total)

            if i < total - 1 and not self._stopped:
                time.sleep(3)

        status = "All drafts generated." if not self._stopped else "Generation stopped."
        self.sig_status.emit(status)
        self.sig_done.emit(drafts)


# ── Preview-All Dialog ─────────────────────────────────────────────────────────

class PreviewAllDialog(QDialog):
    TABLE_STYLE = f"""
        QTableWidget {{
            background-color: {C['base']};
            alternate-background-color: {C['mantle']};
            gridline-color: transparent;
            border: none;
        }}
        QTableWidget::item {{ padding: 5px 8px; border-radius: 0px; }}
        QTableWidget::item:selected {{ background-color: {C['surface0']}; color: {C['text']}; }}
        QHeaderView::section {{
            background-color: {C['surface0']}; color: {C['blue']};
            border: none; border-bottom: 1px solid {C['surface1']};
            padding: 7px 6px; font-weight: 600;
        }}
    """

    def __init__(self, drafts: List[DraftEmail], parent=None):
        super().__init__(parent)
        self._drafts          = drafts
        self._current_row     = -1
        self._blocking        = False
        self.setWindowTitle(f"Review {len(drafts)} Generated Email(s)")
        self.setMinimumSize(960, 620)
        self.resize(1140, 700)
        self._build_ui()
        if drafts:
            self.table.selectRow(0)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(12)
        root.setContentsMargins(16, 16, 16, 16)

        # Header bar
        hdr = QWidget()
        hdr.setStyleSheet(f"background-color: {C['surface0']}; border-radius: 8px;")
        hh = QHBoxLayout(hdr)
        hh.setContentsMargins(14, 10, 14, 10)
        info_lbl = QLabel(f"  {len(self._drafts)} email draft(s) ready — review, edit, then send")
        info_lbl.setStyleSheet(f"color: {C['blue']}; font-weight: 600; font-size: 13px; background: transparent;")
        hh.addWidget(info_lbl)
        hh.addStretch()
        hint = QLabel("Click a row to preview & edit  ·  uncheck to skip")
        hint.setStyleSheet(f"color: {C['overlay0']}; font-size: 11px; background: transparent;")
        hh.addWidget(hint)
        root.addWidget(hdr)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(1)
        splitter.addWidget(self._build_left())
        splitter.addWidget(self._build_right())
        splitter.setSizes([380, 680])
        root.addWidget(splitter, 1)
        root.addWidget(self._build_bottom())

    def _build_left(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 0, 6, 0)
        v.setSpacing(6)

        self.table = QTableWidget(len(self._drafts), 3)
        self.table.setHorizontalHeaderLabels(["", "Recipient", "Subject"])
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        hh.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        hh.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(0, 34)
        self.table.setColumnWidth(1, 186)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet(self.TABLE_STYLE)

        for r, d in enumerate(self._drafts):
            self.table.setRowHeight(r, 50)
            chk = QTableWidgetItem()
            chk.setFlags(
                Qt.ItemFlag.ItemIsUserCheckable |
                Qt.ItemFlag.ItemIsEnabled |
                Qt.ItemFlag.ItemIsSelectable
            )
            chk.setCheckState(Qt.CheckState.Checked)
            self.table.setItem(r, 0, chk)
            self.table.setItem(r, 1, QTableWidgetItem(f"{d.name}\n{d.email}"))
            self.table.setItem(r, 2, QTableWidgetItem(d.subject))

        self.table.currentCellChanged.connect(lambda cur, _cc, _pr, _pc: self._on_row_changed(cur))
        self.table.itemChanged.connect(self._on_item_changed)
        v.addWidget(self.table)

        bw = QWidget()
        bh = QHBoxLayout(bw)
        bh.setContentsMargins(0, 0, 0, 0)
        b_all  = QPushButton("Select All")
        b_none = QPushButton("Deselect All")
        b_all.clicked.connect(lambda: self._bulk_check(True))
        b_none.clicked.connect(lambda: self._bulk_check(False))
        bh.addWidget(b_all)
        bh.addWidget(b_none)
        bh.addStretch()
        v.addWidget(bw)
        return w

    def _build_right(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(6, 0, 0, 0)
        v.setSpacing(10)

        self.lbl_to = QLabel("← Select a draft to preview")
        self.lbl_to.setStyleSheet(f"color: {C['text']}; font-weight: 600; font-size: 13px;")
        self.lbl_to.setWordWrap(True)
        v.addWidget(self.lbl_to)

        self.lbl_context = QLabel("")
        self.lbl_context.setStyleSheet(
            f"color: {C['overlay0']}; font-size: 11px; font-style: italic; "
            f"background-color: {C['surface0']}; border-radius: 5px; padding: 4px 8px;"
        )
        self.lbl_context.setWordWrap(True)
        self.lbl_context.setMaximumHeight(46)
        self.lbl_context.setVisible(False)
        v.addWidget(self.lbl_context)

        subj_row = QHBoxLayout()
        lbl_subj = QLabel("Subject:")
        lbl_subj.setStyleSheet(f"color: {C['subtext']}; font-weight: 500;")
        subj_row.addWidget(lbl_subj)
        subj_row.addStretch()
        lbl_edit_hint = QLabel("✏  editable")
        lbl_edit_hint.setStyleSheet(f"color: {C['peach']}; font-size: 11px;")
        subj_row.addWidget(lbl_edit_hint)
        v.addLayout(subj_row)
        self.inp_subject = QLineEdit()
        self.inp_subject.setPlaceholderText("Select a draft to edit its subject…")
        self.inp_subject.setStyleSheet(
            f"QLineEdit {{ border: 1px solid {C['peach']}; background-color: {C['mantle']}; }}"
            f"QLineEdit:focus {{ border-color: {C['yellow']}; background-color: {C['surface0']}; }}"
        )
        self.inp_subject.textEdited.connect(self._on_subject_edited)
        v.addWidget(self.inp_subject)

        lbl_body = QLabel("Body:")
        lbl_body.setStyleSheet(f"color: {C['subtext']}; font-weight: 500;")
        v.addWidget(lbl_body)
        self.txt_body = QTextEdit()
        self.txt_body.setMinimumHeight(360)
        self.txt_body.setPlaceholderText("Select a draft from the list to preview and edit the email body…")
        self.txt_body.setStyleSheet(
            f"QTextEdit {{ border: 1px solid {C['peach']}; background-color: {C['crust']}; }}"
            f"QTextEdit:focus {{ border-color: {C['yellow']}; }}"
        )
        self.txt_body.textChanged.connect(self._on_body_edited)
        v.addWidget(self.txt_body, 1)
        return w

    def _build_bottom(self) -> QWidget:
        w = QWidget()
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 6, 0, 0)
        h.addStretch()
        btn_cancel = QPushButton("Cancel")
        btn_cancel.setMinimumHeight(40)
        btn_cancel.clicked.connect(self.reject)
        self.btn_send = QPushButton("▶  Send 0 Email(s)")
        self.btn_send.setObjectName("btnStart")
        self.btn_send.setMinimumHeight(40)
        self.btn_send.setMinimumWidth(180)
        self.btn_send.clicked.connect(self.accept)
        h.addWidget(btn_cancel)
        h.addSpacing(8)
        h.addWidget(self.btn_send)
        self._refresh_send_btn()
        return w

    def _on_row_changed(self, row: int):
        self._save_edits()
        self._load_draft(row)

    def _on_item_changed(self, item: QTableWidgetItem):
        if self._blocking or item.column() != 0:
            return
        self._drafts[item.row()].include = (
            item.checkState() == Qt.CheckState.Checked
        )
        self._refresh_send_btn()

    def _save_edits(self):
        r = self._current_row
        if r < 0 or r >= len(self._drafts):
            return
        d = self._drafts[r]
        d.subject = self.inp_subject.text().strip()
        d.body    = self.txt_body.toPlainText().strip()
        item = self.table.item(r, 2)
        if item:
            item.setText(d.subject)

    def _load_draft(self, row: int):
        if row < 0 or row >= len(self._drafts):
            return
        d  = self._drafts[row]
        co = f"  @  {d.company}" if d.company else ""
        self.lbl_to.setText(f"To:  {d.name}{co}  ‹{d.email}›")
        if d.company_context:
            snippet = d.company_context[:220].replace("\n", " ")
            ellipsis = "…" if len(d.company_context) > 220 else ""
            self.lbl_context.setText(f"🔍 {snippet}{ellipsis}")
            self.lbl_context.setVisible(True)
        else:
            self.lbl_context.setVisible(False)
        self.inp_subject.blockSignals(True)
        self.txt_body.blockSignals(True)
        self.inp_subject.setText(d.subject)
        self.txt_body.setPlainText(d.body)
        self.inp_subject.blockSignals(False)
        self.txt_body.blockSignals(False)
        self._current_row = row

    def _on_subject_edited(self, text: str):
        if self._current_row >= 0:
            self._drafts[self._current_row].subject = text
            item = self.table.item(self._current_row, 2)
            if item:
                item.setText(text)

    def _on_body_edited(self):
        if self._current_row >= 0:
            self._drafts[self._current_row].body = self.txt_body.toPlainText().strip()

    def _bulk_check(self, state: bool):
        self._blocking = True
        cs = Qt.CheckState.Checked if state else Qt.CheckState.Unchecked
        for r, d in enumerate(self._drafts):
            d.include = state
            item = self.table.item(r, 0)
            if item:
                item.setCheckState(cs)
        self._blocking = False
        self._refresh_send_btn()

    def _refresh_send_btn(self):
        n = sum(1 for d in self._drafts if d.include)
        self.btn_send.setText(f"▶  Send {n} Email{'s' if n != 1 else ''}")
        self.btn_send.setEnabled(n > 0)

    def selected_drafts(self) -> List[DraftEmail]:
        self._save_edits()
        return [d for d in self._drafts if d.include]


# ── Campaign Thread ────────────────────────────────────────────────────────────

class CampaignThread(QThread):
    sig_log      = pyqtSignal(str, str)
    sig_progress = pyqtSignal(int, int)
    sig_status   = pyqtSignal(str)
    sig_preview  = pyqtSignal(str, str, str, str)
    sig_done     = pyqtSignal(bool, str)

    def __init__(self, cfg: Dict[str, Any], contacts: List[Dict],
                 cv_text: str, cv_path: str, parent=None,
                 drafts: Optional[List[DraftEmail]] = None):
        super().__init__(parent)
        self.cfg      = cfg
        self.contacts = contacts
        self.cv_text  = cv_text
        self.cv_path  = cv_path
        self.drafts   = drafts

        self._mutex   = QMutex()
        self._wait    = QWaitCondition()
        self._paused  = False
        self._stopped = False

        self._preview_event  = threading.Event()
        self._preview_result: Tuple[bool, str, str] = (True, "", "")
        self._log_records: List[Dict] = []
        self._researcher = CompanyResearcher()

    def pause(self):
        self._mutex.lock(); self._paused = True; self._mutex.unlock()

    def resume(self):
        self._mutex.lock()
        self._paused = False; self._wait.wakeAll()
        self._mutex.unlock()

    def stop(self):
        self._mutex.lock()
        self._stopped = True; self._paused = False; self._wait.wakeAll()
        self._mutex.unlock()

    def set_preview_result(self, send: bool, subject: str, body: str):
        self._preview_result = (send, subject, body)
        self._preview_event.set()

    def _check_paused(self) -> bool:
        self._mutex.lock()
        while self._paused and not self._stopped:
            self._wait.wait(self._mutex)
        stopped = self._stopped
        self._mutex.unlock()
        return stopped

    def _log(self, msg: str, level: str = "info", recipient: str = "",
             success: Optional[bool] = None):
        self.sig_log.emit(msg, level)
        rec: Dict[str, Any] = {
            "ts": datetime.now().isoformat(timespec="seconds"),
            "level": level, "msg": msg, "recipient": recipient,
        }
        if success is not None:
            rec["success"] = success
        self._log_records.append(rec)

    def _flush_log(self):
        try:
            existing: List = []
            if LOG_FILE.exists():
                try:
                    existing = json.loads(LOG_FILE.read_text())
                except Exception:
                    pass
            existing.extend(self._log_records)
            LOG_FILE.write_text(json.dumps(existing, indent=2, ensure_ascii=False))
        except Exception as e:
            self.sig_log.emit(f"Log save error: {e}", "warn")

    def run(self):
        cfg     = self.cfg
        mailer  = EmailSender(cfg["email"], cfg["password"])
        delay_s = cfg["delay_min"] * 60
        sent = failed = 0

        if self.drafts is not None:
            total = len(self.drafts)
            self._log(f"Sending {total} pre-reviewed email(s), {cfg['delay_min']} min delay")
            self.sig_progress.emit(0, total)
            for i, draft in enumerate(self.drafts):
                if self._check_paused():
                    self._log("Stopped by user.", "warn"); break
                self.sig_status.emit(f"Sending to {draft.name} <{draft.email}>…")
                try:
                    mailer.send(
                        to=draft.email, subject=draft.subject, body=draft.body,
                        cv_path=self.cv_path if cfg["attach_cv"] else None,
                    )
                    self._log(
                        f"[{i+1}/{total}] ✓ Sent → {draft.name} <{draft.email}> | {draft.subject}",
                        "success", draft.email, True,
                    )
                    sent += 1
                except smtplib.SMTPAuthenticationError:
                    self._log("SMTP auth failed — check credentials.", "error", draft.email, False)
                    failed += 1; break
                except (smtplib.SMTPException, OSError) as exc:
                    self._log(f"[{i+1}/{total}] SMTP error: {exc}", "error", draft.email, False)
                    failed += 1
                except Exception as exc:
                    self._log(f"[{i+1}/{total}] Error: {exc}", "error", draft.email, False)
                    failed += 1
                self.sig_progress.emit(i + 1, total)
                if i < total - 1 and not self._stopped:
                    remaining = delay_s
                    while remaining > 0 and not self._stopped:
                        if self._check_paused(): break
                        tick = min(remaining, 5)
                        time.sleep(tick); remaining -= tick
                        m, s = divmod(int(remaining), 60)
                        self.sig_status.emit(f"Next email in {m}m {s:02d}s…")
            self._flush_log()
            summary = f"Done: {sent} sent, {failed} failed / {total} total."
            self._log(summary); self.sig_status.emit("Idle")
            self.sig_done.emit(failed == 0, summary)
            return

        llm   = LLMClient(cfg["provider"], cfg["api_key"], cfg.get("model", ""))
        total = len(self.contacts)
        sent = failed = 0

        self._log(f"Campaign started — {total} contact(s), {cfg['delay_min']} min delay")
        self.sig_progress.emit(0, total)

        for i, contact in enumerate(self.contacts):
            if self._check_paused():
                self._log("Stopped by user.", "warn")
                break

            email = get_field(contact, "email", "e-mail", "email_address", "mail")
            name  = get_field(contact, "name", "full_name", "first_name",
                              default=f"Contact {i + 1}")

            if not email:
                self._log(f"[{i+1}/{total}] No email for '{name}' — skipped", "warn")
                self.sig_progress.emit(i + 1, total)
                continue
            if not valid_email(email):
                self._log(f"[{i+1}/{total}] Invalid address '{email}' — skipped", "warn", email)
                failed += 1
                self.sig_progress.emit(i + 1, total)
                continue

            company_context = ""
            if cfg.get("research"):
                company = get_field(contact, "company", "organization", "employer", "firm",
                                    default="")
                website = get_field(contact, "website", "url", "homepage", "site")
                if company:
                    self.sig_status.emit(f"Researching {company}…")
                    self._log(f"[{i+1}/{total}] Researching company: {company}", "info", email)
                    company_context = self._researcher.research(company, website)
                    if company_context:
                        self._log(
                            f"[{i+1}/{total}] Context found ({len(company_context)} chars)",
                            "info", email,
                        )
                    else:
                        self._log(f"[{i+1}/{total}] No context found for '{company}'", "warn", email)

            self.sig_status.emit(f"Generating email for {name}…")
            self._log(f"[{i+1}/{total}] Generating → {name} <{email}>", "info", email)
            try:
                raw = llm.generate(
                    build_prompt(
                        self.cv_text, contact, cfg["goal"], company_context,
                        sender_name=cfg.get("sender_name", ""),
                        sender_email=cfg.get("sender_email", ""),
                        sender_phone=cfg.get("sender_phone", ""),
                        sender_extra=cfg.get("sender_extra", ""),
                        template_text=cfg.get("template_text", ""),
                    )
                )
                subject, body = parse_llm_json(raw)
                if not body.strip():
                    raise ValueError("LLM returned empty body")
                subject = subject or cfg["goal"]
            except Exception as exc:
                self._log(f"[{i+1}/{total}] LLM error: {exc}", "error", email, False)
                failed += 1
                self.sig_progress.emit(i + 1, total)
                continue

            time.sleep(3)

            if cfg.get("preview"):
                self._preview_event.clear()
                self.sig_preview.emit(name, email, subject, body)
                self._preview_event.wait(timeout=300)
                send, subject, body = self._preview_result
                if not send:
                    self._log(f"[{i+1}/{total}] Skipped by user: {email}", "warn", email)
                    self.sig_progress.emit(i + 1, total)
                    continue

            self.sig_status.emit(f"Sending to {name} <{email}>…")
            try:
                mailer.send(
                    to=email, subject=subject, body=body,
                    cv_path=self.cv_path if cfg["attach_cv"] else None,
                )
                self._log(
                    f"[{i+1}/{total}] ✓ Sent → {name} <{email}> | {subject}",
                    "success", email, True,
                )
                sent += 1
            except smtplib.SMTPAuthenticationError:
                self._log("SMTP auth failed — check credentials.", "error", email, False)
                failed += 1
                break
            except (smtplib.SMTPException, OSError) as exc:
                self._log(f"[{i+1}/{total}] SMTP error: {exc}", "error", email, False)
                failed += 1
            except Exception as exc:
                self._log(f"[{i+1}/{total}] Send error: {exc}", "error", email, False)
                failed += 1

            self.sig_progress.emit(i + 1, total)

            if i < total - 1 and not self._stopped:
                remaining = delay_s
                while remaining > 0 and not self._stopped:
                    if self._check_paused():
                        break
                    tick = min(remaining, 5)
                    time.sleep(tick)
                    remaining -= tick
                    m, s = divmod(int(remaining), 60)
                    self.sig_status.emit(f"Next email in {m}m {s:02d}s…")

        self._flush_log()
        summary = f"Done: {sent} sent, {failed} failed / {total} total."
        self._log(summary)
        self.sig_status.emit("Idle")
        self.sig_done.emit(failed == 0, summary)


# ── Main Window ────────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} — Email Outreach")
        self.setMinimumSize(1020, 700)
        self.resize(1200, 800)

        self._cv_path       = ""
        self._csv_path      = ""
        self._template_path = ""
        self._template_text = ""
        self._contacts: List[Dict] = []
        self._cv_text   = ""
        self._thread: Optional[CampaignThread] = None
        self._gen_thread: Optional[GenerationThread] = None
        self._cv_info_thread: Optional[CVInfoThread] = None
        self._is_paused = False
        self._sender_name  = ""
        self._sender_phone = ""
        self._sender_extra = ""
        self._cv_llm_extracted = False

        self._build_ui()
        self._load_settings()
        self._connect_save_signals()

    # ── Persistence ───────────────────────────────────────────────────────────
    def _save_settings(self):
        data = {
            "goal":          self.inp_goal.text().strip(),
            "provider":      self.cmb_provider.currentText(),
            "model":         self.inp_model.text().strip(),
            "api_key":       self.inp_api_key.text().strip(),
            "sender_email":  self.inp_sender.text().strip(),
            "app_password":  self.inp_password.text().strip(),
            "delay_min":     self.spn_delay.value(),
            "attach_cv":     self.chk_attach.isChecked(),
            "preview":       self.chk_preview.isChecked(),
            "research":      self.chk_research.isChecked(),
            "cv_path":       self._cv_path,
            "csv_path":      self._csv_path,
            "template_path": self._template_path,
        }
        try:
            SETTINGS_FILE.write_text(json.dumps(data, indent=2, ensure_ascii=False))
        except Exception:
            pass

    def _load_settings(self):
        if not SETTINGS_FILE.exists():
            return
        try:
            data = json.loads(SETTINGS_FILE.read_text())
        except Exception:
            return

        if data.get("goal"):
            self.inp_goal.setText(data["goal"])
        if data.get("provider") and data["provider"] in [
            self.cmb_provider.itemText(i) for i in range(self.cmb_provider.count())
        ]:
            self.cmb_provider.setCurrentText(data["provider"])
        if data.get("model"):
            self.inp_model.setText(data["model"])
        if data.get("api_key"):
            self.inp_api_key.setText(data["api_key"])
        if data.get("sender_email"):
            self.inp_sender.setText(data["sender_email"])
        if data.get("app_password"):
            self.inp_password.setText(data["app_password"])
        if data.get("delay_min"):
            self.spn_delay.setValue(int(data["delay_min"]))
        self.chk_attach.setChecked(bool(data.get("attach_cv", False)))
        self.chk_preview.setChecked(bool(data.get("preview", False)))
        self.chk_research.setChecked(bool(data.get("research", False)))

        for attr, lbl_attr, loader, path_key in [
            ("_cv_path",       "lbl_cv",       self._load_cv_file,       "cv_path"),
            ("_csv_path",      "lbl_csv",       self._load_csv_file,      "csv_path"),
            ("_template_path", "lbl_template",  self._load_template_file, "template_path"),
        ]:
            path = data.get(path_key, "")
            if path and Path(path).exists():
                loader(path)

    def _load_cv_file(self, path: str):
        self._cv_path = path
        self._cv_text = parse_cv(path)
        name = Path(path).name
        self.lbl_cv.setText(name)
        self.lbl_cv.setStyleSheet(f"color: {C['text']}; font-size: 12px;")
        if not self._cv_text.startswith("["):
            self._cv_llm_extracted = False
            self._start_cv_extraction()

    def _load_csv_file(self, path: str):
        try:
            self._csv_path = path
            self._contacts = load_contacts(path)
            name = Path(path).name
            self.lbl_csv.setText(name)
            self.lbl_csv.setStyleSheet(f"color: {C['text']}; font-size: 12px;")
            n = len(self._contacts)
            self.lbl_count.setText(f"✓  {n} contact{'s' if n != 1 else ''} loaded")
            self.progress.setRange(0, n)
            self.progress.setFormat(f"0 / {n}")
        except Exception:
            pass

    def _load_template_file(self, path: str):
        text = load_template(path)
        if text.startswith("["):
            return
        self._template_path = path
        self._template_text = text
        self.lbl_template.setText(Path(path).name)
        self.lbl_template.setStyleSheet(f"color: {C['text']}; font-size: 12px;")

    def _connect_save_signals(self):
        self.inp_goal.editingFinished.connect(self._save_settings)
        self.inp_api_key.editingFinished.connect(self._save_settings)
        self.inp_sender.editingFinished.connect(self._save_settings)
        self.inp_password.editingFinished.connect(self._save_settings)
        self.inp_model.editingFinished.connect(self._save_settings)
        self.cmb_provider.currentTextChanged.connect(lambda _: self._save_settings())
        self.spn_delay.valueChanged.connect(lambda _: self._save_settings())
        self.chk_attach.toggled.connect(lambda _: self._save_settings())
        self.chk_preview.toggled.connect(lambda _: self._save_settings())
        self.chk_research.toggled.connect(lambda _: self._save_settings())

    # ── Build UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        vbox = QVBoxLayout(root)
        vbox.setSpacing(0)
        vbox.setContentsMargins(0, 0, 0, 0)

        vbox.addWidget(self._ui_header())

        body = QWidget()
        body_lay = QVBoxLayout(body)
        body_lay.setSpacing(10)
        body_lay.setContentsMargins(14, 12, 14, 10)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(1)
        splitter.addWidget(self._ui_left())
        splitter.addWidget(self._ui_right())
        splitter.setSizes([420, 680])
        body_lay.addWidget(splitter, 1)
        body_lay.addWidget(self._ui_progress_bar())

        vbox.addWidget(body, 1)

    def _ui_header(self) -> QWidget:
        w = QWidget()
        w.setFixedHeight(64)
        w.setStyleSheet(
            f"background: qlineargradient(x1:0,y1:0,x2:1,y2:0,"
            f"stop:0 {C['mantle']}, stop:1 {C['surface0']});"
            f"border-bottom: 1px solid {C['surface0']};"
        )
        h = QHBoxLayout(w)
        h.setContentsMargins(18, 0, 18, 0)

        # Title group
        col = QVBoxLayout()
        col.setSpacing(2)
        title = QLabel(APP_NAME)
        title.setObjectName("lblTitle")
        title.setFont(QFont("Kufam", 20, QFont.Weight.Bold))
        title.setStyleSheet(f"font-size: 20px; font-weight: bold; color: {C['blue']}; background: transparent;")
        sub = QLabel("LLM-powered personalized email campaigns")
        sub.setObjectName("lblSubtitle")
        sub.setStyleSheet(f"font-size: 11px; color: {C['overlay0']}; background: transparent;")
        col.addWidget(title)
        col.addWidget(sub)
        h.addLayout(col)
        h.addStretch()

        # Status pill
        self.lbl_status_pill = QLabel("● Idle")
        self.lbl_status_pill.setObjectName("lblBadge")
        self.lbl_status_pill.setStyleSheet(
            f"background-color: {C['surface0']}; border: 1px solid {C['surface1']};"
            f"border-radius: 12px; padding: 4px 14px; color: {C['overlay2']}; font-size: 12px;"
        )
        h.addWidget(self.lbl_status_pill)
        h.addSpacing(12)

        self.btn_test = QPushButton("⚡  Test SMTP")
        self.btn_test.setObjectName("btnTest")
        self.btn_test.setFixedHeight(36)
        self.btn_test.clicked.connect(self._send_test)
        h.addWidget(self.btn_test)
        return w

    # ── Left panel ────────────────────────────────────────────────────────────
    def _ui_left(self) -> QScrollArea:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        inner = QWidget()
        lay = QVBoxLayout(inner)
        lay.setSpacing(14)
        lay.setContentsMargins(2, 4, 10, 4)
        lay.addWidget(self._grp_files())
        lay.addWidget(self._grp_config())
        lay.addWidget(self._grp_settings())
        lay.addWidget(self._grp_controls())
        lay.addStretch()

        scroll.setWidget(inner)
        return scroll

    def _file_row(self, label: str, lbl_attr: str, slot) -> QWidget:
        row = QWidget()
        h = QHBoxLayout(row)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(8)
        lbl_label = QLabel(label)
        lbl_label.setStyleSheet(f"color: {C['subtext']}; font-size: 12px;")
        lbl_label.setFixedWidth(88)
        h.addWidget(lbl_label)
        lbl = QLabel("No file selected")
        lbl.setStyleSheet(f"color: {C['overlay0']}; font-style: italic; font-size: 12px;")
        lbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        setattr(self, lbl_attr, lbl)
        h.addWidget(lbl, 1)
        btn = QPushButton("Browse")
        btn.setObjectName("btnBrowse")
        btn.setFixedWidth(80)
        btn.setFixedHeight(30)
        btn.clicked.connect(slot)
        h.addWidget(btn)
        return row

    def _grp_files(self) -> QGroupBox:
        grp = QGroupBox("📁  Files")
        lay = QVBoxLayout(grp)
        lay.setSpacing(10)
        lay.setContentsMargins(12, 8, 12, 12)
        lay.addWidget(self._file_row("CV / Resume:", "lbl_cv",  self._browse_cv))

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet(f"background: {C['surface0']}; max-height: 1px;")
        lay.addWidget(sep)

        lay.addWidget(self._file_row("Contacts CSV:", "lbl_csv", self._browse_csv))

        self.lbl_count = QLabel("")
        self.lbl_count.setStyleSheet(f"color: {C['green']}; font-size: 11px; padding-left: 96px;")
        lay.addWidget(self.lbl_count)

        sep2 = QFrame()
        sep2.setFrameShape(QFrame.Shape.HLine)
        sep2.setStyleSheet(f"background: {C['surface0']}; max-height: 1px;")
        lay.addWidget(sep2)

        lay.addWidget(self._file_row("Template:", "lbl_template", self._browse_template))

        lbl_tmpl_hint = QLabel("Optional — .txt, .pdf or .docx  ·  LLM will follow its structure")
        lbl_tmpl_hint.setStyleSheet(
            f"color: {C['overlay0']}; font-size: 11px; padding-left: 96px; font-style: italic;"
        )
        lay.addWidget(lbl_tmpl_hint)
        return grp

    def _grp_config(self) -> QGroupBox:
        grp = QGroupBox("🔑  Configuration")
        form = QFormLayout(grp)
        form.setSpacing(10)
        form.setContentsMargins(12, 8, 12, 12)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.inp_goal = QLineEdit()
        self.inp_goal.setPlaceholderText("e.g. summer AI / Edge AI engineering internship")
        form.addRow("Goal:", self.inp_goal)

        self.cmb_provider = QComboBox()
        self.cmb_provider.addItems([
            "OpenAI", "Gemini", "Groq", "OpenRouter", "Anthropic",
            "Mistral", "Mistral Codestral", "Together AI", "Cohere",
            "Cerebras", "NVIDIA NIM", "GitHub Models", "Fireworks",
            "Nebius", "Novita", "Hyperbolic", "SambaNova", "Scaleway",
            "Inference.net",
        ])
        self.cmb_provider.currentTextChanged.connect(self._on_provider)
        form.addRow("LLM Provider:", self.cmb_provider)

        self.inp_model = QLineEdit()
        self.inp_model.setPlaceholderText("gpt-4o-mini")
        self.inp_model.setText("gpt-4o-mini")
        form.addRow("Model:", self.inp_model)

        self.inp_api_key = QLineEdit()
        self.inp_api_key.setEchoMode(QLineEdit.EchoMode.Password)
        self.inp_api_key.setPlaceholderText("Paste your API key…")
        form.addRow("API Key:", self.inp_api_key)

        sep_lbl = QLabel("── SMTP ──────────────────────────────")
        sep_lbl.setStyleSheet(f"color: {C['surface2']}; font-size: 11px;")
        form.addRow("", sep_lbl)

        self.inp_sender = QLineEdit()
        self.inp_sender.setPlaceholderText("your@gmail.com")
        form.addRow("Gmail:", self.inp_sender)

        self.inp_password = QLineEdit()
        self.inp_password.setEchoMode(QLineEdit.EchoMode.Password)
        self.inp_password.setPlaceholderText("16-char App Password")
        form.addRow("App Password:", self.inp_password)

        return grp

    def _grp_settings(self) -> QGroupBox:
        grp = QGroupBox("⚙️  Settings")
        form = QFormLayout(grp)
        form.setSpacing(10)
        form.setContentsMargins(12, 8, 12, 12)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        delay_w = QWidget()
        dh = QHBoxLayout(delay_w)
        dh.setContentsMargins(0, 0, 0, 0)
        self.spn_delay = QSpinBox()
        self.spn_delay.setRange(1, 120)
        self.spn_delay.setValue(10)
        self.spn_delay.setSuffix(" min")
        self.spn_delay.setFixedWidth(100)
        dh.addWidget(self.spn_delay)
        dh.addStretch()
        form.addRow("Send delay:", delay_w)

        self.chk_attach   = QCheckBox("Attach CV to every email")
        self.chk_preview  = QCheckBox("Preview each email before sending")
        self.chk_research = QCheckBox("Research each company online before writing")
        self.chk_research.setStyleSheet(f"color: {C['sky']};")
        form.addRow("", self.chk_attach)
        form.addRow("", self.chk_preview)
        form.addRow("", self.chk_research)
        return grp

    def _grp_controls(self) -> QGroupBox:
        grp = QGroupBox("▶  Campaign")
        outer = QVBoxLayout(grp)
        outer.setContentsMargins(12, 8, 12, 12)
        outer.setSpacing(8)

        top_row = QHBoxLayout()
        self.btn_preview_all = QPushButton("📧  Generate Drafts")
        self.btn_preview_all.setObjectName("btnGenerate")
        self.btn_preview_all.setMinimumHeight(44)
        self.btn_preview_all.setToolTip(
            "Generate all emails with the LLM first, review them, then send."
        )
        self.btn_preview_all.clicked.connect(self._generate_drafts)

        self.btn_start = QPushButton("▶  Start Campaign")
        self.btn_start.setObjectName("btnStart")
        self.btn_start.setMinimumHeight(44)
        self.btn_start.setToolTip("Generate and send emails one by one (no preview).")
        self.btn_start.clicked.connect(self._start)
        top_row.addWidget(self.btn_preview_all)
        top_row.addWidget(self.btn_start)
        outer.addLayout(top_row)

        bot_row = QHBoxLayout()
        self.btn_pause = QPushButton("⏸  Pause")
        self.btn_pause.setObjectName("btnPause")
        self.btn_pause.setMinimumHeight(36)
        self.btn_pause.setEnabled(False)
        self.btn_pause.clicked.connect(self._toggle_pause)

        self.btn_stop = QPushButton("⏹  Stop")
        self.btn_stop.setObjectName("btnStop")
        self.btn_stop.setMinimumHeight(36)
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self._stop)
        bot_row.addWidget(self.btn_pause)
        bot_row.addWidget(self.btn_stop)
        outer.addLayout(bot_row)

        return grp

    # ── Right panel ───────────────────────────────────────────────────────────
    def _ui_right(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        v.setSpacing(10)
        v.setContentsMargins(10, 4, 2, 4)

        # Status bar
        sr = QWidget()
        sr.setStyleSheet(
            f"background-color: {C['surface0']}; border-radius: 8px;"
        )
        sh = QHBoxLayout(sr)
        sh.setContentsMargins(14, 8, 14, 8)
        lbl = QLabel("Status")
        lbl.setStyleSheet(f"font-weight: 600; color: {C['subtext']}; font-size: 12px; background: transparent;")
        sh.addWidget(lbl)
        self.lbl_status = QLabel("Idle — configure inputs and press Start")
        self.lbl_status.setObjectName("lblStatus")
        self.lbl_status.setWordWrap(True)
        self.lbl_status.setStyleSheet(f"color: {C['green']}; font-style: italic; background: transparent;")
        sh.addWidget(self.lbl_status, 1)
        v.addWidget(sr)

        # Log
        log_grp = QGroupBox("📋  Activity Log")
        lv = QVBoxLayout(log_grp)
        lv.setContentsMargins(8, 8, 8, 8)
        lv.setSpacing(6)

        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        lv.addWidget(self.log_view)

        footer = QHBoxLayout()
        self.lbl_log_count = QLabel("0 entries")
        self.lbl_log_count.setStyleSheet(f"color: {C['overlay0']}; font-size: 11px;")
        footer.addWidget(self.lbl_log_count)
        footer.addStretch()
        btn_clear = QPushButton("Clear")
        btn_clear.setFixedWidth(72)
        btn_clear.setFixedHeight(28)
        btn_clear.clicked.connect(self._clear_log)
        footer.addWidget(btn_clear)
        lv.addLayout(footer)
        v.addWidget(log_grp, 1)
        return w

    def _ui_progress_bar(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 2, 0, 0)
        v.setSpacing(0)
        self.progress = QProgressBar()
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.progress.setFormat("Ready")
        self.progress.setFixedHeight(26)
        v.addWidget(self.progress)
        return w

    # ── Slots ─────────────────────────────────────────────────────────────────
    def _browse_cv(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select CV / Resume", "",
            "Documents (*.pdf *.docx *.doc);;All Files (*)",
        )
        if not path:
            return
        self._load_cv_file(path)
        if self._cv_text.startswith("["):
            self._append_log(f"CV parse warning: {self._cv_text}", "warn")
        else:
            self._append_log(f"CV loaded: {Path(path).name} ({len(self._cv_text)} chars)", "success")
        self._save_settings()

    def _browse_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Email Template", "",
            "Text / Documents (*.txt *.pdf *.docx *.doc);;All Files (*)",
        )
        if not path:
            return
        text = load_template(path)
        if text.startswith("["):
            self._append_log(f"Template warning: {text}", "warn")
            return
        self._load_template_file(path)
        self._append_log(f"Template loaded: {Path(path).name} ({len(text)} chars)", "success")
        self._save_settings()

    def _browse_csv(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Contacts CSV", "",
            "CSV Files (*.csv);;All Files (*)",
        )
        if not path:
            return
        try:
            self._load_csv_file(path)
            self.progress.setValue(0)
            n = len(self._contacts)
            self._append_log(f"CSV loaded: {Path(path).name} — {n} contacts", "success")
            self._save_settings()
        except Exception as e:
            QMessageBox.critical(self, "CSV Error", f"Failed to load CSV:\n{e}")

    def _start_cv_extraction(self):
        if self._cv_info_thread and self._cv_info_thread.isRunning():
            return
        self._cv_info_thread = CVInfoThread(
            self._cv_text,
            provider=self.cmb_provider.currentText(),
            api_key=self.inp_api_key.text().strip(),
            model=self.inp_model.text().strip(),
        )
        self._cv_info_thread.sig_done.connect(self._on_cv_info)
        self._cv_info_thread.start()

    def _on_cv_info(self, info: dict):
        self._cv_llm_extracted = True
        self._sender_name  = info.get("name",  "").strip()
        self._sender_phone = info.get("phone", "").strip()
        self._sender_extra = info.get("extra", "").strip()

        extracted_email = info.get("email", "").strip()
        if extracted_email and not self.inp_sender.text().strip():
            self.inp_sender.setText(extracted_email)

        found = [k for k in ("name", "phone", "email", "extra") if info.get(k)]
        if found:
            self._append_log(
                f"Sender info extracted from CV: {', '.join(found)}", "success"
            )
        else:
            self._append_log("Could not extract sender info from CV (LLM will derive from CV text).", "warn")

    _PROVIDER_DEFAULTS = {
        "OpenAI":            "gpt-4o-mini",
        "Gemini":            "gemini-2.0-flash",
        "Groq":              "llama-3.3-70b-versatile",
        "OpenRouter":        "google/gemini-2.0-flash-exp:free",
        "Anthropic":         "claude-3-5-haiku-20241022",
        "Mistral":           "mistral-small-latest",
        "Mistral Codestral": "codestral-latest",
        "Together AI":       "meta-llama/Llama-3.3-70B-Instruct-Turbo",
        "Cohere":            "command-r-plus-08-2024",
        "Cerebras":          "llama-3.3-70b",
        "NVIDIA NIM":        "meta/llama-3.3-70b-instruct",
        "GitHub Models":     "gpt-4o-mini",
        "Fireworks":         "accounts/fireworks/models/llama-v3p3-70b-instruct",
        "Nebius":            "meta-llama/Llama-3.3-70B-Instruct",
        "Novita":            "meta-llama/llama-3.3-70b-instruct",
        "Hyperbolic":        "meta-llama/Llama-3.3-70B-Instruct",
        "SambaNova":         "Meta-Llama-3.3-70B-Instruct",
        "Scaleway":          "llama-3.3-70b-instruct",
        "Inference.net":     "meta-llama/llama-3.3-70b-instruct/fp-8",
    }

    def _on_provider(self, text: str):
        default = self._PROVIDER_DEFAULTS.get(text, "")
        self.inp_model.setPlaceholderText(default)
        if not self.inp_model.text().strip():
            self.inp_model.setText(default)

    def _validate(self) -> Optional[str]:
        if not self._cv_path:
            return "Please select a CV / Resume file."
        if not self._contacts:
            return "Please load a contacts CSV file."
        if not self.inp_goal.text().strip():
            return "Please enter the email goal."
        if not self.inp_api_key.text().strip():
            return "Please enter the LLM API key."
        if not self.inp_sender.text().strip():
            return "Please enter your Gmail address."
        if not valid_email(self.inp_sender.text().strip()):
            return "The Gmail address is not valid."
        if not self.inp_password.text().strip():
            return "Please enter the Gmail App Password."
        return None

    def _build_cfg(self) -> Dict[str, Any]:
        return {
            "goal":         self.inp_goal.text().strip(),
            "provider":     self.cmb_provider.currentText(),
            "api_key":      self.inp_api_key.text().strip(),
            "model":        self.inp_model.text().strip(),
            "email":        self.inp_sender.text().strip(),
            "password":     self.inp_password.text().strip(),
            "delay_min":    self.spn_delay.value(),
            "attach_cv":    self.chk_attach.isChecked(),
            "preview":      self.chk_preview.isChecked(),
            "research":     self.chk_research.isChecked(),
            "sender_name":   self._sender_name,
            "sender_email":  self.inp_sender.text().strip(),
            "sender_phone":  self._sender_phone,
            "sender_extra":  self._sender_extra,
            "template_text": self._template_text,
        }

    def _ensure_cv_extracted(self):
        if self._cv_text and not self._cv_llm_extracted and self.inp_api_key.text().strip():
            self._append_log("Extracting sender info from CV via LLM…", "info")
            self._start_cv_extraction()
            if self._cv_info_thread:
                self._cv_info_thread.wait()

    def _start(self):
        err = self._validate()
        if err:
            QMessageBox.warning(self, "Missing Input", err)
            return

        self._ensure_cv_extracted()
        cfg = self._build_cfg()
        self.progress.setValue(0)
        self._is_paused = False
        self._set_running(True)

        self._thread = CampaignThread(
            cfg, self._contacts, self._cv_text, self._cv_path, self
        )
        self._thread.sig_log.connect(self._on_log)
        self._thread.sig_progress.connect(self._on_progress)
        self._thread.sig_status.connect(self._on_status)
        self._thread.sig_preview.connect(self._on_preview)
        self._thread.sig_done.connect(self._on_done)
        self._thread.start()

    def _toggle_pause(self):
        if not self._thread:
            return
        if self._is_paused:
            self._thread.resume()
            self._is_paused = False
            self.btn_pause.setText("⏸  Pause")
            self._on_status("Resumed…")
        else:
            self._thread.pause()
            self._is_paused = True
            self.btn_pause.setText("▶  Resume")
            self._on_status("Paused — click Resume to continue")

    def _generate_drafts(self):
        err = self._validate()
        if err:
            QMessageBox.warning(self, "Missing Input", err)
            return

        self._ensure_cv_extracted()
        cfg = self._build_cfg()
        cfg["preview"] = False
        self.progress.setValue(0)
        self._set_running(True)
        self._on_status("Generating drafts — please wait…")

        self._gen_thread = GenerationThread(cfg, self._contacts, self._cv_text, self)
        self._gen_thread.sig_log.connect(self._on_log)
        self._gen_thread.sig_progress.connect(self._on_progress)
        self._gen_thread.sig_status.connect(self._on_status)
        self._gen_thread.sig_done.connect(self._on_gen_done)
        self._gen_thread.start()

    def _on_gen_done(self, drafts: list):
        self._set_running(False)
        self._on_status("Idle")
        if not drafts:
            QMessageBox.warning(self, "No Drafts", "No valid drafts were generated.")
            return

        dlg = PreviewAllDialog(drafts, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            selected = dlg.selected_drafts()
            if selected:
                self._send_drafts(selected)

    def _send_drafts(self, drafts: List[DraftEmail]):
        cfg = self._build_cfg()
        cfg["preview"] = False
        cfg["research"] = False
        self.progress.setRange(0, len(drafts))
        self.progress.setValue(0)
        self._set_running(True)

        self._thread = CampaignThread(
            cfg, [], self._cv_text, self._cv_path, self, drafts=drafts
        )
        self._thread.sig_log.connect(self._on_log)
        self._thread.sig_progress.connect(self._on_progress)
        self._thread.sig_status.connect(self._on_status)
        self._thread.sig_done.connect(self._on_done)
        self._thread.start()

    def _stop(self):
        if self._thread:
            self._thread.stop()
            self._thread.set_preview_result(False, "", "")
        if self._gen_thread:
            self._gen_thread.stop()
        self._on_status("Stopping…")

    # ── Thread callbacks ──────────────────────────────────────────────────────
    def _on_log(self, msg: str, level: str):
        self._append_log(msg, level)

    def _on_progress(self, current: int, total: int):
        self.progress.setRange(0, total)
        self.progress.setValue(current)
        self.progress.setFormat(f"{current} / {total}")

    def _on_status(self, msg: str):
        self.lbl_status.setText(msg)
        self.lbl_status_pill.setText(f"● {msg[:48]}")

    def _on_preview(self, name: str, email: str, subject: str, body: str):
        dlg = PreviewDialog(name, email, subject, body, self)
        result = dlg.exec()
        if self._thread:
            if result == QDialog.DialogCode.Accepted:
                self._thread.set_preview_result(True, dlg.subject, dlg.body)
            else:
                self._thread.set_preview_result(False, "", "")

    def _on_done(self, success: bool, summary: str):
        self._set_running(False)
        self._append_log(summary, "success" if success else "warn")
        self._on_status("Idle")
        QMessageBox.information(self, "Campaign Complete", summary)

    # ── Test email ────────────────────────────────────────────────────────────
    def _send_test(self):
        sender = self.inp_sender.text().strip()
        pwd    = self.inp_password.text().strip()
        if not sender or not pwd:
            QMessageBox.warning(
                self, "Missing Credentials",
                "Enter your Gmail address and App Password first.",
            )
            return
        try:
            EmailSender(sender, pwd).send(
                to=sender,
                subject="[Test] Email Outreach App",
                body=(
                    "This is a test email from your Email Outreach Automation app.\n\n"
                    "SMTP credentials are working correctly!"
                ),
            )
            self._append_log(f"Test email sent to {sender}", "success")
            QMessageBox.information(self, "Test Email Sent",
                                    f"Test email delivered to {sender}")
        except Exception as e:
            self._append_log(f"Test email failed: {e}", "error")
            QMessageBox.critical(self, "Test Failed", str(e))

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _set_running(self, running: bool):
        self.btn_preview_all.setEnabled(not running)
        self.btn_start.setEnabled(not running)
        self.btn_pause.setEnabled(running)
        self.btn_stop.setEnabled(running)
        if not running:
            self.btn_pause.setText("⏸  Pause")
            self._is_paused = False
            self.lbl_status_pill.setStyleSheet(
                f"background-color: {C['surface0']}; border: 1px solid {C['surface1']};"
                f"border-radius: 12px; padding: 4px 14px; color: {C['overlay2']}; font-size: 12px;"
            )
        else:
            self.lbl_status_pill.setStyleSheet(
                f"background-color: rgba(166,227,161,0.15); border: 1px solid {C['green']};"
                f"border-radius: 12px; padding: 4px 14px; color: {C['green']}; font-size: 12px;"
            )

    _log_entry_count = 0

    def _clear_log(self):
        self.log_view.clear()
        self._log_entry_count = 0
        self.lbl_log_count.setText("0 entries")

    def _append_log(self, msg: str, level: str = "info"):
        colors = {
            "info":    C['text'],
            "success": C['green'],
            "warn":    C['yellow'],
            "error":   C['red'],
        }
        icons = {
            "info":    "·",
            "success": "✓",
            "warn":    "⚠",
            "error":   "✗",
        }
        color = colors.get(level, C['text'])
        icon  = icons.get(level, "·")
        ts = datetime.now().strftime("%H:%M:%S")
        safe = msg.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        html = (
            f'<span style="color:{C["surface2"]}">[{ts}]</span> '
            f'<span style="color:{color}">{icon} {safe}</span>'
        )
        self.log_view.append(html)
        cur = self.log_view.textCursor()
        cur.movePosition(QTextCursor.MoveOperation.End)
        self.log_view.setTextCursor(cur)
        self._log_entry_count += 1
        self.lbl_log_count.setText(f"{self._log_entry_count} entr{'y' if self._log_entry_count == 1 else 'ies'}")

    def closeEvent(self, event):
        self._save_settings()
        if self._cv_info_thread and self._cv_info_thread.isRunning():
            self._cv_info_thread.wait(2000)
        active = (
            (self._thread and self._thread.isRunning()) or
            (self._gen_thread and self._gen_thread.isRunning())
        )
        if active:
            reply = QMessageBox.question(
                self, "Task Running",
                "A generation or campaign is in progress. Stop it and close?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.No:
                event.ignore()
                return
            if self._gen_thread:
                self._gen_thread.stop()
                self._gen_thread.wait(3000)
            if self._thread:
                self._thread.stop()
                self._thread.set_preview_result(False, "", "")
                self._thread.wait(5000)
        event.accept()


# ── Entry Point ────────────────────────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setStyleSheet(DARK_STYLE)
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
