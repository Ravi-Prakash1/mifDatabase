#!/usr/bin/env python3
"""
mif_common.py — shared configuration and helpers for the MIF portal automation.

Both fetch_news.py and enrich_data.py import from here so that file paths, the
Excel schema contract, the de-duplication cache, provenance/flagging rules and
the LLM provider layer all live in exactly one place.

Design rules that the rest of the automation depends on:
  * The Excel files stay the single source of truth. Scripts only APPEND
    (news editions) or FILL EMPTY cells (enrichment) — never overwrite a cell
    a human has verified.
  * Every machine-written value carries a source URL and is flagged
    "Estimated" so a reviewer can tell auto-filled data from verified data.
  * Nothing here prints secrets. The free LLM key (GROQ_API_KEY or
    GEMINI_API_KEY) is read from the environment only.
"""

import os
import re
import json
import hashlib
from datetime import datetime, timezone

# ─── PATHS ────────────────────────────────────────────────────────────────────
# Everything is resolved relative to the repository root (the parent of this
# scripts/ directory) so the scripts work the same locally and in CI.
REPO_ROOT   = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

PROSPECTS_FILE   = os.path.join(REPO_ROOT, "MIF_Prospects_Master_Database.xlsx")
EXHIBITIONS_FILE = os.path.join(REPO_ROOT, "MIF_Global_Exhibitions_FY_2026_27.xlsx")
NEWS_FILE        = os.path.join(REPO_ROOT, "MIF_News_Intelligence.xlsx")

DATA_DIR         = os.path.join(REPO_ROOT, "data")
NEWS_SEEN_FILE   = os.path.join(DATA_DIR, "news_seen.json")

# ─── EXCEL SCHEMA CONTRACT ────────────────────────────────────────────────────
# These MUST match what generate_portal_v7_News.py::parse_news() reads. Do not
# reorder or rename without updating the generator too.
NEWS_SHEETS = {
    "Briefing":   ["Date", "Type", "Input Cost Note"],
    "Indicators": ["Date", "Indicator", "Value", "Unit", "Movement", "Driver", "Source"],
    "Trend":      ["Date", "HR Coil", "USD/INR"],
    "News":       ["Date", "Section", "Segment", "Status", "Company", "Headline",
                   "What Happened", "Why It Matters", "Source", "Importance",
                   "Marquee", "Dashboard"],
}

# Prospects "Master Database" sheet — column header -> 0-based index (56 cols).
PROSPECTS_SHEET = "Master Database"
AUDIT_SHEET     = "Fill Audit Log"          # cols: Company, Field, New Value, Method, Source, Date
AUDIT_COLS      = ["Company", "Field", "New Value", "Method", "Source", "Date"]

# Provenance / quality columns in the prospects sheet.
COL_DATA_CONFIDENCE = "Data Confidence"
COL_LAST_VERIFIED   = "Last Verified Date"
COL_VERIF_METHOD    = "Verification Method"
COL_SOURCE_LINKS    = "Source Hyperlink(s)"
COL_EST_FLAG        = "Estimated vs Verified Flag"
COL_COMPANY_NAME    = "Company Name"

# Fields enrichment is allowed to fill (sparse, high-value). Anything not in this
# list is left untouched even if empty.
ENRICHABLE_PROSPECT_FIELDS = [
    "CIN", "GSTIN", "Legal Form", "Year Est.", "Parent / Group", "HQ Address",
    "Plant Locations", "Company Website", "Products", "Director Name",
    "Procurement Head", "Sales / BD Head", "Revenue (Rs. Cr)",
    "EBITDA or PAT (Rs. Cr)", "Revenue Growth YoY / 3-yr CAGR",
    "Capex Announced (Last 24 months)", "Current Steel Suppliers",
    "Approximate Annual Steel Tonnage (TPA)", "Buying Pattern",
    "Customer OEMs Served",
]

# Never treat these values as "present" — they mean the cell is effectively empty.
_EMPTY_TOKENS = {"", "nan", "none", "nat", "<na>", "n/a", "na", "-", "—", "tbd", "unknown"}


def is_empty(v):
    """True if a cell value should be considered blank/fillable."""
    if v is None:
        return True
    return str(v).strip().lower() in _EMPTY_TOKENS


def is_verified(flag_value):
    """A cell/row is protected from overwrite when its flag says 'Verified'."""
    return "verified" in str(flag_value or "").strip().lower()


# Note/section rows masquerading as company names. Enrichment skips these so we
# never spend a web-search call on a heading like "Some More Identified Vendors".
_JUNK_NAME = re.compile(
    r"\b(identified|vendors|sourcing|combination|various|miscellaneous|others?|"
    r"segment|category|section|see above|as above|to be|tbd|n/?a|list|"
    r"done globally|locked contract|best trip)\b", re.I)


def looks_like_company(name):
    """Conservative filter: True if `name` plausibly names a real company/event
    rather than a heading or free-text note row."""
    if is_empty(name):
        return False
    s = str(name).strip()
    if s.lower() in ("company name", "company", "name", "tbd", "n/a", "na"):
        return False
    if s.startswith(("----", "▼")):
        return False
    if _JUNK_NAME.search(s):
        return False
    # Real company/event names are short; sentence-like rows are not.
    if len(s.split()) > 7:
        return False
    return True


def today_iso():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d")


# ─── NEWS DE-DUPLICATION CACHE ────────────────────────────────────────────────
def _story_key(headline, link):
    raw = (str(headline).strip().lower() + "|" + str(link).strip().lower())
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]


def load_seen():
    """Return the set of story keys already published, plus the raw dict."""
    try:
        with open(NEWS_SEEN_FILE, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return set(data.get("keys", [])), data
    except (FileNotFoundError, json.JSONDecodeError):
        return set(), {"keys": [], "updated": ""}


def save_seen(keys):
    os.makedirs(DATA_DIR, exist_ok=True)
    payload = {"keys": sorted(keys), "updated": today_iso()}
    with open(NEWS_SEEN_FILE, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2)


def mark_seen(seen_set, headline, link):
    """Add a story to the seen set; return True if it was new."""
    k = _story_key(headline, link)
    if k in seen_set:
        return False
    seen_set.add(k)
    return True


# ─── KEYWORD / QUERY CONFIG (free Google News RSS + industry feeds) ───────────
# Broad market/commodity themes always fetched.
NEWS_TOPIC_QUERIES = [
    "HR coil steel price India",
    "steel price India HRC",
    "India automotive OEM capex expansion",
    "roll forming India manufacturing",
    "India steel tube pipe manufacturer expansion",
    "India construction infrastructure steel demand",
    "elevator manufacturer India",
    "India solar module mounting structure steel",
]

# Static industry RSS feeds (free, public). Google News queries are built at
# run time in fetch_news.py from these plus the company list.
INDUSTRY_RSS = [
    "https://economictimes.indiatimes.com/industry/indl-goods/svs/steel/rssfeeds/13358380.cms",
    "https://www.moneycontrol.com/rss/business.xml",
]

# Free, no-key FX endpoint for the USD/INR indicator & trend point.
FX_URL = "https://api.frankfurter.app/latest?from=USD&to=INR"


# ─── LLM PROVIDER LAYER (free-only, pluggable) ────────────────────────────────
# Which engine summarises the news / normalises enrichment fields. Resolved from
# env: "groq" if GROQ_API_KEY set, "gemini" if GEMINI_API_KEY set, else "none"
# (rule-based fallback, no network LLM call). Override with MIF_LLM_PROVIDER.
# All calls use plain `requests` — no paid SDKs, no web-search tool (free tiers
# lack it), so the model only ever sees text WE supply, keeping output grounded.
import requests as _requests

GROQ_URL   = "https://api.groq.com/openai/v1/chat/completions"
GEMINI_URL = ("https://generativelanguage.googleapis.com/v1beta/models/"
              "{model}:generateContent")


class LLMUnavailable(RuntimeError):
    """Raised when no free LLM key is configured; callers fall back to rules."""


def resolve_provider():
    p = os.environ.get("MIF_LLM_PROVIDER", "auto").strip().lower()
    if p in ("groq", "gemini", "none"):
        return p
    # auto
    if os.environ.get("GROQ_API_KEY"):
        return "groq"
    if os.environ.get("GEMINI_API_KEY"):
        return "gemini"
    return "none"


def llm_available():
    return resolve_provider() != "none"


def llm_json(prompt, system=None, max_tokens=4000):
    """Send `prompt` to the configured FREE provider and parse JSON from the
    reply. Raises LLMUnavailable when provider == 'none' (or the key is missing)
    so the caller can drop to its rule-based path."""
    provider = resolve_provider()
    if provider == "groq":
        text = _call_groq(prompt, system, max_tokens)
    elif provider == "gemini":
        text = _call_gemini(prompt, system, max_tokens)
    else:
        raise LLMUnavailable("no free LLM configured (MIF_LLM_PROVIDER=none)")
    return _parse_json(text)


def _call_groq(prompt, system, max_tokens):
    key = os.environ.get("GROQ_API_KEY")
    if not key:
        raise LLMUnavailable("GROQ_API_KEY is not set")
    model = os.environ.get("MIF_LLM_MODEL", "llama-3.1-8b-instant")
    messages = []
    if system:
        messages.append({"role": "system", "content": system})
    messages.append({"role": "user", "content": prompt})
    body = {"model": model, "messages": messages, "max_tokens": max_tokens,
            "temperature": 0.2, "response_format": {"type": "json_object"}}
    r = _requests.post(GROQ_URL, timeout=90,
                       headers={"Authorization": f"Bearer {key}",
                                "Content-Type": "application/json"},
                       json=body)
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"]


def _call_gemini(prompt, system, max_tokens):
    key = os.environ.get("GEMINI_API_KEY")
    if not key:
        raise LLMUnavailable("GEMINI_API_KEY is not set")
    model = os.environ.get("MIF_LLM_MODEL", "gemini-2.5-flash")
    url = GEMINI_URL.format(model=model)
    body = {"contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"maxOutputTokens": max_tokens, "temperature": 0.2,
                                 "responseMimeType": "application/json"}}
    if system:
        body["systemInstruction"] = {"parts": [{"text": system}]}
    r = _requests.post(url, timeout=90, params={"key": key},
                       headers={"Content-Type": "application/json"}, json=body)
    r.raise_for_status()
    return r.json()["candidates"][0]["content"]["parts"][0]["text"]


_JSON_BLOCK = re.compile(r"```(?:json)?\s*(.*?)```", re.S)


def _parse_json(text):
    m = _JSON_BLOCK.search(text)
    candidate = m.group(1).strip() if m else text.strip()
    # Fall back to the first {...} or [...] span if the model wrapped it in prose.
    if not m:
        span = re.search(r"(\{.*\}|\[.*\])", candidate, re.S)
        if span:
            candidate = span.group(1)
    return json.loads(candidate)


# ─── LOGGING ──────────────────────────────────────────────────────────────────
def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)
