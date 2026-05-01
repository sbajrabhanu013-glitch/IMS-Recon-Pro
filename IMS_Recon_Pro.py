import base64
import hashlib
import json
import os
import pickle
import re
import sqlite3
from dataclasses import dataclass
from datetime import date, datetime
from html import escape
from io import BytesIO
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ============================================================
# IMS RECON PRO | @BAJRABHANU
# Premium India-themed GST IMS Reconciliation Website/App
# ============================================================

APP_TITLE = "IMS Recon Pro"
APP_TAGLINE = "Intelligent GST IMS Reconciliation & Action Management Platform"
COPYRIGHT_OWNER = "@BAJRABHANU"
COPYRIGHT_NOTICE = "Copyright © @BAJRABHANU. All rights reserved."
APP_VERSION = "1.0.0"
APP_DB = "ims_recon_pro.db"
BASE_DIR = Path(__file__).resolve().parent

IMS_SHEETS = ["B2B", "B2BA", "B2B-DN", "B2B-DNA", "B2B-CN", "B2B-CNA"]
AMENDMENT_SHEETS = {"B2BA", "B2B-DNA", "B2B-CNA"}
TAX_COLS = ["taxable_value", "igst", "cgst", "sgst", "cess"]
MONEY_COLS = ["invoice_value", *TAX_COLS]
ACTIONS = ["No Action", "Accepted", "Rejected", "Pending", "Review"]
RISK_ORDER = ["Low", "Medium", "High", "Critical"]

USERS = {
    "MainAdmin": {"password": "Adminpwd", "role": "Main Admin"},
    "User1": {"password": "Userpwd1", "role": "Sub User - Upload & Reco"},
    "User2": {"password": "Userpwd2", "role": "Sub User - Review"},
}

COLUMN_ALIASES = {
    "supplier_gstin": [
        "supplier gstin", "gstin of supplier", "gstin", "ctin", "counterparty gstin", "party gstin",
        "vendor gstin", "gstin/uın", "gstin/uin", "recipient gstin",
    ],
    "supplier_name": [
        "supplier name", "trade/legal name", "trade legal name", "legal name", "party name", "vendor name",
        "name", "taxpayer name",
    ],
    "document_type": [
        "document type", "doc type", "invoice type", "type", "nature of document", "section",
    ],
    "document_no": [
        "document number", "document no", "doc no", "invoice number", "invoice no", "invoice", "bill no",
        "voucher number", "note number", "cdnr no", "credit note number", "debit note number",
    ],
    "document_date": [
        "document date", "doc date", "invoice date", "date", "bill date", "note date", "cdnr date",
    ],
    "invoice_value": [
        "invoice value", "invoice value(inr)", "invoice value(rs)", "total invoice value", "document value",
        "gross value", "total value", "note value", "value",
    ],
    "taxable_value": [
        "taxable value", "taxable value(inr)", "taxable value(rs)", "taxable amount", "assessable value", "net value",
    ],
    "igst": ["integrated tax", "integrated tax(₹)", "integrated tax(rs)", "igst", "igst amount"],
    "cgst": ["central tax", "central tax(₹)", "central tax(rs)", "cgst", "cgst amount"],
    "sgst": ["state/ut tax", "state tax", "state/ut tax(₹)", "sgst", "utgst", "sgst amount"],
    "cess": ["cess", "cess(₹)", "cess amount"],
    "itc_available": ["itc available", "itc availability", "eligible itc", "itc eligibility", "eligibility"],
    "ims_status": ["status", "ims status", "recipient action", "recipient status", "action"],
    "remarks": ["remarks", "remark", "reason", "comments", "comment"],
    "pos": ["place of supply", "pos", "state", "supply state"],
    "period": ["period", "return period", "tax period", "month"],
}

GSTIN_STATE_CODES = {
    "01": "Jammu & Kashmir", "02": "Himachal Pradesh", "03": "Punjab", "04": "Chandigarh", "05": "Uttarakhand",
    "06": "Haryana", "07": "Delhi", "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim",
    "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur", "15": "Mizoram", "16": "Tripura", "17": "Meghalaya",
    "18": "Assam", "19": "West Bengal", "20": "Jharkhand", "21": "Odisha", "22": "Chhattisgarh", "23": "Madhya Pradesh",
    "24": "Gujarat", "25": "Daman & Diu", "26": "Dadra & Nagar Haveli", "27": "Maharashtra", "28": "Andhra Pradesh",
    "29": "Karnataka", "30": "Goa", "31": "Lakshadweep", "32": "Kerala", "33": "Tamil Nadu", "34": "Puducherry",
    "35": "Andaman & Nicobar Islands", "36": "Telangana", "37": "Andhra Pradesh", "38": "Ladakh", "97": "Other Territory",
    "99": "Centre Jurisdiction",
}

st.set_page_config(
    page_title=f"{APP_TITLE} | {COPYRIGHT_OWNER}",
    page_icon="🇮🇳",
    layout="wide",
    initial_sidebar_state="expanded",
)

@dataclass
class UploadPack:
    raw: pd.DataFrame
    std: pd.DataFrame
    schema: Dict[str, Optional[str]]
    label: str

# ============================================================
# DATABASE AND SESSION
# ============================================================

def get_conn() -> sqlite3.Connection:
    return sqlite3.connect(APP_DB, check_same_thread=False)


def init_db() -> None:
    with get_conn() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS user_state (
                username TEXT PRIMARY KEY,
                saved_at TEXT NOT NULL,
                payload BLOB NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                event_time TEXT NOT NULL,
                username TEXT,
                role TEXT,
                event_type TEXT NOT NULL,
                detail TEXT NOT NULL
            )
            """
        )


def init_state() -> None:
    defaults = {
        "logged_in": False,
        "username": "",
        "role": "",
        "client_name": "",
        "client_gstin": "",
        "return_period": datetime.today().strftime("%b-%Y"),
        "prepared_by": "",
        "reviewed_by": "",
        "period_status": "Open",
        "purchase_pack": None,
        "ims_pack": None,
        "ims_json_pack": None,
        "primary_ims_source": "IMS Utility",
        "recon_result": pd.DataFrame(),
        "action_table": pd.DataFrame(),
        "active_page": "Dashboard",
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


def log_event(event_type: str, detail: str) -> None:
    try:
        with get_conn() as conn:
            conn.execute(
                "INSERT INTO audit_log (event_time, username, role, event_type, detail) VALUES (?, ?, ?, ?, ?)",
                (
                    datetime.now().isoformat(timespec="seconds"),
                    st.session_state.get("username", ""),
                    st.session_state.get("role", ""),
                    event_type,
                    detail,
                ),
            )
    except Exception:
        pass


def serialize_state() -> bytes:
    keys = [
        "client_name", "client_gstin", "return_period", "prepared_by", "reviewed_by", "period_status",
        "purchase_pack", "ims_pack", "ims_json_pack", "primary_ims_source", "recon_result", "action_table",
    ]
    data = {k: st.session_state.get(k) for k in keys}
    return pickle.dumps(data)


def save_user_state() -> None:
    if not st.session_state.get("logged_in"):
        return
    try:
        payload = serialize_state()
        with get_conn() as conn:
            conn.execute(
                "REPLACE INTO user_state (username, saved_at, payload) VALUES (?, ?, ?)",
                (st.session_state["username"], datetime.now().isoformat(timespec="seconds"), payload),
            )
    except Exception as exc:
        st.warning(f"Auto-save skipped: {exc}")


def load_user_state(username: str) -> None:
    try:
        with get_conn() as conn:
            row = conn.execute("SELECT payload FROM user_state WHERE username=?", (username,)).fetchone()
        if row:
            data = pickle.loads(row[0])
            for k, v in data.items():
                st.session_state[k] = v
    except Exception:
        pass


def reset_current_user_data() -> None:
    username = st.session_state.get("username", "")
    with get_conn() as conn:
        conn.execute("DELETE FROM user_state WHERE username=?", (username,))
    for key in ["purchase_pack", "ims_pack", "ims_json_pack", "recon_result", "action_table"]:
        st.session_state[key] = pd.DataFrame() if key in ["recon_result", "action_table"] else None
    log_event("Data Reset", "Current user data deleted completely.")

# ============================================================
# PREMIUM UI
# ============================================================

def inject_css() -> None:
    st.markdown(
        """
        <style>
        :root {
            --navy:#0B1F3A; --navy2:#102A4C; --ashoka:#1A4E8A; --saffron:#FF9933; --green:#138808;
            --bg:#F4F7FB; --card:#FFFFFF; --line:#D7E1EC; --muted:#64748B; --red:#DC2626; --amber:#F59E0B;
        }
        html, body, [data-testid="stAppViewContainer"] {
            background: radial-gradient(circle at top left, rgba(255,153,51,.12), transparent 26%),
                        radial-gradient(circle at top right, rgba(19,136,8,.10), transparent 28%),
                        linear-gradient(180deg,#f7fafc 0%,#eef4f9 44%,#f8fafc 100%);
            color: #0f172a;
        }
        .main .block-container { max-width: 1580px; padding-top: .7rem; padding-bottom: 2rem; }
        [data-testid="stSidebar"] { background: linear-gradient(180deg,#071a31 0%,#0B1F3A 48%,#102A4C 100%); }
        [data-testid="stSidebar"] * { color: #E5EEF8 !important; }
        [data-testid="stSidebar"] div[data-testid="stRadio"] label { border-radius: 14px; }
        .tricolor-strip { height: 7px; border-radius: 999px; background: linear-gradient(90deg,var(--saffron) 0 33.3%,#fff 33.3% 66.6%,var(--green) 66.6% 100%); box-shadow:0 8px 20px rgba(15,23,42,.08); margin-bottom:.75rem; border:1px solid rgba(11,31,58,.08); }
        .topbar {
            display:flex;align-items:center;justify-content:space-between;gap:1rem;
            background: linear-gradient(135deg, rgba(11,31,58,.98), rgba(26,78,138,.92));
            border:1px solid rgba(255,255,255,.20); border-radius: 22px; padding: .9rem 1rem;
            box-shadow: 0 24px 55px rgba(11,31,58,.24); color:#fff; margin-bottom:1rem; position:relative; overflow:hidden;
        }
        .topbar::after { content:""; position:absolute; right:220px; top:-70px; width:190px; height:190px; border:10px solid rgba(255,255,255,.08); border-radius:50%; box-shadow: inset 0 0 0 1px rgba(255,255,255,.08); }
        .brand { display:flex; align-items:center; gap:.85rem; position:relative; z-index:1; }
        .brand-logo { width:54px;height:54px;border-radius:18px;display:grid;place-items:center;background:#fff;color:var(--ashoka);font-size:1.55rem;font-weight:900; box-shadow: inset 0 -10px 20px rgba(26,78,138,.10); }
        .brand-title { font-size:1.25rem; font-weight:950; line-height:1.1; letter-spacing:.02em; }
        .brand-sub { font-size:.78rem; color:rgba(255,255,255,.78); margin-top:.15rem; }
        .topchips { display:flex; gap:.4rem; flex-wrap:wrap; justify-content:flex-end; position:relative; z-index:1; }
        .chip { border:1px solid rgba(255,255,255,.25); background:rgba(255,255,255,.10); border-radius:999px; padding:.33rem .62rem; font-size:.76rem; font-weight:800; white-space:nowrap; }
        .owner-chip { background:linear-gradient(90deg,rgba(255,153,51,.20),rgba(19,136,8,.20)); color:#fff; border-color:rgba(255,255,255,.35); }
        .hero-card { border:1px solid var(--line); background:rgba(255,255,255,.90); border-radius:24px; padding:1.1rem; box-shadow:0 18px 45px rgba(15,23,42,.08); margin-bottom:1rem; }
        .hero-title { font-size: clamp(1.35rem,2.5vw,2.25rem); font-weight:950; color:#0B1F3A; line-height:1.08; }
        .hero-note { color:var(--muted); margin-top:.3rem; font-size:.95rem; }
        .kpi-card { border:1px solid var(--line); background:linear-gradient(180deg,#fff,#f8fbff); border-radius:22px; padding:1rem; box-shadow:0 15px 34px rgba(15,23,42,.07); min-height:124px; position:relative; overflow:hidden; }
        .kpi-card:before { content:""; position:absolute; inset:0 0 auto 0; height:5px; background:linear-gradient(90deg,var(--saffron),#fff,var(--green)); }
        .kpi-icon { font-size:1.55rem; margin-bottom:.35rem; }
        .kpi-label { color:var(--muted); font-size:.76rem; font-weight:900; text-transform:uppercase; letter-spacing:.05em; }
        .kpi-value { color:#0B1F3A; font-size:1.65rem; font-weight:950; margin-top:.2rem; }
        .kpi-note { color:var(--muted); font-size:.78rem; margin-top:.15rem; }
        .section-title { display:flex;justify-content:space-between;align-items:center;gap:1rem; border:1px solid var(--line); border-radius:18px; padding:.75rem .95rem; background:#fff; box-shadow:0 10px 24px rgba(15,23,42,.05); margin:.9rem 0 .7rem 0; }
        .section-title strong { color:#0B1F3A; font-size:1rem; }
        .section-title span { color:var(--muted); font-size:.8rem; }
        .copyright-floating { position:fixed; right:18px; bottom:16px; z-index:999; color:rgba(11,31,58,.22); font-weight:950; letter-spacing:.08em; pointer-events:none; }
        .sidebar-brand { border:1px solid rgba(255,255,255,.14); border-radius:20px; padding:.9rem; background:rgba(255,255,255,.06); margin-bottom:1rem; }
        .sidebar-title { font-weight:950; font-size:1.16rem; color:#fff; }
        .sidebar-sub { color:#cbd5e1; font-size:.74rem; margin-top:.2rem; }
        .pill { display:inline-flex; align-items:center; gap:.3rem; padding:.22rem .55rem; border-radius:999px; border:1px solid var(--line); background:#f8fafc; font-size:.78rem; font-weight:800; color:#334155; }
        .pill-ok { background:#ECFDF5; border-color:#BBF7D0; color:#166534; }
        .pill-warn { background:#FFFBEB; border-color:#FDE68A; color:#92400E; }
        .pill-risk { background:#FEF2F2; border-color:#FECACA; color:#991B1B; }
        .premium-panel { border:1px solid var(--line); border-radius:24px; padding:1rem; background:rgba(255,255,255,.92); box-shadow:0 18px 45px rgba(15,23,42,.07); }
        .upload-card { border:1px dashed #B7C7D8; background:linear-gradient(180deg,#fff,#f8fbff); border-radius:24px; padding:1rem; min-height:220px; box-shadow:0 13px 30px rgba(15,23,42,.06); }
        .upload-card h3 { margin:0 0 .2rem 0; color:#0B1F3A; }
        .upload-card p { color:var(--muted); font-size:.82rem; }
        .stButton button, .stDownloadButton button { border-radius:14px !important; font-weight:850 !important; border:1px solid #B7C7D8 !important; }
        div[data-testid="stDataFrame"] { border:1px solid var(--line); border-radius:18px; overflow:hidden; }
        .login-bg { min-height:72vh; display:grid; place-items:center; }
        .login-card { width:min(480px,96%); border:1px solid rgba(255,255,255,.5); border-radius:28px; padding:1.25rem; background:rgba(255,255,255,.90); box-shadow:0 30px 80px rgba(11,31,58,.20); }
        .login-logo { width:72px;height:72px;border-radius:24px; display:grid;place-items:center; background:linear-gradient(135deg,var(--saffron),#fff,var(--green)); color:var(--ashoka); font-size:2rem; margin:auto; box-shadow:0 18px 40px rgba(15,23,42,.15); }
        .login-title { text-align:center; color:#0B1F3A; font-size:1.7rem; font-weight:950; margin-top:.75rem; }
        .login-sub { text-align:center; color:var(--muted); margin-bottom:.9rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def chakra_svg(size: int = 24) -> str:
    spokes = "".join([
        f'<line x1="50" y1="50" x2="50" y2="8" transform="rotate({i*15} 50 50)" stroke="#1A4E8A" stroke-width="2" opacity=".85" />'
        for i in range(24)
    ])
    return f'<svg width="{size}" height="{size}" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg"><circle cx="50" cy="50" r="43" fill="none" stroke="#1A4E8A" stroke-width="5"/><circle cx="50" cy="50" r="5" fill="#1A4E8A"/>{spokes}</svg>'


def icon(name: str) -> str:
    icons = {
        "dashboard": "◆", "setup": "▣", "upload": "⇧", "viewer": "☷", "recon": "⟳", "action": "✓",
        "risk": "⚠", "vendor": "✉", "report": "▤", "ai": "✦", "admin": "♛", "logout": "↪",
    }
    return icons.get(name, "•")


def render_topbar() -> None:
    user = st.session_state.get("username", "Guest")
    role = st.session_state.get("role", "")
    client = st.session_state.get("client_name") or "Client not set"
    gstin = st.session_state.get("client_gstin") or "GSTIN not set"
    period = st.session_state.get("return_period") or "Period not set"
    st.markdown('<div class="tricolor-strip"></div>', unsafe_allow_html=True)
    st.markdown(
        f"""
        <div class="topbar">
            <div class="brand">
                <div class="brand-logo">{chakra_svg(34)}</div>
                <div>
                    <div class="brand-title">{APP_TITLE}</div>
                    <div class="brand-sub">{APP_TAGLINE}</div>
                </div>
            </div>
            <div class="topchips">
                <span class="chip">{escape(client)}</span>
                <span class="chip">{escape(gstin)}</span>
                <span class="chip">{escape(period)}</span>
                <span class="chip">{escape(user)} · {escape(role)}</span>
                <span class="chip owner-chip">{COPYRIGHT_OWNER}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(f'<div class="copyright-floating">{COPYRIGHT_OWNER}</div>', unsafe_allow_html=True)


def hero(title: str, note: str) -> None:
    st.markdown(
        f"""
        <div class="hero-card">
            <div class="hero-title">{escape(title)}</div>
            <div class="hero-note">{escape(note)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section(title: str, note: str = "") -> None:
    st.markdown(
        f"""
        <div class="section-title">
            <div><strong>{escape(title)}</strong><br><span>{escape(note)}</span></div>
            <div class="pill">{COPYRIGHT_OWNER}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def kpi(label: str, value, note: str = "", symbol: str = "◆") -> None:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-icon">{symbol}</div>
            <div class="kpi-label">{escape(label)}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{escape(note)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ============================================================
# DATA HELPERS
# ============================================================

def clean_header(value: object) -> str:
    text = str(value or "").strip().lower()
    text = text.replace("\n", " ").replace("₹", "rs")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^a-z0-9%/() -]", "", text)
    return text.strip()


def find_col(df: pd.DataFrame, logical: str) -> Optional[str]:
    aliases = COLUMN_ALIASES.get(logical, [logical])
    normalized = {clean_header(c): c for c in df.columns}
    for alias in aliases:
        c = clean_header(alias)
        if c in normalized:
            return normalized[c]
    for alias in aliases:
        c = clean_header(alias)
        for norm, original in normalized.items():
            if c and (c in norm or norm in c):
                return original
    return None


def detect_schema(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    return {k: find_col(df, k) for k in COLUMN_ALIASES}


def make_unique_columns(headers: Iterable[object]) -> List[str]:
    seen: Dict[str, int] = {}
    out = []
    for i, h in enumerate(headers, start=1):
        name = str(h or "").strip()
        if not name or name.lower().startswith("unnamed"):
            name = f"Column {i}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        out.append(name)
    return out


def detect_header_row(raw: pd.DataFrame) -> int:
    alias_words = [clean_header(x) for vals in COLUMN_ALIASES.values() for x in vals]
    best_idx = raw.index[0]
    best_score = -1
    for idx, row in raw.head(35).iterrows():
        text = " ".join(clean_header(x) for x in row.tolist())
        score = sum(1 for a in alias_words if len(a) > 2 and a in text)
        if score > best_score:
            best_idx, best_score = idx, score
    return int(best_idx)


def normalize_sheet_frame(sheet_df: pd.DataFrame) -> pd.DataFrame:
    raw = sheet_df.dropna(how="all").dropna(how="all", axis=1)
    if raw.empty:
        return pd.DataFrame()
    header_idx = detect_header_row(raw)
    header_pos = list(raw.index).index(header_idx)
    headers = make_unique_columns(raw.loc[header_idx].fillna("").astype(str).tolist())
    body = raw.iloc[header_pos + 1:].copy()
    body.columns = headers
    body = body.dropna(how="all")
    body = body.loc[:, [c for c in body.columns if str(c).strip()]]
    return body


def read_any_table(file, label: str, target_sheets: Optional[List[str]] = None) -> pd.DataFrame:
    if file is None:
        return pd.DataFrame()
    name = getattr(file, "name", str(file))
    try:
        if str(name).lower().endswith(".csv"):
            df = pd.read_csv(file, dtype=object)
            df["_source_sheet"] = "CSV"
            return df
        sheets = pd.read_excel(file, sheet_name=None, dtype=object, header=None, engine="openpyxl")
        frames = []
        for sheet_name, sheet_df in sheets.items():
            if target_sheets and sheet_name not in target_sheets:
                continue
            norm = normalize_sheet_frame(sheet_df)
            if not norm.empty:
                norm["_source_sheet"] = sheet_name
                frames.append(norm)
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    except Exception as exc:
        st.error(f"{label}: unable to read file. {exc}")
        return pd.DataFrame()


def normalize_doc_no(value: object) -> str:
    text = str(value or "").strip().upper()
    text = re.sub(r"\.0$", "", text)
    return re.sub(r"[^A-Z0-9]", "", text)


def normalize_gstin(value: object) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper().strip())


def validate_gstin(value: object) -> bool:
    text = normalize_gstin(value)
    return bool(re.match(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$", text)) and text[:2] in GSTIN_STATE_CODES


def to_num(s: pd.Series) -> pd.Series:
    cleaned = (
        s.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False)
        .str.replace("INR", "", case=False, regex=False)
        .str.replace("RS.", "", case=False, regex=False)
        .str.replace("RS", "", case=False, regex=False)
        .str.strip()
    )
    cleaned = cleaned.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    cleaned = cleaned.replace({"": "0", "nan": "0", "None": "0", "NaT": "0"})
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def to_date(s: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(s, errors="coerce")
    excel_dates = pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
    parsed = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return parsed.fillna(excel_dates)


def infer_document_type(row: pd.Series, source_sheet: str = "") -> str:
    text = " ".join(str(row.get(k, "")) for k in ["document_type", "remarks", "ims_status"]).lower() + " " + source_sheet.lower()
    if "cn" in source_sheet.lower() or "credit" in text or "cr note" in text:
        return "Credit Note"
    if "dn" in source_sheet.lower() or "debit" in text or "dr note" in text:
        return "Debit Note"
    if "amend" in text or source_sheet.upper() in AMENDMENT_SHEETS:
        return "Amendment"
    return "Invoice"


def standardize(df: pd.DataFrame, label: str, source_type: str) -> Tuple[pd.DataFrame, Dict[str, Optional[str]]]:
    if df.empty:
        return pd.DataFrame(), {}
    schema = detect_schema(df)
    out = pd.DataFrame(index=df.index)
    def col(logical, default=""):
        c = schema.get(logical)
        return df[c] if c and c in df.columns else pd.Series([default] * len(df), index=df.index)
    out["supplier_gstin"] = col("supplier_gstin").map(normalize_gstin)
    out["supplier_name"] = col("supplier_name").astype(str).str.strip()
    out["document_type_raw"] = col("document_type").astype(str).str.strip()
    out["document_no"] = col("document_no").astype(str).str.strip()
    out["document_norm"] = out["document_no"].map(normalize_doc_no)
    out["document_date"] = to_date(col("document_date", pd.NaT))
    for c in MONEY_COLS:
        out[c] = to_num(col(c, 0))
    out["total_tax"] = out[["igst", "cgst", "sgst", "cess"]].sum(axis=1)
    out["itc_available"] = col("itc_available", "").astype(str).str.strip()
    out["ims_status"] = col("ims_status", "").astype(str).str.strip()
    out["remarks"] = col("remarks", "").astype(str).str.strip()
    out["pos"] = col("pos", "").astype(str).str.strip()
    out["period"] = col("period", "").astype(str).str.strip()
    out["source_sheet"] = df["_source_sheet"] if "_source_sheet" in df.columns else ""
    out["document_type"] = out.apply(lambda r: infer_document_type(r, str(r.get("source_sheet", ""))), axis=1)
    out["source_type"] = source_type
    out["source_label"] = label
    out["gstin_valid"] = out["supplier_gstin"].map(validate_gstin)
    out["state_name"] = out["supplier_gstin"].str[:2].map(GSTIN_STATE_CODES).fillna("")
    out["match_key"] = out["supplier_gstin"].astype(str) + "|" + out["document_norm"].astype(str)
    out["data_quality_score"] = out.apply(data_quality_score, axis=1)
    out["row_id"] = range(1, len(out) + 1)
    out = out[(out["supplier_gstin"].ne("") | out["document_norm"].ne(""))].copy()
    return out, schema


def data_quality_score(row: pd.Series) -> int:
    score = 100
    if not row.get("gstin_valid", False) and row.get("supplier_gstin"):
        score -= 25
    if not str(row.get("document_norm", "")).strip():
        score -= 30
    if pd.isna(row.get("document_date")):
        score -= 15
    if abs(float(row.get("taxable_value", 0))) < .01 and abs(float(row.get("total_tax", 0))) > .01:
        score -= 20
    if float(row.get("igst", 0) or 0) and (float(row.get("cgst", 0) or 0) or float(row.get("sgst", 0) or 0)):
        score -= 10
    return max(0, min(100, score))


def make_pack(file, label: str, source_type: str, target_sheets: Optional[List[str]] = None) -> Optional[UploadPack]:
    raw = read_any_table(file, label, target_sheets)
    if raw.empty:
        return None
    std, schema = standardize(raw, label, source_type)
    if std.empty:
        st.warning(f"{label}: no usable rows detected.")
        return None
    pack = UploadPack(raw=raw, std=std, schema=schema, label=label)
    log_event("Upload", f"{label}: {len(std):,} standardized rows loaded.")
    save_user_state()
    return pack


def flatten_json_records(obj, records: Optional[List[dict]] = None) -> List[dict]:
    if records is None:
        records = []
    if isinstance(obj, dict):
        keys = {str(k).lower() for k in obj.keys()}
        likely_doc = any(k in keys for k in ["inum", "invoice_number", "invoice number", "doc_num", "nt_num", "ctin", "gstin"]) and any(
            k in keys for k in ["txval", "taxable_value", "iamt", "camt", "samt", "val", "invoice_value"]
        )
        if likely_doc:
            records.append(obj.copy())
        for v in obj.values():
            flatten_json_records(v, records)
    elif isinstance(obj, list):
        for item in obj:
            flatten_json_records(item, records)
    return records


def parse_ims_json(file) -> Optional[UploadPack]:
    if file is None:
        return None
    try:
        content = file.read()
        if isinstance(content, bytes):
            content = content.decode("utf-8", errors="ignore")
        obj = json.loads(content)
        rows = flatten_json_records(obj)
        if not rows:
            st.error("No IMS-like document rows detected in JSON.")
            return None
        df = pd.json_normalize(rows)
        df["_source_sheet"] = "IMS_JSON"
        # Add extra alias columns from GST JSON short names when available.
        rename_map = {}
        for c in df.columns:
            lc = c.lower().split(".")[-1]
            if lc in {"ctin"}: rename_map[c] = "Supplier GSTIN"
            elif lc in {"inum", "nt_num", "doc_num"}: rename_map[c] = "Document Number"
            elif lc in {"idt", "nt_dt", "doc_dt"}: rename_map[c] = "Document Date"
            elif lc in {"val"}: rename_map[c] = "Invoice Value"
            elif lc in {"txval"}: rename_map[c] = "Taxable Value"
            elif lc in {"iamt"}: rename_map[c] = "IGST"
            elif lc in {"camt"}: rename_map[c] = "CGST"
            elif lc in {"samt"}: rename_map[c] = "SGST"
            elif lc in {"csamt"}: rename_map[c] = "Cess"
        df = df.rename(columns=rename_map)
        std, schema = standardize(df, "IMS JSON", "IMS_JSON")
        pack = UploadPack(raw=df, std=std, schema=schema, label="IMS JSON")
        log_event("Upload", f"IMS JSON: {len(std):,} standardized rows loaded.")
        save_user_state()
        return pack
    except Exception as exc:
        st.error(f"Unable to parse IMS JSON: {exc}")
        return None

# ============================================================
# RECONCILIATION ENGINE
# ============================================================

def aggregate(df: pd.DataFrame, suffix: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    w = df.copy()
    w = w[w["match_key"].astype(str).ne("|")]
    agg = {
        "supplier_name": "first", "document_type": "first", "document_no": "first", "document_date": "min",
        "invoice_value": "sum", "taxable_value": "sum", "igst": "sum", "cgst": "sum", "sgst": "sum", "cess": "sum",
        "total_tax": "sum", "itc_available": "first", "ims_status": "first", "remarks": "first", "source_sheet": "first",
        "data_quality_score": "mean", "row_id": "count",
    }
    out = w.groupby(["supplier_gstin", "document_norm"], dropna=False).agg(agg).reset_index()
    return out.rename(columns={c: f"{c}_{suffix}" for c in agg.keys()})


def start_reconciliation(purchase: pd.DataFrame, ims: pd.DataFrame, amount_tol: float, date_tol: int, include_amendments: bool) -> pd.DataFrame:
    if purchase is None or purchase.empty or ims is None or ims.empty:
        return pd.DataFrame()
    ims_work = ims.copy()
    if not include_amendments:
        ims_work = ims_work[~ims_work["source_sheet"].isin(AMENDMENT_SHEETS)].copy()
    p = aggregate(purchase, "purchase")
    i = aggregate(ims_work, "ims")
    if p.empty or i.empty:
        return pd.DataFrame()
    m = p.merge(i, on=["supplier_gstin", "document_norm"], how="outer", indicator=True)
    for c in MONEY_COLS + ["total_tax"]:
        m[f"{c}_diff"] = m.get(f"{c}_purchase", 0).fillna(0) - m.get(f"{c}_ims", 0).fillna(0)
    pdate = pd.to_datetime(m.get("document_date_purchase"), errors="coerce")
    idate = pd.to_datetime(m.get("document_date_ims"), errors="coerce")
    m["date_diff_days"] = (pdate - idate).dt.days.abs().fillna(0).astype("Int64")
    both = m["_merge"].eq("both")
    amount_ok = m["taxable_value_diff"].abs().le(amount_tol) & m["total_tax_diff"].abs().le(amount_tol)
    tax_head_ok = m[["igst_diff", "cgst_diff", "sgst_diff", "cess_diff"]].abs().le(amount_tol).all(axis=1)
    date_ok = m["date_diff_days"].fillna(0).le(date_tol)
    m["mismatch_type"] = "Matched"
    m.loc[m["_merge"].eq("left_only"), "mismatch_type"] = "Only in Purchase"
    m.loc[m["_merge"].eq("right_only"), "mismatch_type"] = "Only in IMS"
    m.loc[both & amount_ok & ~tax_head_ok, "mismatch_type"] = "Tax Head Mismatch"
    m.loc[both & ~amount_ok & date_ok, "mismatch_type"] = "Value / Tax Mismatch"
    m.loc[both & amount_ok & tax_head_ok & ~date_ok, "mismatch_type"] = "Date Mismatch"
    m.loc[both & ~amount_ok & ~date_ok, "mismatch_type"] = "Value and Date Mismatch"
    m["risk_score"] = m.apply(calculate_risk_score, axis=1)
    m["risk_level"] = m["risk_score"].map(risk_level)
    m["recommended_ims_action"] = m.apply(recommend_action, axis=1)
    m["reason"] = m.apply(action_reason, axis=1)
    m["vendor_followup_required"] = m["recommended_ims_action"].isin(["Pending", "Review"]) | m["mismatch_type"].isin(["Only in Purchase", "Value / Tax Mismatch", "Tax Head Mismatch"])
    m["confidence_score"] = m.apply(confidence_score, axis=1)
    m["final_user_action"] = m["recommended_ims_action"]
    m["user_remarks"] = ""
    m = m.sort_values(["risk_score", "supplier_gstin", "document_norm"], ascending=[False, True, True])
    log_event("Reconciliation", f"IMS reconciliation completed with {len(m):,} result rows.")
    return m


def calculate_risk_score(row: pd.Series) -> int:
    score = 0
    mt = str(row.get("mismatch_type", ""))
    tax_impact = abs(float(row.get("total_tax_diff", 0) or 0))
    taxable_impact = abs(float(row.get("taxable_value_diff", 0) or 0))
    if mt == "Matched": score += 5
    if mt == "Only in IMS": score += 35
    if mt == "Only in Purchase": score += 25
    if "Value" in mt: score += 35
    if "Tax Head" in mt: score += 30
    if "Date" in mt: score += 12
    if tax_impact >= 100000 or taxable_impact >= 500000: score += 25
    elif tax_impact >= 10000 or taxable_impact >= 100000: score += 15
    doc_type = f"{row.get('document_type_purchase','')} {row.get('document_type_ims','')}`".lower()
    if "credit" in doc_type: score += 15
    if str(row.get("source_sheet_ims", "")).upper() in AMENDMENT_SHEETS: score += 15
    return min(100, score)


def risk_level(score: int) -> str:
    if score >= 76: return "Critical"
    if score >= 51: return "High"
    if score >= 21: return "Medium"
    return "Low"


def recommend_action(row: pd.Series) -> str:
    mt = str(row.get("mismatch_type", ""))
    itc = str(row.get("itc_available_purchase", "")).lower()
    source_sheet = str(row.get("source_sheet_ims", "")).upper()
    if "not" in itc or "no" == itc.strip() or "ineligible" in itc:
        return "Rejected"
    if mt == "Matched" and source_sheet not in AMENDMENT_SHEETS:
        return "Accepted"
    if mt == "Matched" and source_sheet in AMENDMENT_SHEETS:
        return "Pending"
    if mt == "Only in IMS":
        return "Pending"
    if mt == "Only in Purchase":
        return "No Action"
    if mt in {"Value / Tax Mismatch", "Tax Head Mismatch", "Date Mismatch", "Value and Date Mismatch"}:
        return "Pending"
    return "Review"


def action_reason(row: pd.Series) -> str:
    mt = str(row.get("mismatch_type", ""))
    if mt == "Matched":
        return "Document found in Purchase Register and IMS with values within tolerance."
    if mt == "Only in IMS":
        return "Document is available in IMS but not booked in Purchase Register; keep pending until books/vendor confirmation."
    if mt == "Only in Purchase":
        return "Document is booked in Purchase Register but not available in IMS; vendor follow-up required, no IMS action available."
    if mt == "Tax Head Mismatch":
        return "Total may be close but IGST/CGST/SGST/Cess split differs; verify POS and tax classification."
    if "Value" in mt:
        return "Taxable value or tax amount differs; verify invoice/CN/DN/amendment before accepting."
    if mt == "Date Mismatch":
        return "Document number matched but date differs beyond tolerance; verify document date."
    return "Manual review required."


def confidence_score(row: pd.Series) -> int:
    if row.get("_merge") != "both": return 0
    score = 100
    score -= min(40, int(abs(float(row.get("total_tax_diff", 0) or 0)) // 1000))
    score -= min(30, int(abs(float(row.get("taxable_value_diff", 0) or 0)) // 10000))
    score -= min(15, int(row.get("date_diff_days", 0) or 0))
    if row.get("mismatch_type") == "Tax Head Mismatch": score -= 12
    return max(0, min(100, score))

# ============================================================
# EXPORTS
# ============================================================

def excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        meta = pd.DataFrame([
            {"Field": "Application", "Value": APP_TITLE},
            {"Field": "Version", "Value": APP_VERSION},
            {"Field": "Owner", "Value": COPYRIGHT_OWNER},
            {"Field": "Generated On", "Value": datetime.now().strftime("%d-%b-%Y %H:%M:%S")},
            {"Field": "Client", "Value": st.session_state.get("client_name", "")},
            {"Field": "GSTIN", "Value": st.session_state.get("client_gstin", "")},
            {"Field": "Return Period", "Value": st.session_state.get("return_period", "")},
        ])
        meta.to_excel(writer, index=False, sheet_name="BAJRABHANU")
        used = {"BAJRABHANU"}
        for sheet_name, df in sheets.items():
            safe = re.sub(r"[\[\]:*?/\\]", "_", str(sheet_name))[:31] or "Sheet"
            if safe in used:
                safe = f"{safe[:27]}_{len(used)}"
            used.add(safe)
            data = df.copy() if isinstance(df, pd.DataFrame) and not df.empty else pd.DataFrame({"Message": ["No data available"]})
            # Avoid PyArrow/Excel mixed object issues.
            for col in data.columns:
                if data[col].dtype == "object":
                    data[col] = data[col].astype(str)
            data.to_excel(writer, index=False, sheet_name=safe)
            ws = writer.sheets[safe]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            header_fill = PatternFill("solid", fgColor="0B1F3A")
            header_font = Font(color="FFFFFF", bold=True)
            thin = Side(style="thin", color="D9E2EC")
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            for row in ws.iter_rows():
                for cell in row:
                    cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
            for col_cells in ws.columns:
                letter = get_column_letter(col_cells[0].column)
                max_len = max(len(str(c.value or "")) for c in col_cells[:200])
                ws.column_dimensions[letter].width = min(max_len + 2, 38)
        wb = writer.book
        wb.properties.creator = COPYRIGHT_OWNER
        wb.properties.title = APP_TITLE
        wb.properties.subject = COPYRIGHT_NOTICE
    buffer.seek(0)
    return buffer.getvalue()

# ============================================================
# PAGES
# ============================================================

def login_page() -> None:
    st.markdown('<div class="tricolor-strip"></div>', unsafe_allow_html=True)
    st.markdown('<div class="login-bg"><div class="login-card">', unsafe_allow_html=True)
    st.markdown(f'<div class="login-logo">{chakra_svg(44)}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="login-title">{APP_TITLE}</div><div class="login-sub">{APP_TAGLINE}<br>{COPYRIGHT_OWNER}</div>', unsafe_allow_html=True)
    user = st.text_input("User ID")
    pwd = st.text_input("Password", type="password")
    if st.button("Secure Login", width="stretch", type="primary"):
        if user in USERS and USERS[user]["password"] == pwd:
            st.session_state.logged_in = True
            st.session_state.username = user
            st.session_state.role = USERS[user]["role"]
            load_user_state(user)
            log_event("Login", "User logged in successfully.")
            st.rerun()
        else:
            st.error("Invalid User ID or Password")
    st.caption("Demo users: MainAdmin / User1 / User2")
    st.markdown('</div></div>', unsafe_allow_html=True)


def sidebar() -> str:
    with st.sidebar:
        st.markdown(
            f"""
            <div class="sidebar-brand">
                <div style="font-size:2rem;">{chakra_svg(36)}</div>
                <div class="sidebar-title">{APP_TITLE}</div>
                <div class="sidebar-sub">{APP_TAGLINE}</div>
                <div style="margin-top:.6rem;font-weight:950;color:#FF9933;">{COPYRIGHT_OWNER}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        pages = [
            f"{icon('dashboard')} Dashboard",
            f"{icon('setup')} Client Setup",
            f"{icon('upload')} Upload Center",
            f"{icon('viewer')} IMS Data Viewer",
            f"{icon('recon')} Reconciliation Workspace",
            f"{icon('action')} IMS Action Center",
            f"{icon('risk')} Risk & Exception Center",
            f"{icon('vendor')} Vendor Follow-up",
            f"{icon('report')} Reports & Export",
            f"{icon('ai')} AI Insight Desk",
            f"{icon('admin')} Admin Panel",
        ]
        selected = st.radio("Navigation", pages, label_visibility="collapsed")
        st.divider()
        st.caption(f"Logged in: {st.session_state.get('username')} · {st.session_state.get('role')}")
        if st.button(f"{icon('logout')} Logout", width="stretch"):
            save_user_state()
            log_event("Logout", "User logged out.")
            for key in ["logged_in", "username", "role"]:
                st.session_state[key] = False if key == "logged_in" else ""
            st.rerun()
    return selected.split(" ", 1)[1]


def dashboard_page() -> None:
    hero("IMS Command Dashboard", "Premium control room for IMS actions, ITC exposure, vendor follow-up and reconciliation status.")
    purchase = st.session_state.get("purchase_pack")
    ims = active_ims_pack()
    recon = st.session_state.get("recon_result", pd.DataFrame())
    action = st.session_state.get("action_table", pd.DataFrame())
    c = st.columns(5)
    with c[0]: kpi("Purchase Docs", f"{len(purchase.std):,}" if purchase else "0", "Books data", "▣")
    with c[1]: kpi("IMS Docs", f"{len(ims.std):,}" if ims else "0", "Utility / JSON", "⇧")
    with c[2]: kpi("Reco Rows", f"{len(recon):,}" if isinstance(recon, pd.DataFrame) else "0", "Latest run", "⟳")
    with c[3]:
        accepted = int((action.get("final_user_action", pd.Series(dtype=str)) == "Accepted").sum()) if isinstance(action, pd.DataFrame) and not action.empty else 0
        kpi("Accepted", f"{accepted:,}", "Final action", "✓")
    with c[4]:
        high = int(action.get("risk_level", pd.Series(dtype=str)).isin(["High", "Critical"]).sum()) if isinstance(action, pd.DataFrame) and not action.empty else 0
        kpi("High Risk", f"{high:,}", "Needs review", "⚠")

    section("Action Summary", "Status distribution and high-level risk view")
    if isinstance(action, pd.DataFrame) and not action.empty:
        a1, a2 = st.columns(2)
        with a1:
            fig = px.pie(action, names="final_user_action", title="IMS Final Action Mix", hole=.58)
            fig.update_layout(height=390, paper_bgcolor="rgba(0,0,0,0)", font={"color":"#0B1F3A"})
            st.plotly_chart(fig, width="stretch")
        with a2:
            fig = px.bar(action.groupby("risk_level", dropna=False).size().reset_index(name="Count"), x="risk_level", y="Count", title="Risk Summary")
            fig.update_layout(height=390, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, width="stretch")
    else:
        st.info("Upload data and run reconciliation to activate the dashboard.")

    section("Quick Actions", "Move directly into the next step")
    q1, q2, q3, q4 = st.columns(4)
    with q1:
        if st.button("⇧ Upload Data", width="stretch"):
            st.session_state.nav_override = "Upload Center"
            st.rerun()
    with q2:
        if st.button("⟳ Start Reconciliation", width="stretch"):
            st.session_state.nav_override = "Reconciliation Workspace"
            st.rerun()
    with q3:
        if st.button("✓ Action Center", width="stretch"):
            st.session_state.nav_override = "IMS Action Center"
            st.rerun()
    with q4:
        if st.button("▤ Reports", width="stretch"):
            st.session_state.nav_override = "Reports & Export"
            st.rerun()


def client_setup_page() -> None:
    hero("Client Setup", "Set client identity, GSTIN, return period and review controls before starting IMS reconciliation.")
    section("Client Identity", "These details will appear in reports and exports")
    c1, c2, c3 = st.columns(3)
    with c1: st.session_state.client_name = st.text_input("Client Name", st.session_state.client_name)
    with c2: st.session_state.client_gstin = st.text_input("Client GSTIN", st.session_state.client_gstin).upper()
    with c3: st.session_state.return_period = st.text_input("Return Period", st.session_state.return_period)
    c4, c5, c6 = st.columns(3)
    with c4: st.session_state.prepared_by = st.text_input("Prepared By", st.session_state.prepared_by)
    with c5: st.session_state.reviewed_by = st.text_input("Reviewed By", st.session_state.reviewed_by)
    with c6: st.session_state.period_status = st.selectbox("Period Status", ["Open", "Under Review", "Finalized", "Locked"], index=["Open", "Under Review", "Finalized", "Locked"].index(st.session_state.period_status) if st.session_state.period_status in ["Open", "Under Review", "Finalized", "Locked"] else 0)
    if st.session_state.client_gstin and not validate_gstin(st.session_state.client_gstin):
        st.warning("Client GSTIN format/state code looks invalid.")
    if st.button("Save Client Setup", width="stretch", type="primary"):
        save_user_state()
        log_event("Client Setup", "Client setup saved.")
        st.success("Client setup saved.")


def upload_center_page() -> None:
    hero("Upload Center", "Upload Purchase Register, GST IMS Utility .xlsm or direct IMS JSON in a guided premium workflow.")
    section("Upload Sources", "Recommended Phase 1 source: Purchase Register + populated GST IMS Utility")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown('<div class="upload-card"><h3>▣ Purchase Register</h3><p>Upload books data with invoices, credit notes and debit notes.</p>', unsafe_allow_html=True)
        f = st.file_uploader("Purchase Register", type=["xlsx", "xls", "csv"], key="purchase_upload")
        if f:
            pack = make_pack(f, "Purchase Register", "PURCHASE")
            if pack: st.session_state.purchase_pack = pack; st.success(f"Loaded {len(pack.std):,} rows")
        st.markdown('</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="upload-card"><h3>⇧ GST IMS Utility</h3><p>Upload populated GST offline utility .xlsm after importing IMS JSON in Excel.</p>', unsafe_allow_html=True)
        f = st.file_uploader("GST IMS Utility .xlsm", type=["xlsm", "xlsx"], key="ims_utility_upload")
        if f:
            pack = make_pack(f, "GST IMS Utility", "IMS_UTILITY", target_sheets=IMS_SHEETS)
            if pack: st.session_state.ims_pack = pack; st.success(f"Loaded {len(pack.std):,} IMS rows")
        st.markdown('</div>', unsafe_allow_html=True)
    with c3:
        st.markdown('<div class="upload-card"><h3>☷ IMS JSON</h3><p>Upload raw IMS JSON. Use this when utility file is not available.</p>', unsafe_allow_html=True)
        f = st.file_uploader("IMS JSON", type=["json"], key="ims_json_upload")
        if f:
            pack = parse_ims_json(f)
            if pack: st.session_state.ims_json_pack = pack; st.success(f"Loaded {len(pack.std):,} JSON rows")
        st.markdown('</div>', unsafe_allow_html=True)

    st.session_state.primary_ims_source = st.radio("Primary IMS Source", ["IMS Utility", "IMS JSON"], horizontal=True, index=0 if st.session_state.primary_ims_source == "IMS Utility" else 1)

    section("Data Health Check", "Review upload counts and validation position")
    rows = []
    for name, pack in [("Purchase Register", st.session_state.get("purchase_pack")), ("IMS Utility", st.session_state.get("ims_pack")), ("IMS JSON", st.session_state.get("ims_json_pack"))]:
        if pack:
            rows.append({
                "Source": name, "Rows": len(pack.std), "Invalid GSTIN": int((~pack.std["gstin_valid"]).sum()),
                "Missing Document No": int(pack.std["document_norm"].eq("").sum()),
                "Avg Data Quality": round(float(pack.std["data_quality_score"].mean()), 2),
            })
    if rows:
        st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
    else:
        st.info("No data uploaded yet.")
    save_user_state()


def active_ims_pack() -> Optional[UploadPack]:
    return st.session_state.get("ims_json_pack") if st.session_state.get("primary_ims_source") == "IMS JSON" else st.session_state.get("ims_pack")


def data_viewer_page() -> None:
    hero("IMS Data Viewer", "Inspect standardized Purchase and IMS records before reconciliation.")
    tabs = st.tabs(["Purchase Register", "IMS Utility", "IMS JSON", "Active IMS Combined"])
    packs = [st.session_state.get("purchase_pack"), st.session_state.get("ims_pack"), st.session_state.get("ims_json_pack"), active_ims_pack()]
    for tab, pack in zip(tabs, packs):
        with tab:
            if pack and not pack.std.empty:
                st.caption(f"{pack.label}: {len(pack.std):,} rows")
                st.dataframe(pack.std.head(2000), width="stretch", hide_index=True)
            else:
                st.info("No data available.")


def reconciliation_page() -> None:
    hero("Reconciliation Workspace", "Run exact and tolerance-based IMS reconciliation only when you click Start Reconciliation.")
    purchase = st.session_state.get("purchase_pack")
    ims = active_ims_pack()
    if not purchase or not ims:
        st.warning("Please upload Purchase Register and IMS Utility/JSON first.")
        return
    section("Reconciliation Controls", "Keep fuzzy matching off for large files; Phase 1 uses fast exact/normalized matching")
    c1, c2, c3 = st.columns(3)
    with c1: amount_tol = st.number_input("Amount Tolerance", min_value=0.0, value=2.0, step=1.0)
    with c2: date_tol = st.number_input("Date Tolerance Days", min_value=0, value=2, step=1)
    with c3: include_amendments = st.checkbox("Include amendment sheets", value=True)
    if st.button("▶ Start IMS Reconciliation", width="stretch", type="primary"):
        with st.spinner("Reconciling Purchase Register with IMS data..."):
            result = start_reconciliation(purchase.std, ims.std, amount_tol, int(date_tol), include_amendments)
            st.session_state.recon_result = result
            st.session_state.action_table = build_action_table(result)
            save_user_state()
        st.success(f"Reconciliation completed: {len(result):,} rows")

    result = st.session_state.get("recon_result", pd.DataFrame())
    if isinstance(result, pd.DataFrame) and not result.empty:
        section("Reconciliation Result", "Showing first 2,000 rows for speed; download full report from Reports page")
        summary = result.groupby("mismatch_type", dropna=False).size().reset_index(name="Count")
        st.dataframe(summary, width="stretch", hide_index=True)
        st.dataframe(result.head(2000), width="stretch", hide_index=True)


def build_action_table(recon: pd.DataFrame) -> pd.DataFrame:
    if recon.empty:
        return pd.DataFrame()
    cols = [
        "supplier_gstin", "supplier_name_purchase", "supplier_name_ims", "document_type_purchase", "document_type_ims",
        "document_no_purchase", "document_no_ims", "document_date_purchase", "document_date_ims",
        "invoice_value_purchase", "invoice_value_ims", "taxable_value_purchase", "taxable_value_ims",
        "igst_diff", "cgst_diff", "sgst_diff", "cess_diff", "total_tax_diff", "mismatch_type", "risk_level",
        "recommended_ims_action", "final_user_action", "reason", "vendor_followup_required", "confidence_score", "user_remarks",
    ]
    return recon[[c for c in cols if c in recon.columns]].copy()


def action_center_page() -> None:
    hero("IMS Action Center", "Review and finalize Accepted, Pending, Rejected and No Action status before preparing reports.")
    action = st.session_state.get("action_table", pd.DataFrame())
    if action.empty:
        st.info("Run reconciliation first.")
        return
    section("Action Summary", "Final actions can be reviewed and modified")
    st.dataframe(action.groupby(["final_user_action", "risk_level"], dropna=False).size().reset_index(name="Count"), width="stretch", hide_index=True)
    section("Bulk Actions", "Use carefully after reviewing risk category")
    b1, b2, b3, b4 = st.columns(4)
    with b1:
        if st.button("Accept Low-Risk Matched", width="stretch"):
            mask = action["risk_level"].eq("Low") & action["mismatch_type"].eq("Matched")
            action.loc[mask, "final_user_action"] = "Accepted"
    with b2:
        if st.button("Pending IMS Only", width="stretch"):
            action.loc[action["mismatch_type"].eq("Only in IMS"), "final_user_action"] = "Pending"
    with b3:
        if st.button("Pending High Risk", width="stretch"):
            action.loc[action["risk_level"].isin(["High", "Critical"]), "final_user_action"] = "Pending"
    with b4:
        if st.button("Clear to Recommended", width="stretch"):
            action["final_user_action"] = action["recommended_ims_action"]
    edited = st.data_editor(
        action.head(5000),
        width="stretch",
        hide_index=True,
        column_config={"final_user_action": st.column_config.SelectboxColumn("Final User Action", options=ACTIONS)},
        num_rows="fixed",
    )
    if st.button("Save Final Actions", width="stretch", type="primary"):
        # Merge visible edits back for first rows.
        action.update(edited)
        st.session_state.action_table = action
        save_user_state()
        log_event("Action Save", "Final IMS actions saved.")
        st.success("Actions saved.")


def risk_center_page() -> None:
    hero("Risk & Exception Center", "Focused review of high value, tax mismatch, credit note and amendment risks.")
    action = st.session_state.get("action_table", pd.DataFrame())
    if action.empty:
        st.info("Run reconciliation first.")
        return
    risk = action[action["risk_level"].isin(["High", "Critical"])].copy()
    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi("Critical", int((action["risk_level"] == "Critical").sum()), "Urgent", "⚠")
    with c2: kpi("High", int((action["risk_level"] == "High").sum()), "Review", "△")
    with c3: kpi("Tax Diff", f"{action['total_tax_diff'].abs().sum():,.2f}" if "total_tax_diff" in action else "0", "Absolute", "₹")
    with c4: kpi("Vendor Follow-up", int(action["vendor_followup_required"].sum()), "Cases", "✉")
    section("High Risk Exceptions", "Showing top 2,000 risky rows")
    st.dataframe(risk.head(2000), width="stretch", hide_index=True)


def vendor_page() -> None:
    hero("Vendor Follow-up", "Vendor-wise pending and mismatch cases for communication and tracking.")
    action = st.session_state.get("action_table", pd.DataFrame())
    if action.empty:
        st.info("Run reconciliation first.")
        return
    vendor_name_col = "supplier_name_purchase" if "supplier_name_purchase" in action.columns else "supplier_name_ims"
    follow = action[action["vendor_followup_required"] == True].copy()
    if follow.empty:
        st.success("No vendor follow-up cases found.")
        return
    summary = follow.groupby(["supplier_gstin", vendor_name_col], dropna=False).agg(
        Cases=("supplier_gstin", "size"),
        Tax_Diff=("total_tax_diff", lambda x: round(float(pd.to_numeric(x, errors="coerce").abs().sum()), 2)),
    ).reset_index().sort_values("Cases", ascending=False)
    st.dataframe(summary, width="stretch", hide_index=True)
    section("Email Draft", "Copy this draft and send to vendor")
    if not summary.empty:
        selected = st.selectbox("Select vendor", summary["supplier_gstin"].astype(str).tolist())
        vendor_cases = follow[follow["supplier_gstin"].astype(str) == selected]
        body = f"""Dear Vendor,\n\nDuring IMS reconciliation for {st.session_state.get('return_period','')}, we observed {len(vendor_cases)} document(s) requiring clarification for GSTIN {selected}.\n\nIssues include missing booking/value mismatch/tax head mismatch/date mismatch as per attached reconciliation details.\n\nPlease confirm and share corrected invoice/CN/DN details at the earliest so that IMS action can be completed timely.\n\nRegards,\n{st.session_state.get('prepared_by') or COPYRIGHT_OWNER}\n"""
        st.text_area("Vendor email draft", body, height=230)


def reports_page() -> None:
    hero("Reports & Export", "Download final IMS reports, action report and vendor follow-up files.")
    action = st.session_state.get("action_table", pd.DataFrame())
    recon = st.session_state.get("recon_result", pd.DataFrame())
    purchase = st.session_state.get("purchase_pack")
    ims = active_ims_pack()
    if action.empty:
        st.info("Run reconciliation first.")
        return
    section("Download Center", "Excel workbooks include @BAJRABHANU metadata and premium formatting")
    sheets = {
        "Final IMS Action": action,
        "Reco Detail": recon,
        "Accepted": action[action["final_user_action"] == "Accepted"],
        "Pending": action[action["final_user_action"] == "Pending"],
        "Rejected": action[action["final_user_action"] == "Rejected"],
        "No Action": action[action["final_user_action"] == "No Action"],
        "High Risk": action[action["risk_level"].isin(["High", "Critical"])],
    }
    if purchase: sheets["Std Purchase"] = purchase.std
    if ims: sheets["Std IMS"] = ims.std
    st.download_button("Download Complete IMS Reconciliation Workbook", data=excel_bytes(sheets), file_name="IMS_Recon_Pro_Final_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch", type="primary")
    st.caption("Phase 1 gives a complete action report. Phase 2 can write actions back into the GST Utility after exact column testing.")


def ai_page() -> None:
    hero("AI Insight Desk", "Rule-based AI-style assistant for IMS actions, ITC exposure and vendor follow-up.")
    action = st.session_state.get("action_table", pd.DataFrame())
    if action.empty:
        st.info("Run reconciliation first.")
        return
    query = st.text_input("Ask IMS Recon Pro", placeholder="Example: Which invoices should I accept? Which vendors need follow-up?")
    if query:
        q = query.lower()
        if "accept" in q:
            df = action[action["final_user_action"].eq("Accepted")]
            st.success(f"{len(df):,} document(s) are currently safe/marked for acceptance based on final action.")
            st.dataframe(df.head(500), width="stretch", hide_index=True)
        elif "vendor" in q or "follow" in q:
            df = action[action["vendor_followup_required"] == True]
            st.warning(f"{len(df):,} document(s) require vendor follow-up.")
            st.dataframe(df.head(500), width="stretch", hide_index=True)
        elif "risk" in q or "itc" in q:
            df = action[action["risk_level"].isin(["High", "Critical"])]
            st.error(f"{len(df):,} high/critical risk document(s) identified.")
            st.dataframe(df.head(500), width="stretch", hide_index=True)
        else:
            st.info("Try asking about accept, vendor follow-up, ITC risk, pending, rejected, or credit notes.")
    section("Auto Insights", "System generated recommendations")
    insights = [
        f"Accepted recommended/final: {(action['final_user_action']=='Accepted').sum():,}",
        f"Pending action cases: {(action['final_user_action']=='Pending').sum():,}",
        f"Vendor follow-up cases: {action['vendor_followup_required'].sum():,}",
        f"High/Critical risk cases: {action['risk_level'].isin(['High','Critical']).sum():,}",
    ]
    for item in insights:
        st.markdown(f"<span class='pill pill-ok'>✦ {escape(item)}</span>", unsafe_allow_html=True)


def admin_page() -> None:
    hero("Admin Panel", "Control data reset, audit trail and user-wise stored data.")
    section("Data Reset", "This deletes current user's uploaded/reconciled data from local database")
    confirm = st.text_input("Type DELETE to reset your complete data")
    if st.button("Reset My Complete Data", width="stretch"):
        if confirm == "DELETE":
            reset_current_user_data()
            st.success("Your data has been deleted.")
            st.rerun()
        else:
            st.error("Please type DELETE exactly.")
    section("Audit Log", "Recent user activity")
    try:
        with get_conn() as conn:
            log = pd.read_sql("SELECT * FROM audit_log ORDER BY id DESC LIMIT 500", conn)
        st.dataframe(log, width="stretch", hide_index=True)
    except Exception:
        st.info("No audit log available.")


def main() -> None:
    init_db()
    init_state()
    inject_css()
    if not st.session_state.get("logged_in"):
        login_page()
        return
    render_topbar()
    page = st.session_state.pop("nav_override", None) or sidebar()
    page_map = {
        "Dashboard": dashboard_page,
        "Client Setup": client_setup_page,
        "Upload Center": upload_center_page,
        "IMS Data Viewer": data_viewer_page,
        "Reconciliation Workspace": reconciliation_page,
        "IMS Action Center": action_center_page,
        "Risk & Exception Center": risk_center_page,
        "Vendor Follow-up": vendor_page,
        "Reports & Export": reports_page,
        "AI Insight Desk": ai_page,
        "Admin Panel": admin_page,
    }
    page_map.get(page, dashboard_page)()
    st.markdown(f"<br><div style='color:#64748B;font-size:.78rem;border-top:1px solid #D7E1EC;padding-top:.8rem;'>{APP_TITLE} | {APP_TAGLINE}<span style='float:right'>{COPYRIGHT_NOTICE}</span></div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
