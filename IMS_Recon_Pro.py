import json
import re
import sqlite3
import pickle
from dataclasses import dataclass
from datetime import datetime, date
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st


# =========================================================
# IMS RECON PRO — FULL WORKING VERSION
# Premium UI + Login + Upload + IMS Utility/JSON + Reco + Action + Export + DB
# Copyright @BAJRABHANU
# =========================================================

APP_TITLE = "IMS Recon Pro"
APP_TAGLINE = "Intelligent GST IMS Reconciliation & Action Management Platform"
COPYRIGHT_OWNER = "@BAJRABHANU"
APP_DB = "ims_recon_pro.db"
ENGINE_VERSION = "2026.05.01"

IMS_SHEETS = ["B2B", "B2BA", "B2B-DN", "B2B-DNA", "B2B-CN", "B2B-CNA"]
ACTION_VALUES = ["No Action", "Accepted", "Rejected", "Pending", "Review"]
USER_MASTER = {
    "MainAdmin": {"password": "Adminpwd", "role": "Main Admin", "name": "Main Admin"},
    "User1": {"password": "Userpwd1", "role": "Sub User", "name": "User One"},
    "User2": {"password": "Userpwd2", "role": "Sub User", "name": "User Two"},
}

MONEY_COLS = ["invoice_value", "taxable_value", "igst", "cgst", "sgst", "cess"]
TAX_COLS = ["igst", "cgst", "sgst", "cess"]

COLUMN_ALIASES = {
    "supplier_gstin": [
        "supplier gstin", "gstin of supplier", "gstin", "ctin", "counterparty gstin",
        "vendor gstin", "party gstin", "gstin/uın of supplier", "gstin/uin of supplier"
    ],
    "supplier_name": [
        "supplier name", "trade/legal name", "trade legal name", "legal name",
        "vendor name", "party name", "name", "supplier legal name"
    ],
    "document_type": [
        "document type", "doc type", "invoice type", "type", "supply type",
        "nature of document", "note type"
    ],
    "document_no": [
        "document number", "document no", "doc no", "invoice number", "invoice no",
        "invoice", "note number", "note no", "bill number", "voucher number"
    ],
    "document_date": [
        "document date", "doc date", "invoice date", "date", "note date", "bill date"
    ],
    "invoice_value": [
        "invoice value", "document value", "total invoice value", "gross value",
        "total value", "invoice value(inr)", "invoice value(rs)", "total document value"
    ],
    "taxable_value": [
        "taxable value", "taxable amount", "taxable value(inr)", "taxable value(rs)",
        "assessable value", "net value", "taxable val"
    ],
    "igst": ["igst", "integrated tax", "integrated tax amount", "igst amount"],
    "cgst": ["cgst", "central tax", "central tax amount", "cgst amount"],
    "sgst": ["sgst", "state tax", "state/ut tax", "utgst", "sgst amount"],
    "cess": ["cess", "cess amount"],
    "itc_available": ["itc available", "itc availability", "eligible itc", "itc eligibility", "eligible"],
    "ims_status": ["status", "ims status", "recipient status", "recipient action", "action"],
    "remarks": ["remarks", "remark", "reason", "comments", "comment"],
    "pos": ["place of supply", "pos", "state", "supply state"],
    "return_period": ["return period", "tax period", "period", "month"],
}


# =========================================================
# STREAMLIT PAGE
# =========================================================

st.set_page_config(
    page_title=f"{APP_TITLE} | {COPYRIGHT_OWNER}",
    page_icon="🇮🇳",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items=None,
)


# =========================================================
# DATABASE
# =========================================================

def get_conn():
    return sqlite3.connect(APP_DB, check_same_thread=False)


def init_db():
    with get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS user_data (
                username TEXT NOT NULL,
                key TEXT NOT NULL,
                value BLOB NOT NULL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY (username, key)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT,
                event_time TEXT,
                event_type TEXT,
                detail TEXT
            )
        """)


def db_save(username: str, key: str, value):
    blob = pickle.dumps(value)
    with get_conn() as conn:
        conn.execute(
            "INSERT OR REPLACE INTO user_data (username, key, value, updated_at) VALUES (?, ?, ?, ?)",
            (username, key, blob, datetime.now().isoformat(timespec="seconds")),
        )


def db_load(username: str, key: str, default=None):
    try:
        with get_conn() as conn:
            row = conn.execute("SELECT value FROM user_data WHERE username=? AND key=?", (username, key)).fetchone()
        if not row:
            return default
        return pickle.loads(row[0])
    except Exception:
        return default


def db_delete_user(username: str):
    with get_conn() as conn:
        conn.execute("DELETE FROM user_data WHERE username=?", (username,))
        conn.execute("DELETE FROM audit_log WHERE username=?", (username,))


def log_event(event_type: str, detail: str):
    username = st.session_state.get("username", "")
    try:
        with get_conn() as conn:
            conn.execute(
                "INSERT INTO audit_log (username, event_time, event_type, detail) VALUES (?, ?, ?, ?)",
                (username, datetime.now().isoformat(timespec="seconds"), event_type, detail),
            )
    except Exception:
        pass


def load_audit(username: str = "") -> pd.DataFrame:
    try:
        with get_conn() as conn:
            if username:
                return pd.read_sql(
                    "SELECT event_time, event_type, detail FROM audit_log WHERE username=? ORDER BY id DESC",
                    conn,
                    params=(username,),
                )
            return pd.read_sql(
                "SELECT username, event_time, event_type, detail FROM audit_log ORDER BY id DESC",
                conn,
            )
    except Exception:
        return pd.DataFrame()


# =========================================================
# SESSION
# =========================================================

def init_state():
    defaults = {
        "logged_in": False,
        "username": "",
        "role": "",
        "display_name": "",
        "page": "Dashboard",
        "client_name": "",
        "client_gstin": "",
        "return_period": datetime.today().strftime("%b-%Y"),
        "purchase_df": pd.DataFrame(),
        "ims_df": pd.DataFrame(),
        "ims_source": "",
        "recon_df": pd.DataFrame(),
        "action_df": pd.DataFrame(),
        "amount_tolerance": 5.0,
        "date_tolerance": 2,
        "include_amendments": True,
        "use_fuzzy": False,
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)


def load_user_state():
    username = st.session_state.get("username")
    if not username:
        return
    for key in [
        "client_name", "client_gstin", "return_period",
        "purchase_df", "ims_df", "ims_source", "recon_df", "action_df"
    ]:
        st.session_state[key] = db_load(username, key, st.session_state.get(key))


def save_user_state(keys: Optional[List[str]] = None):
    username = st.session_state.get("username")
    if not username:
        return
    keys = keys or [
        "client_name", "client_gstin", "return_period",
        "purchase_df", "ims_df", "ims_source", "recon_df", "action_df"
    ]
    for key in keys:
        db_save(username, key, st.session_state.get(key))


# =========================================================
# STYLING
# =========================================================

def inject_css():
    st.markdown("""
    <style>
        :root {
            --gst-dark:#102965;
            --gst-darker:#071b4a;
            --gst-nav:#2f578f;
            --gst-nav-active:#21c7bc;
            --page-bg:#dbe8f7;
            --panel:#ffffff;
            --panel-soft:#f4f8fd;
            --border:#c7d8ec;
            --text:#0c2148;
            --muted:#536986;
            --gold:#f0b429;
            --blue:#1f6fd1;
            --green:#138808;
            --red:#d94332;
        }

        .stApp {
            background:
                linear-gradient(180deg, #e7f0fb 0%, #d6e5f5 38%, #cfe1f3 100%) !important;
        }

        header[data-testid="stHeader"] {
            visibility: visible !important;
            height: auto !important;
            background: rgba(219,232,247,0.92) !important;
            backdrop-filter: blur(8px);
            border-bottom:1px solid rgba(7,27,74,0.08);
        }
        div[data-testid="stToolbar"] { visibility:hidden; height:0; }

        .block-container {
            padding-top: 1rem;
            padding-bottom: 1.5rem;
            max-width: 1260px;
        }

        .gst-shell {
            border-radius: 0 0 22px 22px;
            overflow:hidden;
            box-shadow:0 14px 34px rgba(7,27,74,0.16);
            border:1px solid rgba(255,255,255,0.45);
            margin-bottom:18px;
            background:#ffffff;
        }
        .gst-top-strip {
            background:#071b4a;
            color:#ffffff;
            min-height:28px;
            display:flex;
            justify-content:flex-end;
            align-items:center;
            gap:18px;
            padding:4px 18px;
            font-size:13px;
            font-weight:700;
        }
        .gst-masthead {
            background: linear-gradient(90deg, #112965 0%, #0e2c68 58%, #12346f 100%);
            color:white;
            display:flex;
            align-items:center;
            justify-content:space-between;
            gap:22px;
            padding:22px 30px;
            min-height:118px;
        }
        .gst-brand {
            display:flex;
            align-items:center;
            gap:20px;
            min-width:0;
        }
        .gst-emblem {
            width:72px;
            height:72px;
            border-radius:10px;
            display:flex;
            align-items:center;
            justify-content:center;
            color:#ffffff;
            font-size:44px;
            background:rgba(255,255,255,0.08);
            border:1px solid rgba(255,255,255,0.16);
        }
        .gst-title {
            font-size:34px;
            font-weight:500;
            letter-spacing:.2px;
            line-height:1.08;
            color:#ffffff;
        }
        .gst-subtitle {
            margin-top:8px;
            font-size:20px;
            line-height:1.25;
            color:#ffffff;
            font-weight:400;
        }
        .gst-action-wrap {
            display:flex;
            gap:12px;
            align-items:center;
            justify-content:flex-end;
            flex-wrap:wrap;
        }
        .gst-login-btn {
            min-width:102px;
            text-align:center;
            background:#ffffff;
            color:#102965;
            border:1px solid rgba(255,255,255,0.6);
            border-radius:4px;
            padding:10px 16px;
            font-size:14px;
            font-weight:800;
            letter-spacing:.3px;
        }
        .gst-meta-row {
            background:#eef5ff;
            border-top:1px solid #c6d6eb;
            padding:14px 18px;
            display:grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap:12px;
        }
        .gst-meta-card {
            background:#ffffff;
            border:1px solid #c7d8ec;
            border-radius:10px;
            min-height:62px;
            padding:11px 14px;
            display:flex;
            align-items:center;
            gap:12px;
        }
        .gst-meta-icon {
            width:38px;
            height:38px;
            border-radius:8px;
            background:#edf4ff;
            display:flex;
            align-items:center;
            justify-content:center;
            font-size:18px;
            flex-shrink:0;
        }
        .gst-meta-label {
            color:#5e7390;
            font-size:12px;
            margin-bottom:2px;
        }
        .gst-meta-value {
            color:#071b4a;
            font-weight:800;
            font-size:15px;
            line-height:1.25;
            overflow-wrap:anywhere;
        }

        .gst-nav-title {
            background:#2f578f;
            color:#fff;
            font-size:16px;
            font-weight:800;
            padding:12px 16px;
            border-radius:12px 12px 0 0;
            border:1px solid rgba(255,255,255,0.18);
            border-bottom:0;
        }
        .gst-nav-panel {
            background:#ffffff;
            border:1px solid #b9cce4;
            border-radius:0 0 14px 14px;
            padding:12px;
            margin-bottom:18px;
            box-shadow:0 10px 22px rgba(7,27,74,0.08);
        }

        div[data-testid="stButton"] > button {
            border-radius:6px;
            border:1px solid #2e5a95;
            background:#2f578f;
            color:#ffffff;
            min-height:46px;
            font-weight:700;
            box-shadow:none;
        }
        div[data-testid="stButton"] > button:hover {
            background:#21c7bc;
            color:#071b4a;
            border-color:#21c7bc;
        }
        div[data-testid="stButton"] > button[kind="primary"] {
            background:#102965;
            color:#ffffff;
            border-color:#102965;
        }

        .main-shell,
        .panel,
        .metric-card,
        .small-card {
            background:#ffffff;
            border:1px solid var(--border);
            box-shadow:0 12px 28px rgba(7,27,74,.10);
        }
        .main-shell {
            border-radius:16px;
            overflow:hidden;
            margin-bottom:18px;
        }
        .panel, .metric-card, .small-card {
            border-radius:14px;
            padding:18px 20px;
            height:100%;
        }
        .content-pad {padding:26px; position:relative;}
        .watermark {
            position:absolute; left:50%; top:46%; transform:translate(-50%,-50%);
            font-size:190px; color:rgba(14,41,90,.035); pointer-events:none;
        }
        .headline {font-size:20px;color:#0d4f9b;font-weight:800;}
        .main-title {font-size:30px;font-weight:900;color:#071b4a;line-height:1.24;margin-top:8px;}
        .subcopy {font-size:16px;color:#52637d;margin-top:10px;line-height:1.5;}
        .cta-dark,.cta-light {
            display:inline-block;padding:12px 22px;border-radius:8px;font-weight:800;text-decoration:none;font-size:15px;margin-right:10px;margin-top:18px;
        }
        .cta-dark {background:#102965;color:white;box-shadow:0 10px 18px rgba(11,42,93,.18);}
        .cta-light {background:white;color:#102965;border:1px solid #b9cce4;}

        .metric-top {display:flex;align-items:center;gap:14px;}
        .metric-icon {width:54px;height:54px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:24px;}
        .metric-label {font-size:14px;color:#59708d;}
        .metric-value {font-size:32px;font-weight:900;color:#071b4a;line-height:1.15;}
        .metric-delta {font-size:13px;color:#128047;margin-top:6px;}
        .metric-delta.red {color:#d94332;}
        .panel-title {font-size:18px;font-weight:900;color:#071b4a;}
        .feature-card {background:#f7fbff;border:1px solid #c7d8ec;border-radius:12px;padding:14px 16px;margin-bottom:12px;}
        .feature-card.blue {background:#f2f7ff;border-color:#c7d8ec;}
        .feature-card.green {background:#f4fbf5;border-color:#cce5d1;}
        .feature-title {font-weight:800;color:#0d2d63;font-size:16px;}
        .feature-desc {font-size:13px;color:#5f6f89;line-height:1.4;margin-top:4px;}
        .shield-center {width:130px;height:130px;border-radius:18px;margin:0 auto 18px auto;background:linear-gradient(135deg,#edf4ff,#ffffff);display:flex;align-items:center;justify-content:center;font-size:56px;box-shadow:0 10px 24px rgba(16,34,68,.12);border:1px solid #c7d8ec;}
        .section-title {font-size:26px;font-weight:900;color:#071b4a;margin:8px 0 4px 0;}
        .section-sub {font-size:14px;color:#60748f;margin-bottom:18px;}

        .login-bg {
            min-height:calc(100vh - 40px); display:flex; align-items:center; justify-content:center;
            background:linear-gradient(135deg,#071b4a,#12346f);
            border-radius:18px; position:relative; overflow:hidden;
        }
        .login-card {
            width:460px; background:rgba(255,255,255,.97); border:1px solid rgba(255,255,255,.7);
            border-radius:16px; padding:34px; box-shadow:0 30px 80px rgba(0,0,0,.28);
        }
        .login-title {font-size:34px;font-weight:900;color:#071b4a;text-align:center;}
        .login-sub {font-size:15px;color:#566982;text-align:center;margin-bottom:22px;}
        .copyright-float {display:none;}

        .footer-bar {
            margin-top:18px;
            border-radius:14px;
            background:linear-gradient(90deg,#071b4a 0%,#102965 48%,#071b4a 100%);
            color:white;
            padding:18px 22px;
            box-shadow:0 12px 28px rgba(7,26,61,0.14);
        }
        .foot-item {display:flex;align-items:center;gap:10px;justify-content:center;}
        .foot-main {font-weight:800;}
        .foot-sub {font-size:13px;color:#d4e0ff;}

        .stTextInput input, .stNumberInput input, .stDateInput input, .stSelectbox div[data-baseweb="select"] > div {
            border-radius:8px !important;
        }
        .stForm {
            background:#f8fbff;
            border:1px solid #c7d8ec;
            border-radius:14px;
            padding:16px;
        }

        @media (max-width: 1100px) {
            .gst-masthead {flex-direction:column; align-items:flex-start;}
            .gst-action-wrap {justify-content:flex-start;}
            .gst-meta-row {grid-template-columns: repeat(2, minmax(0, 1fr));}
            .gst-title {font-size:30px;}
            .gst-subtitle {font-size:18px;}
        }
        @media (max-width: 760px) {
            .gst-brand {align-items:flex-start;}
            .gst-emblem {width:58px;height:58px;font-size:34px;}
            .gst-title {font-size:25px;}
            .gst-subtitle {font-size:15px;}
            .gst-meta-row {grid-template-columns: 1fr;}
            .block-container {padding-top:0.7rem;}
        }
    </style>
    """, unsafe_allow_html=True)



# =========================================================
# DATA PROCESSING
# =========================================================

def clean_header(value) -> str:
    text = str(value or "").strip().lower()
    text = text.replace("\n", " ").replace("\r", " ")
    text = text.replace("₹", "rs")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^a-z0-9%/() -]", "", text)
    return text.strip()


def find_col(df: pd.DataFrame, logical: str) -> Optional[str]:
    aliases = COLUMN_ALIASES.get(logical, [logical])
    norm = {clean_header(c): c for c in df.columns}
    for alias in aliases:
        ca = clean_header(alias)
        if ca in norm:
            return norm[ca]
    for alias in aliases:
        ca = clean_header(alias)
        for nc, orig in norm.items():
            if ca and ca in nc:
                return orig
    return None


def normalize_doc_no(x) -> str:
    text = str(x or "").strip().upper()
    text = re.sub(r"\.0$", "", text)
    return re.sub(r"[^A-Z0-9]", "", text)


def normalize_gstin(x) -> str:
    text = str(x or "").strip().upper()
    return re.sub(r"[^A-Z0-9]", "", text)


def validate_gstin(x) -> bool:
    return bool(re.match(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$", normalize_gstin(x)))


def to_number(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series(dtype="float64")
    clean = (
        s.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("INR", "", case=False, regex=False)
        .str.replace("RS.", "", case=False, regex=False)
        .str.replace("RS", "", case=False, regex=False)
        .str.replace("₹", "", regex=False)
        .str.strip()
        .str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    )
    clean = clean.replace({"": "0", "nan": "0", "None": "0", "NaT": "0"})
    return pd.to_numeric(clean, errors="coerce").fillna(0.0)


def to_date(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series(dtype="datetime64[ns]")
    numeric = pd.to_numeric(s, errors="coerce")
    excel_dates = pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
    parsed = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return parsed.fillna(excel_dates)


def sign_for_doc_type(x) -> int:
    text = str(x or "").lower()
    if any(k in text for k in ["credit", "cn", "cdn", "refund"]):
        return -1
    return 1


def detect_header_row(raw: pd.DataFrame) -> int:
    alias_words = [clean_header(a) for v in COLUMN_ALIASES.values() for a in v if len(clean_header(a)) >= 3]
    best_idx = raw.index[0]
    best_score = -1
    for idx, row in raw.head(30).iterrows():
        row_text = " ".join(clean_header(c) for c in row.tolist())
        score = sum(1 for a in alias_words if a and a in row_text)
        if score > best_score:
            best_score = score
            best_idx = idx
    return int(best_idx)


def normalize_sheet(sheet_df: pd.DataFrame) -> pd.DataFrame:
    raw = sheet_df.dropna(how="all").dropna(how="all", axis=1)
    if raw.empty:
        return pd.DataFrame()
    header_idx = detect_header_row(raw)
    header_pos = list(raw.index).index(header_idx)
    headers = []
    seen = {}
    for i, h in enumerate(raw.loc[header_idx].fillna("").astype(str).tolist(), start=1):
        name = str(h or "").strip() or f"Column {i}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        headers.append(name)
    body = raw.iloc[header_pos + 1:].copy()
    body.columns = headers
    body = body.dropna(how="all")
    body = body.loc[:, [c for c in body.columns if str(c).strip()]]
    return body


def read_excel_all_sheets(file, wanted_sheets: Optional[List[str]] = None) -> Dict[str, pd.DataFrame]:
    sheets = pd.read_excel(file, sheet_name=None, dtype=object, header=None, engine="openpyxl")
    out = {}
    for name, df in sheets.items():
        if wanted_sheets and name.strip().upper() not in [s.upper() for s in wanted_sheets]:
            continue
        clean = normalize_sheet(df)
        if not clean.empty:
            out[name.strip()] = clean
    return out


def standardize(df: pd.DataFrame, source_label: str, sheet_name: str = "", default_doc_type: str = "") -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    schema = {k: find_col(df, k) for k in COLUMN_ALIASES}
    n = len(df)

    def get(logical, default=""):
        col = schema.get(logical)
        if col and col in df.columns:
            return df[col]
        return pd.Series([default] * n, index=df.index)

    out = pd.DataFrame(index=df.index)
    out["supplier_gstin"] = get("supplier_gstin", "").map(normalize_gstin)
    out["supplier_name"] = get("supplier_name", "").astype(str).str.strip()
    out["document_type"] = get("document_type", default_doc_type).astype(str).replace({"": default_doc_type})
    out["document_no"] = get("document_no", "").astype(str).str.strip()
    out["document_norm"] = out["document_no"].map(normalize_doc_no)
    out["document_date"] = to_date(get("document_date", pd.NaT))
    out["invoice_value"] = to_number(get("invoice_value", 0))
    out["taxable_value"] = to_number(get("taxable_value", 0))
    out["igst"] = to_number(get("igst", 0))
    out["cgst"] = to_number(get("cgst", 0))
    out["sgst"] = to_number(get("sgst", 0))
    out["cess"] = to_number(get("cess", 0))
    out["total_tax"] = out[TAX_COLS].sum(axis=1)
    out["itc_available"] = get("itc_available", "Yes").astype(str)
    out["ims_status"] = get("ims_status", "No Action").astype(str)
    out["remarks"] = get("remarks", "").astype(str)
    out["pos"] = get("pos", "").astype(str)
    out["return_period"] = get("return_period", "").astype(str)
    out["source"] = source_label
    out["ims_sheet"] = sheet_name
    out["gstin_valid"] = out["supplier_gstin"].map(validate_gstin)
    out["data_quality"] = out.apply(row_quality_score, axis=1)

    sign = out["document_type"].map(sign_for_doc_type)
    for c in MONEY_COLS + ["total_tax"]:
        out[c] = out[c] * sign.where(out[c] >= 0, 1)

    out = out[
        ["source", "ims_sheet", "supplier_gstin", "supplier_name", "document_type", "document_no",
         "document_norm", "document_date", "invoice_value", "taxable_value", "igst", "cgst", "sgst", "cess",
         "total_tax", "itc_available", "ims_status", "remarks", "pos", "return_period", "gstin_valid",
         "data_quality"]
    ]
    out = out[(out["supplier_gstin"].astype(str).str.len() > 0) | (out["document_norm"].astype(str).str.len() > 0)]
    return out.reset_index(drop=True)


def row_quality_score(row) -> int:
    score = 100
    if not validate_gstin(row.get("supplier_gstin", "")) and str(row.get("supplier_gstin", "")).strip():
        score -= 25
    if not str(row.get("document_norm", "")).strip():
        score -= 30
    if pd.isna(pd.to_datetime(row.get("document_date"), errors="coerce")):
        score -= 15
    if abs(float(row.get("taxable_value", 0) or 0)) < 0.01 and abs(float(row.get("total_tax", 0) or 0)) > 0.01:
        score -= 20
    return max(0, min(100, score))


def read_purchase_file(file) -> pd.DataFrame:
    if file is None:
        return pd.DataFrame()
    suffix = Path(file.name).suffix.lower()
    if suffix == ".csv":
        raw = pd.read_csv(file, dtype=object)
        return standardize(raw, "Purchase Register")
    sheets = read_excel_all_sheets(file)
    frames = []
    for sheet, raw in sheets.items():
        std = standardize(raw, "Purchase Register", sheet)
        if not std.empty:
            frames.append(std)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def infer_doc_type_from_sheet(sheet: str) -> str:
    s = sheet.upper()
    if "CN" in s:
        return "Credit Note"
    if "DN" in s:
        return "Debit Note"
    if "A" in s and "B2B" in s:
        return "Amended Invoice"
    return "Invoice"


def read_ims_utility(file) -> pd.DataFrame:
    if file is None:
        return pd.DataFrame()
    sheets = read_excel_all_sheets(file, IMS_SHEETS)
    frames = []
    for sheet, raw in sheets.items():
        default_doc = infer_doc_type_from_sheet(sheet)
        std = standardize(raw, "IMS Utility", sheet, default_doc)
        if not std.empty:
            frames.append(std)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def flatten_json(obj, parent_key="", sep="_"):
    items = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            new_key = f"{parent_key}{sep}{k}" if parent_key else str(k)
            items.extend(flatten_json(v, new_key, sep=sep).items())
    elif isinstance(obj, list):
        if all(isinstance(i, dict) for i in obj):
            return {parent_key: obj}
        return {parent_key: obj}
    else:
        return {parent_key: obj}
    return dict(items)


def extract_records_from_json(obj) -> List[dict]:
    records = []
    def walk(x, path=""):
        if isinstance(x, list):
            if x and all(isinstance(i, dict) for i in x):
                for item in x:
                    flat = flatten_json(item)
                    flat["_json_path"] = path
                    records.append(flat)
            else:
                for i, item in enumerate(x):
                    walk(item, f"{path}.{i}")
        elif isinstance(x, dict):
            for k, v in x.items():
                walk(v, f"{path}.{k}" if path else k)
    walk(obj)
    return records


def read_ims_json(file) -> pd.DataFrame:
    if file is None:
        return pd.DataFrame()
    data = json.load(file)
    records = extract_records_from_json(data)
    if not records:
        return pd.DataFrame()
    raw = pd.DataFrame(records)
    # Rename likely JSON fields into friendlier headings by fuzzy search
    return standardize(raw, "IMS JSON", "JSON")


def aggregate(df: pd.DataFrame, label: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    work = df.copy()
    work = work[(work["supplier_gstin"].astype(str) != "") & (work["document_norm"].astype(str) != "")]
    if work.empty:
        return pd.DataFrame()

    agg_map = {
        "supplier_name": "first",
        "document_type": "first",
        "document_no": "first",
        "document_date": "min",
        "invoice_value": "sum",
        "taxable_value": "sum",
        "igst": "sum",
        "cgst": "sum",
        "sgst": "sum",
        "cess": "sum",
        "total_tax": "sum",
        "itc_available": "first",
        "ims_status": "first",
        "remarks": "first",
        "ims_sheet": "first",
        "data_quality": "mean",
    }
    out = work.groupby(["supplier_gstin", "document_norm"], dropna=False).agg(agg_map).reset_index()
    rename = {col: f"{col}_{label}" for col in out.columns if col not in ["supplier_gstin", "document_norm"]}
    return out.rename(columns=rename)


def calculate_recon(purchase: pd.DataFrame, ims: pd.DataFrame, amount_tol: float, date_tol: int, include_amendments: bool) -> pd.DataFrame:
    if purchase.empty or ims.empty:
        return pd.DataFrame()

    ims_work = ims.copy()
    if not include_amendments and "ims_sheet" in ims_work.columns:
        ims_work = ims_work[~ims_work["ims_sheet"].astype(str).str.upper().isin(["B2BA", "B2B-DNA", "B2B-CNA"])]

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
    ponly = m["_merge"].eq("left_only")
    ionly = m["_merge"].eq("right_only")

    amount_ok = (
        m["taxable_value_diff"].abs().le(amount_tol)
        & m["total_tax_diff"].abs().le(amount_tol)
        & m["invoice_value_diff"].abs().le(max(amount_tol, 1))
    )
    tax_head_ok = (
        m["igst_diff"].abs().le(amount_tol)
        & m["cgst_diff"].abs().le(amount_tol)
        & m["sgst_diff"].abs().le(amount_tol)
        & m["cess_diff"].abs().le(amount_tol)
    )
    date_ok = m["date_diff_days"].fillna(0).le(date_tol)

    m["mismatch_type"] = "Matched"
    m.loc[ponly, "mismatch_type"] = "Only in Purchase Register"
    m.loc[ionly, "mismatch_type"] = "Only in IMS"
    m.loc[both & amount_ok & ~tax_head_ok, "mismatch_type"] = "Tax Head Mismatch"
    m.loc[both & ~amount_ok & date_ok, "mismatch_type"] = "Value / Tax Mismatch"
    m.loc[both & amount_ok & tax_head_ok & ~date_ok, "mismatch_type"] = "Date Mismatch"
    m.loc[both & ~amount_ok & ~date_ok, "mismatch_type"] = "Value and Date Mismatch"

    m["risk_score"] = m.apply(risk_score, axis=1)
    m["risk_level"] = m["risk_score"].map(risk_level)
    m["recommended_action"] = m.apply(recommend_action, axis=1)
    m["reason"] = m.apply(recommend_reason, axis=1)
    m["vendor_followup_required"] = m["mismatch_type"].isin(["Only in Purchase Register", "Value / Tax Mismatch", "Tax Head Mismatch", "Value and Date Mismatch", "Only in IMS"])
    m["final_user_action"] = m["recommended_action"]
    m["user_remarks"] = ""
    m["confidence_score"] = m.apply(confidence_score, axis=1)

    # Presentation columns
    m["supplier_name"] = m.get("supplier_name_purchase").fillna(m.get("supplier_name_ims"))
    m["document_type"] = m.get("document_type_purchase").fillna(m.get("document_type_ims"))
    m["document_no"] = m.get("document_no_purchase").fillna(m.get("document_no_ims"))
    m["document_date"] = pdate.fillna(idate)

    priority_order = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}
    m["_risk_order"] = m["risk_level"].map(priority_order).fillna(9)
    m = m.sort_values(["_risk_order", "supplier_gstin", "document_norm"]).drop(columns=["_risk_order"])
    return m.reset_index(drop=True)


def risk_score(row) -> int:
    score = 0
    typ = str(row.get("mismatch_type", ""))
    total_tax_diff = abs(float(row.get("total_tax_diff", 0) or 0))
    taxable_diff = abs(float(row.get("taxable_value_diff", 0) or 0))
    if typ == "Only in IMS":
        score += 30
    if typ == "Only in Purchase Register":
        score += 25
    if "Value" in typ:
        score += 25
    if "Tax Head" in typ:
        score += 20
    if "Date" in typ:
        score += 10
    if total_tax_diff >= 100000 or taxable_diff >= 500000:
        score += 25
    if "CN" in str(row.get("ims_sheet_ims", "")).upper() or "credit" in str(row.get("document_type_ims", "")).lower():
        score += 15
    return min(100, score)


def risk_level(score) -> str:
    score = int(score or 0)
    if score >= 75:
        return "Critical"
    if score >= 50:
        return "High"
    if score >= 21:
        return "Medium"
    return "Low"


def recommend_action(row) -> str:
    typ = str(row.get("mismatch_type", ""))
    itc = str(row.get("itc_available_purchase", "Yes")).lower()
    if "no" in itc or "ineligible" in itc or "not" in itc:
        return "Rejected"
    if typ == "Matched":
        return "Accepted"
    if typ == "Only in IMS":
        return "Pending"
    if typ == "Only in Purchase Register":
        return "No Action"
    if typ in ["Value / Tax Mismatch", "Tax Head Mismatch", "Value and Date Mismatch", "Date Mismatch"]:
        return "Pending"
    return "Review"


def recommend_reason(row) -> str:
    typ = str(row.get("mismatch_type", ""))
    if typ == "Matched":
        return "Purchase Register and IMS values are matched within selected tolerance."
    if typ == "Only in IMS":
        return "Document appears in IMS but is not found in Purchase Register. Keep Pending until booking/vendor confirmation."
    if typ == "Only in Purchase Register":
        return "Document exists in books but not in IMS. No IMS action possible; follow up with vendor if ITC expected."
    if typ == "Tax Head Mismatch":
        return "Total tax may be close but IGST/CGST/SGST/Cess split differs. Check POS and tax head classification."
    if typ == "Value / Tax Mismatch":
        return "Amount difference detected. Compare invoice copy, credit/debit note treatment and amendment."
    if typ == "Date Mismatch":
        return "Invoice matched by GSTIN and document number, but document date differs beyond tolerance."
    if typ == "Value and Date Mismatch":
        return "Both amount and date differences detected. Manual review required before IMS action."
    return "Review required."


def confidence_score(row) -> int:
    typ = str(row.get("mismatch_type", ""))
    if typ == "Matched":
        return 100
    if typ == "Date Mismatch":
        return 82
    if typ == "Tax Head Mismatch":
        return 75
    if typ == "Value / Tax Mismatch":
        return 65
    if typ.startswith("Only"):
        return 35
    return 50


def recon_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    return df.groupby(["mismatch_type", "risk_level", "recommended_action"], dropna=False).agg(
        Records=("mismatch_type", "size"),
        Taxable_Diff=("taxable_value_diff", "sum"),
        Tax_Diff=("total_tax_diff", "sum"),
        Avg_Confidence=("confidence_score", "mean"),
    ).reset_index().round(2)


def to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        meta = pd.DataFrame([
            {"Field": "Application", "Value": APP_TITLE},
            {"Field": "Owner", "Value": COPYRIGHT_OWNER},
            {"Field": "Engine Version", "Value": ENGINE_VERSION},
            {"Field": "Generated On", "Value": datetime.now().strftime("%d-%b-%Y %H:%M:%S")},
            {"Field": "Client", "Value": st.session_state.get("client_name", "")},
            {"Field": "GSTIN", "Value": st.session_state.get("client_gstin", "")},
            {"Field": "Return Period", "Value": st.session_state.get("return_period", "")},
            {"Field": "Generated By", "Value": st.session_state.get("username", "")},
        ])
        meta.to_excel(writer, index=False, sheet_name="BAJRABHANU")
        used = {"BAJRABHANU"}
        for sheet_name, df in sheets.items():
            safe = re.sub(r"[\[\]:*?/\\]", "_", str(sheet_name))[:31] or "Sheet"
            base = safe
            i = 1
            while safe in used:
                safe = f"{base[:27]}_{i}"
                i += 1
            used.add(safe)
            data = df if isinstance(df, pd.DataFrame) and not df.empty else pd.DataFrame({"Message": ["No data available"]})
            data.to_excel(writer, index=False, sheet_name=safe)
            ws = writer.sheets[safe]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col_cells in ws.columns:
                max_len = 0
                letter = col_cells[0].column_letter
                for cell in col_cells:
                    max_len = max(max_len, len(str(cell.value)) if cell.value is not None else 0)
                ws.column_dimensions[letter].width = min(max_len + 2, 38)
    return buffer.getvalue()


def safe_display_df(df: pd.DataFrame, limit: int = 1000) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.head(limit).copy()
    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = out[col].dt.strftime("%d-%b-%Y").fillna("")
    return out


# =========================================================
# UI HELPERS
# =========================================================

def top_header():
    client_name = st.session_state.get("client_name", "") or "Not set"
    client_gstin = st.session_state.get("client_gstin", "") or "GSTIN pending"
    st.markdown(f"""
    <div class='gst-shell'>
        <div class='gst-top-strip'>
            <span>Skip to Main Content</span>
            <span>◐</span>
            <span>A+</span>
            <span>A-</span>
        </div>
        <div class='gst-masthead'>
            <div class='gst-brand'>
                <div class='gst-emblem'>🦁</div>
                <div>
                    <div class='gst-title'>Goods and Services Tax</div>
                    <div class='gst-subtitle'>IMS Recon Pro — GST Reconciliation Platform</div>
                    <div class='header-note'>Government-style compliance workspace for Purchase Register vs IMS reconciliation</div>
                </div>
            </div>
            <div class='gst-action-wrap'>
                <div class='gst-login-btn'>REGISTER</div>
                <div class='gst-login-btn'>LOGIN</div>
            </div>
        </div>
        <div class='gst-meta-row'>
            <div class='gst-meta-card'>
                <div class='gst-meta-icon'>🗓️</div>
                <div><div class='gst-meta-label'>Today</div><div class='gst-meta-value'>{datetime.today().strftime("%d %b %Y")} • {datetime.today().strftime("%A")}</div></div>
            </div>
            <div class='gst-meta-card'>
                <div class='gst-meta-icon'>👤</div>
                <div><div class='gst-meta-label'>Logged in user</div><div class='gst-meta-value'>{st.session_state.get("display_name", "User")} • {st.session_state.get("role", "")}</div></div>
            </div>
            <div class='gst-meta-card'>
                <div class='gst-meta-icon'>🏢</div>
                <div><div class='gst-meta-label'>Client / GSTIN</div><div class='gst-meta-value'>{client_name} • {client_gstin}</div></div>
            </div>
            <div class='gst-meta-card'>
                <div class='gst-meta-icon'>©</div>
                <div><div class='gst-meta-label'>Copyright</div><div class='gst-meta-value'>{COPYRIGHT_OWNER} • IMS Recon Pro</div></div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def horizontal_nav():
    pages = [
        ("🏠", "Dashboard"),
        ("🔐", "Client Setup"),
        ("📤", "Upload Center"),
        ("🧾", "IMS Data Viewer"),
        ("🔄", "Reconciliation Workspace"),
        ("✅", "Action Center"),
        ("⚠️", "Risk Center"),
        ("📨", "Vendor Follow-up"),
        ("📊", "Reports & Export"),
        ("🧠", "AI Insight Desk"),
        ("👑", "Admin Panel"),
    ]
    st.markdown("<div class='gst-nav-title'>GST IMS Services</div>", unsafe_allow_html=True)
    st.markdown("<div class='gst-nav-panel'>", unsafe_allow_html=True)
    row1 = pages[:6]
    row2 = pages[6:]
    cols = st.columns(len(row1))
    for col, (icon, page) in zip(cols, row1):
        with col:
            label = f"{icon} {page}"
            if st.button(label, key=f"nav_top_{page}", use_container_width=True):
                st.session_state.page = page
                st.rerun()
    cols = st.columns(len(row2))
    for col, (icon, page) in zip(cols, row2):
        with col:
            label = f"{icon} {page}"
            if st.button(label, key=f"nav_top_{page}", use_container_width=True):
                st.session_state.page = page
                st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)


def hero_dashboard():
    st.markdown("<div class='main-shell'><div class='content-pad'><div class='watermark'>◉</div>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([3.8, 1.8, 1.4])
    with col1:
        st.markdown(f"""
        <div class='headline'>☀️ Namaste, {st.session_state.get("display_name", "User")}! 🙏</div>
        <div class='main-title'>Reconcile Today. Stay Compliant.<br>Drive Confidence.</div>
        <div class='subcopy'>AI-powered IMS reconciliation with accuracy,<br>automation & actionable insights.</div>
        <a class='cta-dark'>Go to Workspace →</a><a class='cta-light'>☁️ &nbsp; Upload IMS Data</a>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("<div class='shield-center'>🛡️</div>", unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class='feature-card'><div class='feature-title'>🧠 Smart Reconciliation</div><div class='feature-desc'>AI-driven matching with high accuracy</div></div>
        <div class='feature-card blue'><div class='feature-title'>🛡️ Risk Detection</div><div class='feature-desc'>Identify mismatches & compliance risks</div></div>
        <div class='feature-card green'><div class='feature-title'>📈 Actionable Insights</div><div class='feature-desc'>Real-time dashboards for better decisions</div></div>
        """, unsafe_allow_html=True)
    st.markdown("</div></div>", unsafe_allow_html=True)


def metric_card(icon, label, value, delta="", bg="#edf4ff", fg="#4d8df7", red=False):
    st.markdown(f"""
    <div class='metric-card'>
        <div class='metric-top'>
            <div class='metric-icon' style='background:{bg};color:{fg};'>{icon}</div>
            <div>
                <div class='metric-label'>{label}</div>
                <div class='metric-value'>{value}</div>
                <div class='metric-delta {"red" if red else ""}'>{delta}</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def page_title(title: str, subtitle: str):
    st.markdown(f"<div class='section-title'>{title}</div><div class='section-sub'>{subtitle}</div>", unsafe_allow_html=True)


def show_df(df: pd.DataFrame, limit=1000):
    if df is None or df.empty:
        st.info("No data available.")
    else:
        if len(df) > limit:
            st.caption(f"Showing first {limit:,} rows out of {len(df):,}. Use export for full data.")
        st.dataframe(safe_display_df(df, limit), use_container_width=True, hide_index=True)


# =========================================================
# LOGIN
# =========================================================

def login_page():
    st.markdown("""
    <div class='login-bg'>
        <div class='login-card'>
            <div style='text-align:center;font-size:54px;'>🇮🇳</div>
            <div class='login-title'>IMS Recon Pro</div>
            <div class='login-sub'>Secure GST IMS Reconciliation Platform<br>© @BAJRABHANU</div>
    """, unsafe_allow_html=True)

    username = st.text_input("User ID", placeholder="MainAdmin / User1 / User2")
    password = st.text_input("Password", type="password", placeholder="Enter password")

    if st.button("🔐 Login Securely", use_container_width=True):
        user = USER_MASTER.get(username)
        if user and password == user["password"]:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.role = user["role"]
            st.session_state.display_name = user["name"]
            log_event("Login", "Successful login")
            load_user_state()
            st.rerun()
        else:
            st.error("Invalid User ID or Password. Please check case-sensitive credentials.")

    st.markdown("""
            <div style='margin-top:18px;font-size:13px;color:#63758e;text-align:center;line-height:1.6;'>
                MainAdmin / Adminpwd<br>User1 / Userpwd1<br>User2 / Userpwd2
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)


# =========================================================
# PAGES
# =========================================================

def dashboard_page():
    hero_dashboard()

    p, ims, recon = st.session_state.purchase_df, st.session_state.ims_df, st.session_state.recon_df
    total_itc = float(ims["total_tax"].sum()) if not ims.empty else 0
    matched = int((recon["mismatch_type"] == "Matched").sum()) if not recon.empty else 0
    pending = int((recon["recommended_action"] == "Pending").sum()) if not recon.empty else 0
    accepted = int((recon["recommended_action"] == "Accepted").sum()) if not recon.empty else 0
    highrisk = int(recon["risk_level"].isin(["High", "Critical"]).sum()) if not recon.empty else 0

    cols = st.columns(5)
    with cols[0]: metric_card("📚", "Purchase Rows", f"{len(p):,}", "Books data", "#ffefe2", "#ec8b24")
    with cols[1]: metric_card("📥", "IMS Rows", f"{len(ims):,}", st.session_state.get("ims_source",""), "#ecfaef", "#27a857")
    with cols[2]: metric_card("✅", "Matched", f"{matched:,}", "Accepted ready", "#edf4ff", "#4d8df7")
    with cols[3]: metric_card("📌", "Pending", f"{pending:,}", "Needs action", "#f4eefe", "#8b6cf7")
    with cols[4]: metric_card("⚠️", "High Risk", f"{highrisk:,}", "Review required", "#fff0ed", "#e1563a", True)

    c1, c2, c3 = st.columns([2, 2, 1.3])
    with c1:
        st.markdown("<div class='panel'><div class='panel-title'>Reconciliation Summary</div>", unsafe_allow_html=True)
        if recon.empty:
            st.info("Upload data and start reconciliation.")
        else:
            show_df(recon_summary(recon), 20)
        st.markdown("</div>", unsafe_allow_html=True)

    with c2:
        st.markdown("<div class='panel'><div class='panel-title'>Top Mismatch Reasons</div>", unsafe_allow_html=True)
        if recon.empty:
            st.info("No reconciliation data.")
        else:
            summary = recon["mismatch_type"].value_counts().reset_index()
            summary.columns = ["Mismatch Type", "Count"]
            show_df(summary, 20)
        st.markdown("</div>", unsafe_allow_html=True)

    with c3:
        st.markdown("<div class='small-card'><div class='panel-title'>Compliance Health</div>", unsafe_allow_html=True)
        health = 0 if recon.empty else round((matched / max(len(recon), 1)) * 100, 2)
        st.markdown(f"<div style='font-size:48px;font-weight:900;color:#112244;text-align:center;margin:20px 0;'>{health}%</div>", unsafe_allow_html=True)
        st.progress(int(min(100, health)))
        st.caption("Based on matched records against total reconciliation records.")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='small-card'><div class='panel-title'>AI Insight</div><br>", unsafe_allow_html=True)
        if recon.empty:
            st.write("Upload data to generate AI-like insights.")
        else:
            st.write(generate_ai_insight())
        st.markdown("</div>", unsafe_allow_html=True)


def client_setup_page():
    page_title("Client Setup", "Set client GSTIN, return period and review controls.")
    with st.form("client_form"):
        c1, c2, c3 = st.columns(3)
        with c1:
            client_name = st.text_input("Client Name", st.session_state.client_name)
        with c2:
            client_gstin = st.text_input("Client GSTIN", st.session_state.client_gstin).upper()
        with c3:
            return_period = st.text_input("Return Period", st.session_state.return_period)

        submitted = st.form_submit_button("💾 Save Client Setup", use_container_width=True)
        if submitted:
            st.session_state.client_name = client_name
            st.session_state.client_gstin = normalize_gstin(client_gstin)
            st.session_state.return_period = return_period
            save_user_state(["client_name", "client_gstin", "return_period"])
            log_event("Client Setup", "Client details saved")
            st.success("Client setup saved.")

    if st.session_state.client_gstin:
        if validate_gstin(st.session_state.client_gstin):
            st.success("GSTIN format is valid.")
        else:
            st.warning("GSTIN format appears invalid.")


def upload_center_page():
    page_title("Upload Center", "Upload Purchase Register, GST IMS Utility .xlsm, or IMS JSON.")

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown("<div class='panel'><div class='panel-title'>📚 Purchase Register</div>", unsafe_allow_html=True)
        file = st.file_uploader("Upload Purchase Register", type=["xlsx", "xls", "csv"], key="purchase_upload")
        if file and st.button("Process Purchase Register", use_container_width=True):
            try:
                df = read_purchase_file(file)
                st.session_state.purchase_df = df
                save_user_state(["purchase_df"])
                log_event("Upload", f"Purchase Register uploaded: {len(df):,} rows")
                st.success(f"Purchase Register processed: {len(df):,} rows.")
            except Exception as e:
                st.error(f"Unable to process Purchase Register: {e}")
        st.markdown("</div>", unsafe_allow_html=True)

    with c2:
        st.markdown("<div class='panel'><div class='panel-title'>🧾 GST IMS Utility</div>", unsafe_allow_html=True)
        file = st.file_uploader("Upload GST IMS Utility (.xlsm)", type=["xlsm", "xlsx"], key="ims_util_upload")
        st.caption("Expected sheets: B2B, B2BA, B2B-DN, B2B-DNA, B2B-CN, B2B-CNA")
        if file and st.button("Process IMS Utility", use_container_width=True):
            try:
                df = read_ims_utility(file)
                st.session_state.ims_df = df
                st.session_state.ims_source = "IMS Utility"
                save_user_state(["ims_df", "ims_source"])
                log_event("Upload", f"IMS Utility uploaded: {len(df):,} rows")
                st.success(f"IMS Utility processed: {len(df):,} rows.")
            except Exception as e:
                st.error(f"Unable to process IMS Utility: {e}")
        st.markdown("</div>", unsafe_allow_html=True)

    with c3:
        st.markdown("<div class='panel'><div class='panel-title'>🧬 IMS JSON</div>", unsafe_allow_html=True)
        file = st.file_uploader("Upload IMS JSON", type=["json"], key="ims_json_upload")
        if file and st.button("Process IMS JSON", use_container_width=True):
            try:
                df = read_ims_json(file)
                st.session_state.ims_df = df
                st.session_state.ims_source = "IMS JSON"
                save_user_state(["ims_df", "ims_source"])
                log_event("Upload", f"IMS JSON uploaded: {len(df):,} rows")
                st.success(f"IMS JSON processed: {len(df):,} rows.")
            except Exception as e:
                st.error(f"Unable to process IMS JSON: {e}")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")
    page_title("Data Health Check", "Upload status and quality summary.")
    h1, h2, h3, h4 = st.columns(4)
    p = st.session_state.purchase_df
    ims = st.session_state.ims_df
    with h1: metric_card("📚", "Purchase Rows", f"{len(p):,}", "", "#ffefe2", "#ec8b24")
    with h2: metric_card("📥", "IMS Rows", f"{len(ims):,}", st.session_state.ims_source, "#ecfaef", "#27a857")
    with h3:
        invalid = int((~p["gstin_valid"]).sum()) if not p.empty and "gstin_valid" in p else 0
        metric_card("⚠️", "Purchase Invalid GSTIN", f"{invalid:,}", "", "#fff0ed", "#e1563a", True)
    with h4:
        invalid = int((~ims["gstin_valid"]).sum()) if not ims.empty and "gstin_valid" in ims else 0
        metric_card("🛡️", "IMS Invalid GSTIN", f"{invalid:,}", "", "#edf4ff", "#4d8df7")


def ims_data_viewer_page():
    page_title("IMS Data Viewer", "Review uploaded and standardized data before reconciliation.")
    tabs = st.tabs(["Purchase Register", "IMS Combined", "IMS Sheet Summary"])
    with tabs[0]:
        show_df(st.session_state.purchase_df)
    with tabs[1]:
        show_df(st.session_state.ims_df)
    with tabs[2]:
        ims = st.session_state.ims_df
        if ims.empty:
            st.info("No IMS data uploaded.")
        else:
            show_df(ims.groupby(["ims_sheet", "document_type"], dropna=False).agg(
                Records=("document_no", "size"),
                Taxable=("taxable_value", "sum"),
                Tax=("total_tax", "sum"),
            ).reset_index().round(2))


def reconciliation_page():
    page_title("Reconciliation Workspace", "Run IMS reconciliation only when you click Start Reconciliation.")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.session_state.amount_tolerance = st.number_input("Amount Tolerance ₹", min_value=0.0, value=float(st.session_state.amount_tolerance), step=1.0)
    with c2:
        st.session_state.date_tolerance = st.number_input("Date Tolerance Days", min_value=0, value=int(st.session_state.date_tolerance), step=1)
    with c3:
        st.session_state.include_amendments = st.checkbox("Include Amendment Sheets", value=bool(st.session_state.include_amendments))
    with c4:
        st.session_state.use_fuzzy = st.checkbox("Fuzzy Matching Later", value=False, disabled=True)

    ready = not st.session_state.purchase_df.empty and not st.session_state.ims_df.empty
    if not ready:
        st.warning("Upload Purchase Register and IMS data first.")

    if st.button("🚀 Start IMS Reconciliation", type="primary", use_container_width=True, disabled=not ready):
        with st.spinner("Running IMS reconciliation..."):
            recon = calculate_recon(
                st.session_state.purchase_df,
                st.session_state.ims_df,
                st.session_state.amount_tolerance,
                st.session_state.date_tolerance,
                st.session_state.include_amendments,
            )
            st.session_state.recon_df = recon
            st.session_state.action_df = recon.copy()
            save_user_state(["recon_df", "action_df"])
            log_event("Reconciliation", f"Reconciliation completed: {len(recon):,} rows")
        st.success(f"Reconciliation completed: {len(st.session_state.recon_df):,} rows.")

    recon = st.session_state.recon_df
    if not recon.empty:
        st.markdown("---")
        tabs = st.tabs(["All Results", "Matched", "Value Mismatch", "Only in IMS", "Only in Purchase", "High Risk"])
        with tabs[0]: show_df(recon)
        with tabs[1]: show_df(recon[recon["mismatch_type"] == "Matched"])
        with tabs[2]: show_df(recon[recon["mismatch_type"].isin(["Value / Tax Mismatch", "Tax Head Mismatch", "Value and Date Mismatch"])])
        with tabs[3]: show_df(recon[recon["mismatch_type"] == "Only in IMS"])
        with tabs[4]: show_df(recon[recon["mismatch_type"] == "Only in Purchase Register"])
        with tabs[5]: show_df(recon[recon["risk_level"].isin(["High", "Critical"])])


def action_center_page():
    page_title("IMS Action Center", "Review system recommended actions and finalize user action.")
    df = st.session_state.action_df
    if df.empty:
        st.info("Run reconciliation first.")
        return

    c1, c2, c3, c4 = st.columns(4)
    with c1: metric_card("✅", "Accepted", f"{(df['final_user_action']=='Accepted').sum():,}", "", "#ecfaef", "#27a857")
    with c2: metric_card("📌", "Pending", f"{(df['final_user_action']=='Pending').sum():,}", "", "#fff7ed", "#f4a62a")
    with c3: metric_card("🚫", "Rejected", f"{(df['final_user_action']=='Rejected').sum():,}", "", "#fff0ed", "#e1563a", True)
    with c4: metric_card("🕘", "No Action", f"{(df['final_user_action']=='No Action').sum():,}", "", "#edf4ff", "#4d8df7")

    view_cols = [
        "supplier_gstin", "supplier_name", "document_type", "document_no", "document_date",
        "taxable_value_ims", "total_tax_ims", "mismatch_type", "risk_level", "recommended_action",
        "final_user_action", "reason", "user_remarks"
    ]
    exist_cols = [c for c in view_cols if c in df.columns]
    edit_df = df[exist_cols].copy()

    st.caption("You can edit Final User Action and User Remarks below.")
    edited = st.data_editor(
        edit_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "final_user_action": st.column_config.SelectboxColumn("Final User Action", options=ACTION_VALUES),
            "user_remarks": st.column_config.TextColumn("User Remarks"),
        },
    )

    if st.button("💾 Save Final Actions", use_container_width=True):
        updated = df.copy()
        for col in ["final_user_action", "user_remarks"]:
            if col in edited.columns:
                updated.loc[edited.index, col] = edited[col].values
        st.session_state.action_df = updated
        save_user_state(["action_df"])
        log_event("Action Center", "Final user actions updated")
        st.success("Final actions saved.")


def risk_center_page():
    page_title("Risk & Exception Center", "Focused review of high-risk IMS records.")
    df = st.session_state.recon_df
    if df.empty:
        st.info("Run reconciliation first.")
        return
    high = df[df["risk_level"].isin(["High", "Critical"])]
    show_df(high)
    if not high.empty:
        st.download_button(
            "Download High Risk Report",
            data=to_excel_bytes({"High Risk": high}),
            file_name="IMS_High_Risk_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


def vendor_followup_page():
    page_title("Vendor Follow-up", "Vendor-wise pending, mismatch and follow-up list.")
    df = st.session_state.recon_df
    if df.empty:
        st.info("Run reconciliation first.")
        return
    follow = df[df["vendor_followup_required"] == True].copy()
    if follow.empty:
        st.success("No vendor follow-up cases identified.")
        return

    summary = follow.groupby(["supplier_gstin", "supplier_name"], dropna=False).agg(
        Exceptions=("mismatch_type", "size"),
        Taxable_Diff=("taxable_value_diff", "sum"),
        Tax_Diff=("total_tax_diff", "sum"),
        High_Risk=("risk_level", lambda x: int(x.isin(["High", "Critical"]).sum())),
    ).reset_index().round(2)
    show_df(summary)

    selected = st.selectbox("Generate email draft for vendor", summary["supplier_gstin"].astype(str).tolist())
    part = follow[follow["supplier_gstin"].astype(str) == str(selected)]
    if not part.empty:
        vendor_name = str(part["supplier_name"].dropna().iloc[0]) if "supplier_name" in part else ""
        email = f"""Dear {vendor_name or 'Vendor'},

During our GST IMS reconciliation for {st.session_state.get('return_period', '')}, the following documents require clarification:

Total exception cases: {len(part)}
Total tax difference / exposure: ₹{float(part['total_tax_diff'].abs().sum()):,.2f}

Request you to kindly review the invoices/credit notes/debit notes and share clarification or corrective action at the earliest.

Regards,
{st.session_state.get('display_name', '')}
"""
        st.text_area("Vendor Email Draft", email, height=220)


def reports_page():
    page_title("Reports & Export", "Download IMS reconciliation and action reports.")
    p, ims, recon, action = st.session_state.purchase_df, st.session_state.ims_df, st.session_state.recon_df, st.session_state.action_df
    sheets = {
        "Final Action Report": action,
        "Reconciliation": recon,
        "Summary": recon_summary(recon),
        "Purchase Standardized": p,
        "IMS Standardized": ims,
        "Audit Log": load_audit(st.session_state.username),
    }
    st.download_button(
        "📥 Download Complete IMS Workpaper",
        data=to_excel_bytes(sheets),
        file_name=f"IMS_Recon_Pro_Workpaper_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.markdown("### Available Reports")
    report_rows = [
        ["Final IMS Action Report", len(action)],
        ["Reconciliation Report", len(recon)],
        ["Accepted Report", int((action["final_user_action"] == "Accepted").sum()) if not action.empty else 0],
        ["Pending Report", int((action["final_user_action"] == "Pending").sum()) if not action.empty else 0],
        ["Rejected Report", int((action["final_user_action"] == "Rejected").sum()) if not action.empty else 0],
        ["Vendor Follow-up Report", int((recon["vendor_followup_required"] == True).sum()) if not recon.empty else 0],
    ]
    show_df(pd.DataFrame(report_rows, columns=["Report", "Records"]), 20)


def ai_insight_page():
    page_title("AI Insight Desk", "Rule-based smart GST IMS insights without external API.")
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    st.write(generate_ai_insight(long=True))
    st.markdown("</div>", unsafe_allow_html=True)

    q = st.text_input("Ask IMS Insight Desk", placeholder="Example: Which invoices should I accept?")
    if q:
        ql = q.lower()
        df = st.session_state.recon_df
        if df.empty:
            st.info("Run reconciliation first.")
        elif "accept" in ql:
            show_df(df[df["recommended_action"] == "Accepted"])
        elif "pending" in ql:
            show_df(df[df["recommended_action"] == "Pending"])
        elif "reject" in ql:
            show_df(df[df["recommended_action"] == "Rejected"])
        elif "risk" in ql:
            show_df(df[df["risk_level"].isin(["High", "Critical"])])
        elif "vendor" in ql:
            vendor_followup_page()
        else:
            st.write(generate_ai_insight(long=True))


def generate_ai_insight(long: bool = False) -> str:
    df = st.session_state.recon_df
    if df.empty:
        return "No reconciliation has been run yet. Upload Purchase Register and IMS data, then click Start IMS Reconciliation."
    total = len(df)
    matched = int((df["mismatch_type"] == "Matched").sum())
    pending = int((df["recommended_action"] == "Pending").sum())
    high = int(df["risk_level"].isin(["High", "Critical"]).sum())
    only_ims = int((df["mismatch_type"] == "Only in IMS").sum())
    only_purchase = int((df["mismatch_type"] == "Only in Purchase Register").sum())
    health = round((matched / max(total, 1)) * 100, 2)
    base = (
        f"IMS Health Score is {health}%. Out of {total:,} reconciled documents, "
        f"{matched:,} are clean matches, {pending:,} should be kept pending/reviewed, "
        f"and {high:,} are high/critical risk cases."
    )
    if not long:
        return base
    return (
        base
        + f"\n\nKey observations:\n"
        + f"- {only_ims:,} documents are appearing in IMS but not in Purchase Register.\n"
        + f"- {only_purchase:,} documents are booked in Purchase Register but not found in IMS.\n"
        + "- Accept matched invoices first, keep value/tax mismatch cases pending, and send vendor follow-up for books-only cases.\n"
        + "- Review credit notes and amendment sheets separately before final upload through GST utility."
    )


def admin_page():
    page_title("Admin Panel", "Data reset, audit log and user controls.")

    st.markdown("### Audit Log")
    show_df(load_audit("" if st.session_state.role == "Main Admin" else st.session_state.username), 500)

    st.markdown("### Reset Data")
    st.warning("Reset will delete saved data for the current logged-in user.")
    confirm = st.text_input("Type DELETE to confirm reset")
    if st.button("🗑️ Reset My Complete Data", use_container_width=True):
        if confirm == "DELETE":
            username = st.session_state.username
            db_delete_user(username)
            log_event("Reset", "User data reset")
            for key in ["purchase_df", "ims_df", "recon_df", "action_df"]:
                st.session_state[key] = pd.DataFrame()
            st.session_state.ims_source = ""
            st.success("Your saved data has been deleted.")
        else:
            st.error("Please type DELETE exactly.")


# =========================================================
# MAIN
# =========================================================

def main():
    init_db()
    init_state()
    inject_css()

    if not st.session_state.logged_in:
        login_page()
        return

    top_header()
    horizontal_nav()

    page = st.session_state.page
    if page == "Dashboard":
        dashboard_page()
    elif page == "Client Setup":
        client_setup_page()
    elif page == "Upload Center":
        upload_center_page()
    elif page == "IMS Data Viewer":
        ims_data_viewer_page()
    elif page == "Reconciliation Workspace":
        reconciliation_page()
    elif page == "Action Center":
        action_center_page()
    elif page == "Risk Center":
        risk_center_page()
    elif page == "Vendor Follow-up":
        vendor_followup_page()
    elif page == "Reports & Export":
        reports_page()
    elif page == "AI Insight Desk":
        ai_insight_page()
    elif page == "Admin Panel":
        admin_page()

    st.markdown(f"""
    <div class='footer-bar'>
        <div style='display:flex;justify-content:space-around;gap:20px;flex-wrap:wrap;'>
            <div class='foot-item'><div style='font-size:26px;'>🛡️</div><div><div class='foot-main'>Secure</div><div class='foot-sub'>Enterprise-grade control</div></div></div>
            <div class='foot-item'><div style='font-size:26px;'>✅</div><div><div class='foot-main'>Compliant</div><div class='foot-sub'>GSTN workflow aligned</div></div></div>
            <div class='foot-item'><div style='font-size:26px;'>🔄</div><div><div class='foot-main'>Reliable</div><div class='foot-sub'>Offline + export ready</div></div></div>
            <div class='foot-item'><div style='font-size:26px;'>✨</div><div><div class='foot-main'>Smart</div><div class='foot-sub'>AI-like insights</div></div></div>
        </div>
    </div>
    <div style='text-align:center;color:#5d718e;font-size:14px;margin-top:14px;padding-bottom:10px;'>
        © 2026 IMS Recon Pro • {COPYRIGHT_OWNER} • Designed for India • Built for Compliance
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
