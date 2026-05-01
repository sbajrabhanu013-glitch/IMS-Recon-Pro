import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(
    page_title="IMS Recon Pro",
    page_icon="🇮🇳",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------- STATIC STATE ----------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = True
    st.session_state.user = "Sachin Sharma"
    st.session_state.role = "Super Admin"

# ---------- STYLES ----------
st.markdown("""
<style>
    :root {
        --navy:#071a3d;
        --navy2:#0d2d63;
        --gold:#d89a3f;
        --saffron:#ff9933;
        --green:#138808;
        --light:#f7f9fc;
        --card:#ffffff;
        --border:#e7edf5;
        --text:#112244;
        --muted:#5c6b85;
        --red:#e1563a;
        --orange:#f4a62a;
        --blue:#4d8df7;
        --purple:#8b6cf7;
    }

    .stApp {
        background: linear-gradient(180deg, #fbfcfe 0%, #f2f6fb 100%);
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #06162f 0%, #071a3d 60%, #031126 100%);
        border-right: 1px solid rgba(255,255,255,0.08);
    }

    section[data-testid="stSidebar"] * {
        color: #f4f7fb;
    }

    .sidebar-logo {
        display:flex;
        align-items:center;
        gap:12px;
        padding:10px 4px 18px 4px;
        margin-bottom:10px;
        border-bottom:1px solid rgba(255,255,255,0.08);
    }

    .logo-mark {
        width:42px;height:42px;border-radius:14px;
        background: linear-gradient(135deg,#f3b34d,#f59e0b 45%,#0f6b36);
        display:flex;align-items:center;justify-content:center;
        color:white;font-size:22px;font-weight:700;
        box-shadow: 0 8px 18px rgba(245,158,11,0.25);
    }

    .sidebar-caption {font-size:12px;color:#b8c6dd; margin-top:2px;}

    .menu-item {
        padding:12px 14px;
        margin:8px 0;
        border-radius:16px;
        color:#e9f0ff;
        font-size:16px;
        border:1px solid rgba(255,255,255,0.06);
        background: rgba(255,255,255,0.02);
    }
    .menu-item.active {
        background: linear-gradient(90deg, rgba(255,153,51,0.18), rgba(255,255,255,0.05));
        border:1px solid rgba(255,153,51,0.55);
        box-shadow: inset 0 0 0 1px rgba(255,255,255,0.04);
    }
    .menu-badge {
        float:right;background:#ff7b36;color:white;font-size:11px;
        padding:2px 8px;border-radius:999px;
    }
    .new-badge {
        float:right;background:#33b36b;color:white;font-size:11px;
        padding:2px 8px;border-radius:999px;
    }

    .status-box {
        margin-top:24px;
        padding:16px;border-radius:18px;
        background: linear-gradient(180deg, rgba(255,255,255,0.04), rgba(255,255,255,0.02));
        border: 1px solid rgba(255,255,255,0.08);
    }
    .status-title{font-weight:600;font-size:18px;}
    .status-line{color:#8ef3a2;font-size:14px;margin-top:8px;}
    .status-foot{font-size:12px;color:#bac9df;margin-top:10px;}

    .hero-wrap {
        position: relative;
        overflow:hidden;
        border-radius:24px;
        background: linear-gradient(90deg, #ffffff 0%, #fffef9 48%, #f0f7ef 100%);
        border: 1px solid var(--border);
        padding: 0;
        box-shadow: 0 10px 25px rgba(16,34,68,0.05);
    }

    .hero-top {
        background: linear-gradient(90deg, #081a3d 0%, #09224a 45%, #0d2f63 100%);
        min-height: 116px;
        position: relative;
        overflow:hidden;
        padding: 24px 28px;
        display:flex;
        align-items:center;
        justify-content:space-between;
    }

    .flag-left, .flag-right {
        position:absolute;top:0;height:116px;width:280px;opacity:0.98;
    }
    .flag-left {left:0; background:
        radial-gradient(circle at 28% 48%, #143f89 0 18px, transparent 19px),
        linear-gradient(165deg, transparent 0 18%, #ff9933 18% 33%, #ffffff 33% 48%, #138808 48% 63%, transparent 63% 100%);
        filter: blur(0.2px);
    }
    .flag-right {right:0; background:
        linear-gradient(345deg, transparent 0 18%, #ff9933 18% 33%, #ffffff 33% 48%, #138808 48% 63%, transparent 63% 100%);
        opacity:0.95;
    }
    .chakra-right {
        position:absolute; right:126px; top:18px; font-size:54px; color:rgba(255,255,255,0.88);
    }

    .hero-brand {
        position:relative; z-index:2; display:flex; align-items:center; gap:16px;
    }
    .hero-brand-mark {
        width:56px;height:56px;border-radius:18px;
        background: linear-gradient(135deg,#fff0d3,#ffffff 45%,#fff1d9);
        display:flex;align-items:center;justify-content:center;color:#0d2f63;
        font-size:30px;font-weight:700; box-shadow: 0 10px 24px rgba(0,0,0,0.12);
    }
    .hero-title {font-size:30px;font-weight:800;color:#ffffff;line-height:1.1;}
    .hero-sub {font-size:15px;color:#dfe9ff;margin-top:6px;}

    .hero-meta {
        position:relative; z-index:2; display:flex; align-items:center; gap:28px; color:white;
        font-size:15px;
    }
    .meta-chip {display:flex; align-items:center; gap:10px;}
    .meta-big {font-weight:700;}
    .meta-small {color:#dbe6ff;font-size:13px;}

    .content-pad {padding: 28px; position:relative;}
    .watermark {
        position:absolute; left:50%; top:42%; transform:translate(-50%,-50%);
        font-size:220px; color:rgba(14,41,90,0.04); pointer-events:none;
    }
    .headline {
        font-size:20px; color:#ff8e1a; font-weight:700;
    }
    .main-title {
        font-size:28px; font-weight:800; color:#112244; line-height:1.25; margin-top:8px;
    }
    .subcopy {font-size:16px; color:#52637d; margin-top:10px; line-height:1.5;}
    .cta-row{display:flex; gap:14px; margin-top:18px; flex-wrap:wrap;}
    .cta-dark, .cta-light {
        display:inline-block; padding:12px 22px; border-radius:14px; font-weight:700; text-decoration:none;
        font-size:15px;
    }
    .cta-dark {background:#0b2a5d; color:white; box-shadow:0 10px 18px rgba(11,42,93,.18);}
    .cta-light {background:white; color:#0b2a5d; border:1px solid #d9e3f3;}

    .feature-card {
        background: linear-gradient(180deg, rgba(255,250,241,0.95), rgba(255,250,241,0.8));
        border:1px solid #f0dfc0; border-radius:16px; padding:14px 16px; margin-bottom:12px;
    }
    .feature-card.blue {background:linear-gradient(180deg,#f5f9ff,#eef5ff); border-color:#d6e4ff;}
    .feature-card.green {background:linear-gradient(180deg,#f5fbf3,#eff9ec); border-color:#d8ead0;}
    .feature-title{font-weight:700;color:#23385d;font-size:16px;}
    .feature-desc{font-size:13px;color:#5f6f89;line-height:1.35;margin-top:4px;}
    .shield-center {
        width:130px;height:130px;border-radius:50%; margin:0 auto 18px auto; background: radial-gradient(circle at 30% 30%, #fffef4, #f8f0d2 55%, #ead39f 100%);
        display:flex;align-items:center;justify-content:center; font-size:56px;
        box-shadow: inset 0 0 0 10px rgba(255,255,255,0.65), 0 10px 24px rgba(194,165,97,0.16);
    }

    .metric-card {
        background:#fff; border:1px solid var(--border); border-radius:20px; padding:18px 20px;
        box-shadow:0 8px 18px rgba(16,34,68,.04);
        height:100%;
    }
    .metric-top {display:flex;align-items:center;gap:14px;}
    .metric-icon {
        width:54px;height:54px;border-radius:50%; display:flex;align-items:center;justify-content:center;
        font-size:24px;
    }
    .metric-label{font-size:14px;color:#59708d;}
    .metric-value{font-size:34px;font-weight:800;color:#142748;line-height:1.15;}
    .metric-delta{font-size:13px;color:#12a150;margin-top:6px;}
    .metric-delta.red{color:#e1563a;}

    .panel {
        background:#fff; border:1px solid var(--border); border-radius:22px; padding:22px;
        box-shadow: 0 10px 22px rgba(16,34,68,0.04); height:100%;
    }
    .panel-title{font-size:18px;font-weight:800;color:#17294a;}
    .panel-link{font-size:13px;color:#2b72e3;font-weight:700;}

    .donut {
        width:180px;height:180px;border-radius:50%;
        background: conic-gradient(#29b35e 0 75.72%, #f9a21a 75.72% 89.71%, #ef5137 89.71% 95.88%, #bec8d8 95.88% 100%);
        display:flex;align-items:center;justify-content:center; margin:14px auto;
    }
    .donut::after {
        content:"243\A Total";
        white-space:pre; text-align:center; font-weight:800; color:#13284a; font-size:18px;
        width:110px;height:110px;border-radius:50%; background:white;
        display:flex;align-items:center;justify-content:center; flex-direction:column; line-height:1.2;
        box-shadow: inset 0 0 0 1px #eef3f8;
    }
    .legend-row{display:flex; align-items:center; justify-content:space-between; margin:9px 0; font-size:14px; color:#5b6f88;}
    .legend-left{display:flex; align-items:center; gap:9px;}
    .dot{width:10px;height:10px;border-radius:50%;display:inline-block;}

    .trend-box {
        height:250px; border-radius:18px; background: linear-gradient(180deg,#fbfdfa,#ffffff);
        border:1px solid #edf3fb; padding:18px; position:relative; overflow:hidden;
    }
    .gridline{position:absolute;left:58px;right:20px;border-top:1px dashed #dfe7f3;}
    .g1{top:30px}.g2{top:70px}.g3{top:110px}.g4{top:150px}.g5{top:190px}.g6{top:230px}
    .y-label{position:absolute; left:10px; font-size:12px; color:#7f91ac;}
    .y1{top:20px}.y2{top:60px}.y3{top:100px}.y4{top:140px}.y5{top:180px}.y6{top:220px}
    .trend-svg {position:absolute; left:48px; top:18px; right:12px; bottom:36px; width:calc(100% - 64px); height:200px;}
    .x-axis {position:absolute; left:56px; right:16px; bottom:10px; display:flex; justify-content:space-between; font-size:12px; color:#8193ae;}

    .mini-title {font-size:13px;color:#70819b;font-weight:700;margin-bottom:8px;}

    .list-card {padding:16px; border-radius:18px; background:#fff; border:1px solid #edf1f7;}
    .quick-tile {
        border:1px solid #e8eef6; border-radius:16px; padding:18px 10px; text-align:center;
        background:#fff; box-shadow:0 4px 10px rgba(16,34,68,0.03); height:100%;
    }
    .quick-icon{font-size:30px;margin-bottom:8px;}
    .quick-label{font-size:14px;font-weight:700;color:#2a3e62;}

    .activity-row{display:flex; justify-content:space-between; gap:12px; padding:10px 0; border-bottom:1px solid #edf2f8;}
    .activity-row:last-child{border-bottom:none;}
    .activity-main{display:flex; gap:12px;}
    .act-icon{font-size:20px;}
    .act-title{font-size:15px;font-weight:700;color:#1d3154;}
    .act-sub{font-size:13px;color:#73849c; margin-top:2px;}
    .act-meta{font-size:13px;color:#7b8da6; text-align:right; min-width:95px;}

    .gauge {
        width:220px; height:110px; border-radius:220px 220px 0 0; margin:10px auto 0 auto;
        background: conic-gradient(from 180deg, #f15a24 0 25%, #f7b731 25% 60%, #56c05d 60% 84%, #dfe7f5 84% 100%);
        position:relative; overflow:hidden;
    }
    .gauge::after {
        content:""; position:absolute; left:22px; right:22px; bottom:-88px; height:176px; background:white; border-radius:50%;
        border:1px solid #eef3fa;
    }
    .gauge-value {
        position:relative; margin-top:-56px; text-align:center; font-size:24px; font-weight:800; color:#142748;
    }
    .gauge-sub {text-align:center; font-weight:700; color:#26a147; margin-top:2px;}

    .small-card{
        background:#fff; border:1px solid var(--border); border-radius:20px; padding:22px; box-shadow:0 8px 20px rgba(16,34,68,.04);
        margin-bottom:18px;
    }
    .footer-bar {
        margin-top:16px; border-radius:22px; background:linear-gradient(90deg,#061a3e 0%, #082b61 45%, #061a3e 100%);
        color:white; padding:18px 22px;
    }
    .foot-item{display:flex;align-items:center;gap:10px;justify-content:center;}
    .foot-main{font-weight:700;}
    .foot-sub{font-size:13px;color:#d4e0ff;}

    /* hide default */
    div[data-testid="stToolbar"], header[data-testid="stHeader"] {visibility:hidden; height:0;}
    .block-container {padding-top:1rem; padding-bottom:1rem; max-width: 1500px;}
</style>
""", unsafe_allow_html=True)

# ---------- SIDEBAR ----------
with st.sidebar:
    st.markdown("""
    <div class='sidebar-logo'>
        <div class='logo-mark'>✦</div>
        <div>
            <div style='font-size:28px;font-weight:800;'>IMS Recon Pro</div>
            <div class='sidebar-caption'>GST IMS Suite</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    menu_items = [
        ("🏠", "Dashboard", "active", ""),
        ("🔄", "Reconciliation", "", ""),
        ("✅", "Action Management", "", ""),
        ("📤", "Upload Data", "", ""),
        ("📊", "Reports & Analytics", "", ""),
        ("⚠️", "Exceptions", "", "12"),
        ("📚", "GST Ledgers", "", ""),
        ("🧭", "Compliance Tracker", "", ""),
        ("🕘", "Audit Trail", "", ""),
        ("👥", "Users & Roles", "", ""),
        ("⚙️", "Settings", "", ""),
    ]
    for icon, label, cls, badge in menu_items:
        badge_html = f"<span class='menu-badge'>{badge}</span>" if badge else ""
        st.markdown(f"<div class='menu-item {cls}'>{icon} &nbsp; {label}{badge_html}</div>", unsafe_allow_html=True)

    st.markdown("<div class='menu-item'>🧠 &nbsp; AI Insight Desk <span class='new-badge'>New</span></div>", unsafe_allow_html=True)

    st.markdown("""
    <div class='status-box'>
        <div class='status-title'>System Status</div>
        <div class='status-line'>● All Systems Operational</div>
        <div class='status-foot'>Last updated: 26 May 2025, 09:30 AM</div>
        <div style='opacity:.18; font-size:54px; text-align:right; margin-top:12px;'>🏛️</div>
    </div>
    <div class='status-box' style='display:flex;align-items:center;justify-content:space-between;'>
        <div>
            <div style='font-size:28px;font-weight:800;'>AR</div>
            <div style='font-size:18px;font-weight:700;'>Arjun R.</div>
            <div class='sidebar-caption'>Super Admin</div>
        </div>
        <div style='font-size:26px;'>⋮</div>
    </div>
    """, unsafe_allow_html=True)

# ---------- MAIN HEADER ----------
st.markdown("""
<div class='hero-wrap'>
    <div class='hero-top'>
        <div class='flag-left'></div>
        <div class='flag-right'></div>
        <div class='chakra-right'>🦁</div>
        <div class='hero-brand'>
            <div class='hero-brand-mark'>⬢</div>
            <div>
                <div class='hero-title'>IMS Recon <span style='color:#f6b443'>Pro</span></div>
                <div class='hero-sub'>Intelligent GST IMS Reconciliation Platform</div>
            </div>
        </div>
        <div class='hero-meta'>
            <div class='meta-chip'>
                <div style='font-size:24px;'>🗓️</div>
                <div><div class='meta-big'>26 May 2025</div><div class='meta-small'>Monday</div></div>
            </div>
            <div class='meta-chip'><div style='font-size:24px;'>🔔</div><div class='meta-big'>7</div></div>
            <div class='meta-chip'>
                <div style='font-size:28px;'>👤</div>
                <div><div class='meta-small'>Welcome,</div><div class='meta-big'>Sachin Sharma</div></div>
            </div>
            <div class='meta-big'>© @BAJRABHANU</div>
        </div>
    </div>
    <div class='content-pad'>
        <div class='watermark'>◉</div>
""", unsafe_allow_html=True)

col1, col2, col3 = st.columns([3.8, 1.8, 1.3])
with col1:
    st.markdown("""
    <div class='headline'>☀️ Namaste, Sachin Sharma! 🙏</div>
    <div class='main-title'>Reconcile Today. Stay Compliant.<br>Drive Confidence.</div>
    <div class='subcopy'>AI-powered IMS reconciliation with accuracy,<br>automation & actionable insights.</div>
    <div class='cta-row'>
        <a href='#' class='cta-dark'>Go to Workspace →</a>
        <a href='#' class='cta-light'>☁️ &nbsp; Upload IMS Data</a>
    </div>
    """, unsafe_allow_html=True)
with col2:
    st.markdown("<div class='shield-center'>🛡️</div>", unsafe_allow_html=True)
with col3:
    st.markdown("""
        <div class='feature-card'>
            <div class='feature-title'>🧠 Smart Reconciliation</div>
            <div class='feature-desc'>AI-driven matching with high accuracy</div>
        </div>
        <div class='feature-card blue'>
            <div class='feature-title'>🛡️ Risk Detection</div>
            <div class='feature-desc'>Identify mismatches & compliance risks</div>
        </div>
        <div class='feature-card green'>
            <div class='feature-title'>📈 Actionable Insights</div>
            <div class='feature-desc'>Real-time dashboards for better decisions</div>
        </div>
    """, unsafe_allow_html=True)

st.markdown("</div></div>", unsafe_allow_html=True)

# ---------- KPI ROW ----------
metrics = [
    ("🧩", "Total Clients", "128", "↑ 12% vs last month", "#ffefe2", "#ec8b24", False),
    ("🗂️", "IMS Uploaded (This Month)", "243", "↑ 18% vs last month", "#ecfaef", "#27a857", False),
    ("📋", "Reconciled (This Month)", "198", "↑ 15% vs last month", "#edf4ff", "#4d8df7", False),
    ("👥", "Match Rate", "92.45%", "↑ 3.2% vs last month", "#f4eefe", "#8b6cf7", False),
    ("⚠️", "Pending Actions", "17", "↓ 8% vs last month", "#fff0ed", "#e1563a", True),
]
cols = st.columns(5)
for c, (icon, label, value, delta, bg, fg, red) in zip(cols, metrics):
    with c:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-top'>
                <div class='metric-icon' style='background:{bg};color:{fg};'>{icon}</div>
                <div>
                    <div class='metric-label'>{label}</div>
                    <div class='metric-value'>{value}</div>
                    <div class='metric-delta {'red' if red else ''}'>{delta}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

# ---------- BODY GRID ----------
left, middle, right = st.columns([2.4, 2.6, 1.25])

with left:
    st.markdown("<div class='panel'><div class='panel-title'>Reconciliation Overview <span style='font-weight:600;color:#7f91ac'>(This Month)</span></div>", unsafe_allow_html=True)
    st.markdown("<div class='donut'></div>", unsafe_allow_html=True)
    legend = [
        ("#29b35e", "Matched", "184 (75.72%)"),
        ("#f9a21a", "Mismatched", "34 (13.99%)"),
        ("#ef5137", "Unmatched", "15 (6.17%)"),
        ("#bec8d8", "Not Processed", "10 (4.12%)"),
    ]
    for color, label, val in legend:
        st.markdown(f"<div class='legend-row'><div class='legend-left'><span class='dot' style='background:{color}'></span>{label}</div><div>{val}</div></div>", unsafe_allow_html=True)
    st.markdown("<div style='margin-top:14px;' class='panel-link'>View full reconciliation report →</div></div>", unsafe_allow_html=True)

with middle:
    st.markdown("<div class='panel'><div style='display:flex;justify-content:space-between;align-items:center;'><div class='panel-title'>Match Rate Trend</div><div class='mini-title'>Last 6 Months ▾</div></div>", unsafe_allow_html=True)
    st.markdown("""
    <div class='trend-box'>
        <div class='gridline g1'></div><div class='gridline g2'></div><div class='gridline g3'></div><div class='gridline g4'></div><div class='gridline g5'></div><div class='gridline g6'></div>
        <div class='y-label y1'>100%</div><div class='y-label y2'>80%</div><div class='y-label y3'>60%</div><div class='y-label y4'>40%</div><div class='y-label y5'>20%</div><div class='y-label y6'>0%</div>
        <svg class='trend-svg' viewBox='0 0 700 220' preserveAspectRatio='none'>
            <defs>
              <linearGradient id='fillg' x1='0' y1='0' x2='0' y2='1'>
                <stop offset='0%' stop-color='#86d29b' stop-opacity='0.35'/>
                <stop offset='100%' stop-color='#86d29b' stop-opacity='0.03'/>
              </linearGradient>
            </defs>
            <path d='M0,120 L116,102 L232,94 L348,88 L464,80 L580,68 L580,200 L0,200 Z' fill='url(#fillg)'/>
            <polyline fill='none' stroke='#2bb05a' stroke-width='4' points='0,120 116,102 232,94 348,88 464,80 580,68'/>
            <circle cx='0' cy='120' r='6' fill='#2bb05a'/>
            <circle cx='116' cy='102' r='6' fill='#2bb05a'/>
            <circle cx='232' cy='94' r='6' fill='#2bb05a'/>
            <circle cx='348' cy='88' r='6' fill='#2bb05a'/>
            <circle cx='464' cy='80' r='6' fill='#2bb05a'/>
            <circle cx='580' cy='68' r='6' fill='#2bb05a'/>
            <text x='-10' y='105' fill='#60748e' font-size='12'>86.11%</text>
            <text x='100' y='87' fill='#60748e' font-size='12'>88.27%</text>
            <text x='216' y='79' fill='#60748e' font-size='12'>89.91%</text>
            <text x='332' y='73' fill='#60748e' font-size='12'>90.42%</text>
            <text x='448' y='65' fill='#60748e' font-size='12'>91.37%</text>
            <text x='560' y='53' fill='#60748e' font-size='12'>92.45%</text>
        </svg>
        <div class='x-axis'><span>Dec ’24</span><span>Jan ’25</span><span>Feb ’25</span><span>Mar ’25</span><span>Apr ’25</span><span>May ’25</span></div>
    </div>
    </div>
    """, unsafe_allow_html=True)

with right:
    st.markdown("<div class='small-card'><div class='panel-title'>Compliance Health</div><div class='gauge'></div><div class='gauge-value'>84%</div><div class='gauge-sub'>Good</div><div style='display:flex;justify-content:space-between;color:#8193ad;font-size:12px;margin-top:8px;'><span>0%</span><span>100%</span></div><div style='font-size:13px;color:#6a7b93;margin-top:14px;'>Keep it up! Your compliance health is good.</div><div class='panel-link' style='margin-top:18px;'>View details →</div></div>", unsafe_allow_html=True)
    st.markdown("<div class='small-card'><div style='display:flex;justify-content:space-between;align-items:center;'><div class='panel-title'>AI Insight</div><div class='new-badge'>New</div></div><div style='margin-top:18px;padding:16px;border-radius:16px;background:#f8fbff;border:1px solid #e8eff8;'><div style='font-size:28px;'>✨</div><div style='font-size:15px;color:#465b79;line-height:1.6;margin-top:8px;'>3 clients have recurring vendor mismatches &gt; 5% in the last 3 months. Review recommended.</div></div><div class='panel-link' style='margin-top:18px;'>View AI Insights →</div></div>", unsafe_allow_html=True)
    st.markdown("<div class='small-card'><div class='panel-title'>Upcoming Compliance</div><div style='display:flex;gap:16px;align-items:center;margin-top:18px;'><div style='font-size:34px;color:#3e74d8;'>🗓️</div><div><div style='font-size:15px;color:#667a95;'>GSTR-3B Due Date</div><div style='font-size:30px;font-weight:800;color:#17294a;'>20 Jun 2025</div><div style='font-size:16px;color:#6a7b93;'>(25 Days Left)</div></div></div><div class='panel-link' style='margin-top:18px;'>View all deadlines →</div></div>", unsafe_allow_html=True)

# ---------- SECOND ROW ----------
l2, m2, r2 = st.columns([2.0, 2.0, 1.35])

with l2:
    st.markdown("<div class='panel'><div class='panel-title'>Quick Actions</div>", unsafe_allow_html=True)
    qcols1 = st.columns(3)
    quicks = [
        ("☁️", "Upload IMS File"),
        ("🔗", "Go to Workspace"),
        ("⚠️", "View Mismatches"),
        ("📋", "Action Center"),
        ("👥", "Vendor Follow-up"),
        ("📄", "Generate Report"),
    ]
    for idx, (icon, label) in enumerate(quicks):
        with qcols1[idx % 3]:
            st.markdown(f"<div class='quick-tile'><div class='quick-icon'>{icon}</div><div class='quick-label'>{label}</div></div>", unsafe_allow_html=True)
        if idx % 3 == 2 and idx != len(quicks)-1:
            qcols1 = st.columns(3)
    st.markdown("</div>", unsafe_allow_html=True)

with m2:
    st.markdown("<div class='panel'><div style='display:flex;justify-content:space-between;align-items:center;'><div class='panel-title'>Recent Activity</div><div class='panel-link'>View All</div></div>", unsafe_allow_html=True)
    activities = [
        ("✅", "IMS file \"ABC Limited_IMS_May25.xlsx\" uploaded successfully.", "26 May 2025, 09:15 AM", "By Riya Kapoor"),
        ("🟠", "Reconciliation completed for \"XYZ Pvt Ltd\".", "26 May 2025, 08:45 AM", "By Ankit Verma"),
        ("🔺", "17 new mismatches identified in \"PQR Solutions\".", "26 May 2025, 08:20 AM", "By Neha Singh"),
    ]
    for icon, title, sub, meta in activities:
        st.markdown(f"""
        <div class='activity-row'>
            <div class='activity-main'>
                <div class='act-icon'>{icon}</div>
                <div>
                    <div class='act-title'>{title}</div>
                    <div class='act-sub'>{sub}</div>
                </div>
            </div>
            <div class='act-meta'>{meta}</div>
        </div>
        """, unsafe_allow_html=True)
    st.markdown("<div class='panel-link' style='margin-top:12px;'>View all activity →</div></div>", unsafe_allow_html=True)

with r2:
    st.markdown("<div class='panel'><div style='display:flex;justify-content:space-between;align-items:center;'><div class='panel-title'>Top Mismatch Reasons</div></div>", unsafe_allow_html=True)
    reasons = [
        ("GSTR-2B vs GSTR-3B", "40.2%", "#ef5137", 80),
        ("Invoice Value Difference", "24.7%", "#f9a21a", 52),
        ("Tax Amount Difference", "18.6%", "#4d8df7", 36),
        ("POS Mismatch", "9.8%", "#29b35e", 24),
        ("Others", "6.7%", "#9da8bc", 18),
    ]
    for label, pct, color, width in reasons:
        st.markdown(f"<div style='margin-top:16px;'><div style='display:flex;justify-content:space-between;font-size:14px;color:#495c78;'><span>{label}</span><span>{pct}</span></div><div style='height:10px;border-radius:999px;background:#edf2f8;margin-top:8px;'><div style='width:{width}%;height:100%;background:{color};border-radius:999px;'></div></div></div>", unsafe_allow_html=True)
    st.markdown("<div class='panel-link' style='margin-top:18px;'>View all mismatch reasons →</div></div>", unsafe_allow_html=True)

# ---------- FOOTER ----------
f1, f2, f3, f4 = st.columns(4)
st.markdown("<div class='footer-bar'>", unsafe_allow_html=True)
with f1:
    st.markdown("<div class='foot-item'><div style='font-size:26px;'>🛡️</div><div><div class='foot-main'>Secure</div><div class='foot-sub'>Enterprise-grade security</div></div></div>", unsafe_allow_html=True)
with f2:
    st.markdown("<div class='foot-item'><div style='font-size:26px;'>✅</div><div><div class='foot-main'>Compliant</div><div class='foot-sub'>GSTN Aligned</div></div></div>", unsafe_allow_html=True)
with f3:
    st.markdown("<div class='foot-item'><div style='font-size:26px;'>🔄</div><div><div class='foot-main'>Reliable</div><div class='foot-sub'>99.9% Uptime</div></div></div>", unsafe_allow_html=True)
with f4:
    st.markdown("<div class='foot-item'><div style='font-size:26px;'>✨</div><div><div class='foot-main'>Smart</div><div class='foot-sub'>AI-Powered Insights</div></div></div>", unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<div style='text-align:center;color:#687a95;font-size:14px;margin-top:14px;'>© 2025 IMS Recon Pro. All rights reserved. | Designed for India. Built for Compliance.</div>", unsafe_allow_html=True)
