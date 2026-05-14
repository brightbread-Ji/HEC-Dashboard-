from __future__ import annotations

import base64
import json
import math
import shutil
from html import escape
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


BASE_DIR = Path(__file__).parent
DEFAULT_EXCEL = BASE_DIR / "HEC AOT by team.xlsx"
LOGO_PATH = BASE_DIR / "ipsos-logo.png.png"
REVENUE_ICON_PATH = BASE_DIR / "assets" / "revenue-icon.png"
GM_ICON_PATH = BASE_DIR / "assets" / "gm-icon.png"
AUTH_PATH = BASE_DIR / "config" / "auth.json"
TARGET_PATH = BASE_DIR / "data" / "team_targets.json"
UPLOAD_DIR = BASE_DIR / "data" / "uploads"
PROJECT_FILTER_VERSION = "active-only-20260513"

MONTH_LABELS = {
    1: "Jan",
    2: "Feb",
    3: "Mar",
    4: "Apr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Aug",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dec",
}
MONTHLY_AOT_COLUMNS = ["CF_AOT"] + [f"{m}_AOT" for m in MONTH_LABELS.values()] + ["CO_AOT"]
PROJECT_CATEGORIES = ["去年的项目", "本年结束的项目", "本年执行中项目"]
PROJECT_COLUMNS = [
    "Record_Detail",
    "Job_Start_Date",
    "Job_End_Date",
    "Job Progress%",
    "Parent_Client_Name",
    "Job_Manager",
    "Revenue_Forecast",
    "GM_Forecast_Percentage",
    "YTD_AOT",
    "YTD AOGM%",
    "YTD_TO",
]


st.set_page_config(
    page_title="HEC Team Financial Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


def inject_css() -> None:
    st.markdown(
        """
        <style>
        :root {
            --ipsos-blue: #1f5aa6;
            --ink: #2d3440;
            --muted: #6b7280;
            --line: #e8eaef;
            --soft: #f7f8fb;
            --accent-orange: #ff8a00;
            --accent-yellow: #f4bd16;
            --accent-teal: #0796a8;
            --accent-rose: #cf1f5a;
            --accent-slate: #404a5c;
            --morandi-blue: #8fb6d8;
            --morandi-blue-bg: #e8f0f7;
            --morandi-green: #7da693;
            --morandi-green-bg: #dfeae4;
        }
        .stApp {
            background: #f3f4f7;
            color: var(--ink);
        }
        [data-testid="stSidebar"] {
            background: #ffffff;
            border-right: 1px solid var(--line);
        }
        .block-container {
            max-width: 1500px;
            margin-top: 2.2rem;
            margin-bottom: 1.2rem;
            padding: 1.35rem 1.18rem 1.35rem;
            background: #ffffff;
            border: 1px solid #e6e7eb;
            border-radius: 14px;
            box-shadow: 0 18px 45px rgba(24, 29, 39, 0.10);
        }
        h1, h2, h3 {
            letter-spacing: 0;
            color: var(--ink);
        }
        h1 {
            font-size: 1.72rem !important;
            line-height: 1.2 !important;
            font-weight: 760 !important;
            margin: 0.15rem 0 0.35rem !important;
        }
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
            gap: 10px;
            margin: 12px 0 12px;
        }
        .kpi-card {
            display: grid;
            grid-template-columns: 82px minmax(0, 1fr);
            align-items: center;
            gap: 14px;
            min-height: 118px;
            padding: 14px 16px;
            background: #ffffff;
            border: 1px solid #e7e8ec;
            border-radius: 4px;
            box-shadow: 0 10px 24px rgba(37, 48, 67, 0.10);
        }
        .kpi-card:hover {
            box-shadow: 0 14px 30px rgba(37, 48, 67, 0.14);
            transform: translateY(-1px);
            transition: box-shadow 160ms ease, transform 160ms ease;
        }
        .kpi-card.gauge-kpi {
            grid-template-columns: 118px minmax(0, 1fr);
        }
        .kpi-icon {
            position: relative;
            width: 76px;
            height: 76px;
            border-radius: 20px 20px 20px 4px;
            background: #eef0f8;
            color: var(--accent-teal);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.6rem;
            font-weight: 800;
            overflow: hidden;
        }
        .kpi-icon > span {
            position: relative;
            z-index: 2;
        }
        .kpi-icon.image-icon {
            background: linear-gradient(135deg, #e5f5f2 0%, #edf2f8 100%);
            color: #092763;
        }
        .kpi-icon.image-icon::before,
        .kpi-icon.image-icon::after {
            display: none;
        }
        .kpi-icon.image-icon img {
            position: relative;
            z-index: 2;
            width: 68px;
            height: 68px;
            object-fit: contain;
            filter: drop-shadow(0 4px 8px rgba(9, 39, 99, 0.18));
        }
        .kpi-icon.gm-image-icon {
            background: linear-gradient(135deg, #def3f2 0%, #eef8fa 100%);
            box-shadow: inset 0 0 0 1px rgba(7, 150, 168, 0.14);
        }
        .kpi-icon.gm-image-icon img {
            width: 72px;
            height: 72px;
            filter: drop-shadow(0 5px 9px rgba(7, 150, 168, 0.20));
        }
        .kpi-icon::before {
            content: "";
            position: absolute;
            width: 46px;
            height: 46px;
            right: -10px;
            bottom: -8px;
            border-radius: 50%;
            background: currentColor;
            opacity: 0.16;
        }
        .kpi-icon::after {
            content: "";
            position: absolute;
            left: 15px;
            bottom: 16px;
            width: 44px;
            height: 8px;
            border-radius: 999px;
            background: currentColor;
            opacity: 0.22;
        }
        .kpi-card:nth-child(2) .kpi-icon {
            color: var(--accent-orange);
        }
        .kpi-card:nth-child(3) .kpi-icon {
            color: var(--accent-yellow);
        }
        .kpi-card:nth-child(4) .kpi-icon {
            color: var(--accent-rose);
        }
        .kpi-card:nth-child(5) .kpi-icon {
            color: var(--accent-teal);
        }
        .kpi-icon.gm-icon {
            color: var(--accent-orange);
            align-items: flex-end;
            gap: 4px;
            padding-bottom: 18px;
            font-size: 0.92rem;
            letter-spacing: 0;
        }
        .kpi-icon.gm-icon::before {
            left: 14px;
            right: auto;
            bottom: 18px;
            width: 44px;
            height: 28px;
            border-radius: 4px 4px 0 0;
            background:
                linear-gradient(90deg, currentColor 0 7px, transparent 7px 12px, currentColor 12px 23px, transparent 23px 28px, currentColor 28px 44px);
            opacity: 0.30;
        }
        .kpi-icon.gm-icon::after {
            left: 18px;
            bottom: 48px;
            width: 36px;
            height: 18px;
            border-left: 4px solid currentColor;
            border-top: 4px solid currentColor;
            border-radius: 3px 0 0 0;
            transform: skewX(-18deg);
            background: transparent;
            opacity: 0.75;
        }
        .kpi-icon.jobs-icon {
            color: var(--accent-rose);
        }
        .kpi-icon.jobs-icon::before {
            left: 19px;
            top: 19px;
            width: 38px;
            height: 42px;
            border-radius: 5px;
            background:
                linear-gradient(currentColor 0 0) 11px 10px / 22px 4px no-repeat,
                linear-gradient(currentColor 0 0) 11px 20px / 22px 4px no-repeat,
                linear-gradient(currentColor 0 0) 11px 30px / 22px 4px no-repeat,
                linear-gradient(currentColor 0 0) 4px 10px / 4px 4px no-repeat,
                linear-gradient(currentColor 0 0) 4px 20px / 4px 4px no-repeat,
                linear-gradient(currentColor 0 0) 4px 30px / 4px 4px no-repeat,
                #eef0f8;
            opacity: 1;
            box-shadow: 8px 8px 0 rgba(207, 31, 90, 0.14);
        }
        .kpi-icon.jobs-icon::after {
            right: 9px;
            left: auto;
            bottom: 8px;
            width: 24px;
            height: 24px;
            border-radius: 50%;
            background: currentColor;
            opacity: 0.22;
        }
        .kpi-icon.month-icon {
            color: var(--morandi-blue);
            font-size: 1.05rem;
            letter-spacing: 0.02em;
            text-transform: uppercase;
        }
        .kpi-icon.month-icon::before {
            left: 16px;
            top: 14px;
            right: auto;
            bottom: auto;
            width: 44px;
            height: 48px;
            border-radius: 8px;
            background:
                linear-gradient(currentColor 0 0) 0 0 / 44px 10px no-repeat,
                linear-gradient(currentColor 0 0) 9px 18px / 6px 6px no-repeat,
                linear-gradient(currentColor 0 0) 20px 18px / 6px 6px no-repeat,
                linear-gradient(currentColor 0 0) 31px 18px / 6px 6px no-repeat,
                linear-gradient(currentColor 0 0) 9px 30px / 6px 6px no-repeat,
                linear-gradient(currentColor 0 0) 20px 30px / 6px 6px no-repeat,
                #eef0f8;
            opacity: 0.35;
            box-shadow: 8px 8px 0 rgba(143, 182, 216, 0.22);
        }
        .kpi-icon.month-icon::after {
            display: none;
        }
        .kpi-label {
            color: #667085;
            font-size: 0.9rem;
            line-height: 1.25;
            font-weight: 560;
            word-break: normal;
            overflow-wrap: normal;
            hyphens: none;
        }
        .kpi-value {
            margin-top: 7px;
            color: #3f4652;
            font-size: clamp(1.55rem, 2.35vw, 2.05rem);
            line-height: 1.05;
            font-weight: 500;
            white-space: nowrap;
        }
        .kpi-note {
            margin-top: 8px;
            color: #7a8699;
            font-size: 0.75rem;
            line-height: 1.25;
            word-break: normal;
            overflow-wrap: normal;
        }
        .gauge-meter {
            position: relative;
            width: 112px;
            height: 74px;
        }
        .gauge-arc {
            position: absolute;
            left: 0;
            bottom: 6px;
            width: 112px;
            height: 56px;
            border-radius: 112px 112px 0 0;
            background:
                conic-gradient(
                    from 270deg at 50% 100%,
                    var(--accent-orange) 0 var(--gauge-angle),
                    #e3e5eb var(--gauge-angle) 180deg,
                    transparent 180deg 360deg
                );
        }
        .gauge-arc::after {
            content: "";
            position: absolute;
            left: 18px;
            bottom: 0;
            width: 76px;
            height: 38px;
            border-radius: 76px 76px 0 0;
            background: #ffffff;
        }
        .gauge-needle {
            position: absolute;
            left: 54px;
            bottom: 8px;
            width: 4px;
            height: 43px;
            border-radius: 999px;
            background: #404a5c;
            transform-origin: 50% 100%;
            transform: rotate(var(--needle-angle));
            box-shadow: 0 1px 2px rgba(24, 29, 39, 0.22);
        }
        .gauge-hub {
            position: absolute;
            left: 49px;
            bottom: 2px;
            width: 14px;
            height: 14px;
            border-radius: 50%;
            background: #404a5c;
            border: 3px solid #ffffff;
            box-shadow: 0 1px 4px rgba(24, 29, 39, 0.18);
        }
        .gauge-scale {
            position: absolute;
            left: 0;
            right: 0;
            bottom: 0;
            display: flex;
            justify-content: space-between;
            color: #8a93a3;
            font-size: 0.64rem;
            line-height: 1;
        }
        .gauge-kpi.is-empty .gauge-needle,
        .gauge-kpi.is-empty .gauge-hub {
            opacity: 0.35;
        }
        @media (max-width: 760px) {
            .kpi-grid {
                grid-template-columns: 1fr;
            }
        }
        .section-title {
            margin: 0 0 0.35rem;
            font-size: 0.96rem;
            font-weight: 760;
            color: #252b36;
        }
        .subtle {
            color: var(--muted);
            font-size: 0.92rem;
        }
        .login-panel {
            max-width: 460px;
            margin: 3rem auto 0;
            padding: 26px;
            background: #ffffff;
            border: 1px solid var(--line);
            border-radius: 8px;
            box-shadow: 0 18px 50px rgba(23, 32, 51, 0.08);
        }
        .alert-row {
            padding: 12px 14px;
            border: 1px solid #ffd9a3;
            background: #fff8ed;
            border-radius: 8px;
            margin-bottom: 8px;
        }
        .status-pill {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 999px;
            background: #f3f5fa;
            color: #4c5565;
            font-size: 0.78rem;
            font-weight: 700;
        }
        .chart-panel {
            border: 1px solid #e7e8ec;
            border-radius: 4px;
            background: #ffffff;
            padding: 10px 12px 4px;
            margin-top: 10px;
            box-shadow: 0 2px 10px rgba(24, 29, 39, 0.025);
        }
        .stAlert {
            border-radius: 4px;
            border-color: #e7e8ec;
        }
        [data-testid="stSegmentedControl"] {
            margin: 2.1rem 0 0.9rem;
            position: relative;
            z-index: 0;
        }
        [data-testid="stSegmentedControl"] [role="radiogroup"] {
            display: inline-flex;
            gap: 6px;
            padding: 5px;
            background: #f1f4f8;
            border: 1px solid #dfe4ec;
            border-radius: 8px;
            box-shadow: 0 8px 18px rgba(37, 48, 67, 0.08);
        }
        [data-testid="stSegmentedControl"] label {
            min-height: 34px;
            border-radius: 6px !important;
            border: 1px solid transparent !important;
            color: #4b5563 !important;
            background: transparent !important;
            font-weight: 650 !important;
        }
        [data-testid="stSegmentedControl"] label[data-baseweb="radio"] > div:first-child {
            display: none !important;
        }
        [data-testid="stSegmentedControl"] button[kind^="segmented_control"] {
            min-height: 34px;
            border-radius: 6px !important;
            border: 1px solid transparent !important;
            color: #4b5563 !important;
            background: transparent !important;
            font-weight: 650 !important;
        }
        [data-testid="stSegmentedControl"] button[kind="segmented_controlActive"] {
            background: #0b2e68 !important;
            border-color: #0b2e68 !important;
            color: #ffffff !important;
            box-shadow: 0 5px 12px rgba(11, 46, 104, 0.22);
        }
        [data-testid="stSegmentedControl"] button[kind="segmented_controlActive"] p {
            color: #ffffff !important;
        }
        [data-testid="stSegmentedControl"] label:has(input:checked),
        [data-testid="stSegmentedControl"] label[aria-checked="true"] {
            background: #0b2e68 !important;
            border-color: #0b2e68 !important;
            color: #ffffff !important;
            box-shadow: 0 5px 12px rgba(11, 46, 104, 0.22);
        }
        [data-testid="stSegmentedControl"] label:has(input:checked) p,
        [data-testid="stSegmentedControl"] label[aria-checked="true"] p {
            color: #ffffff !important;
        }
        div[data-baseweb="tag"] {
            background-color: var(--morandi-blue-bg) !important;
            color: #44697e !important;
            border: 1px solid #c7d9e8 !important;
        }
        [data-baseweb="select"] div[data-baseweb="tag"],
        [data-baseweb="select"] span[data-baseweb="tag"],
        [data-baseweb="select"] [data-baseweb="tag"] > span,
        [data-baseweb="select"] [data-baseweb="tag"] > div {
            background-color: var(--morandi-blue-bg) !important;
            color: #44697e !important;
            border-color: #c7d9e8 !important;
        }
        div[data-baseweb="tag"] svg {
            color: #44697e !important;
            fill: #44697e !important;
        }
        [data-baseweb="select"] [data-baseweb="tag"] svg,
        [data-baseweb="select"] [aria-label="Remove"] svg {
            color: #44697e !important;
            fill: #44697e !important;
        }
        [data-testid="stMain"] div[data-baseweb="select"] > div {
            background-color: #ffffff !important;
            border-color: #4b5563 !important;
            box-shadow: inset 0 0 0 1px #4b5563 !important;
        }
        [data-testid="stMain"] div[data-baseweb="select"] svg {
            color: #4b5563 !important;
            fill: #4b5563 !important;
        }
        [data-testid="stMain"] div[data-baseweb="select"] input {
            color: #1f2937 !important;
        }
        [data-testid="stMain"] div[data-baseweb="input"] > div {
            background-color: #ffffff !important;
            border-color: #4b5563 !important;
            box-shadow: inset 0 0 0 1px #4b5563 !important;
        }
        [data-testid="stMain"] div[data-baseweb="input"] input {
            color: #1f2937 !important;
        }
        [data-testid="stSidebar"] [role="radiogroup"] input[type="radio"] {
            accent-color: #0b2e68 !important;
        }
        [data-testid="stDataFrame"] div[role="progressbar"] > div,
        [data-testid="stDataFrame"] [data-testid="stProgress"] > div > div {
            background-color: var(--morandi-green) !important;
        }
        [data-testid="stDataFrame"] div[role="progressbar"],
        [data-testid="stDataFrame"] [data-testid="stProgress"] > div {
            background-color: var(--morandi-green-bg) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def ensure_runtime_dirs() -> None:
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    TARGET_PATH.parent.mkdir(parents=True, exist_ok=True)


@st.cache_data(show_spinner=False)
def load_auth() -> dict:
    with AUTH_PATH.open("r", encoding="utf-8") as f:
        return json.load(f)


def load_targets() -> dict[str, float]:
    if not TARGET_PATH.exists():
        return {}
    with TARGET_PATH.open("r", encoding="utf-8") as f:
        raw = json.load(f)
    return {str(k): float(v or 0) for k, v in raw.items()}


def save_targets(targets: dict[str, float]) -> None:
    with TARGET_PATH.open("w", encoding="utf-8") as f:
        json.dump(targets, f, ensure_ascii=False, indent=2)


def latest_excel_path() -> Path:
    uploads = sorted(UPLOAD_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if uploads:
        return uploads[0]
    return DEFAULT_EXCEL


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    rename_map = {}
    for col in df.columns:
        clean = str(col).strip()
        if clean == "，":
            rename_map[col] = "Joblist_Month"
        elif clean != col:
            rename_map[col] = clean
    df = df.rename(columns=rename_map)

    if "Joblist_Month" not in df.columns and "Joblist_Date" in df.columns:
        dates = pd.to_datetime(df["Joblist_Date"], errors="coerce")
        df["Joblist_Month"] = dates.dt.month

    numeric_prefixes = ("CF_", "Jan_", "Feb_", "Mar_", "Apr_", "May_", "Jun_", "Jul_", "Aug_", "Sep_", "Oct_", "Nov_", "Dec_", "CO_", "YTD_")
    numeric_columns = [
        "Revenue_Forecast",
        "GM_Forecast",
        "GM_Forecast_Percentage",
        "YTD AOGM%",
        "Job Progress%",
        "Billing Progress",
        "AOT_CY_YTD",
        "AOGM_CY_YTD",
        "Sales_CY_YTD",
        "Joblist_Month",
    ]
    for col in df.columns:
        if col.startswith(numeric_prefixes) or col in numeric_columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    for col in ["Job_Start_Date", "Job_End_Date", "Joblist_Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


@st.cache_data(show_spinner="Loading joblist data...")
def load_joblist(path: str, modified_time: float) -> pd.DataFrame:
    del modified_time
    df = pd.read_excel(path, sheet_name=0)
    return normalize_columns(df)


def get_data() -> tuple[pd.DataFrame, Path]:
    path = latest_excel_path()
    if not path.exists():
        st.error("No Excel data file found. Please place HEC AOT by team.xlsx in this folder.")
        st.stop()
    return load_joblist(str(path), path.stat().st_mtime), path


def validate_columns(df: pd.DataFrame, required: list[str]) -> list[str]:
    return [col for col in required if col not in df.columns]


def authenticate(username: str, password: str) -> dict | None:
    auth = load_auth()
    bp = auth.get("bp", {})
    if username.strip().lower() == str(bp.get("username", "")).lower() and password == bp.get("password"):
        return {"role": "bp", "display_name": bp.get("display_name", "Finance BP"), "team": None}

    for item in auth.get("teams", []):
        aliases = {str(item.get("username", "")), str(item.get("team", ""))}
        if username.strip().lower() in {a.lower() for a in aliases} and password == item.get("password"):
            return {"role": "team", "display_name": item.get("team"), "team": item.get("team")}
    return None


def logout() -> None:
    for key in ["user", "selected_team", "scope_choice", "page_choice", "page_tabs"]:
        st.session_state.pop(key, None)
    st.rerun()


def format_currency(value: float) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "N/A"
    sign = "-" if value < 0 else ""
    value = abs(float(value))
    return f"{sign}¥{value / 1_000_000:.2f}M"


def format_percent(value: float | None) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "N/A"
    return f"{float(value) * 100:.1f}%"


def format_percent_cell(value: float | None) -> str:
    if value is None or pd.isna(value):
        return ""
    return format_percent(float(value))


def image_data_uri(path: Path) -> str:
    if not path.exists():
        return ""
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    suffix = path.suffix.lower().lstrip(".") or "png"
    mime = "jpeg" if suffix in {"jpg", "jpeg"} else suffix
    return f"data:image/{mime};base64,{encoded}"


def safe_divide(numerator: float, denominator: float) -> float | None:
    if denominator in (0, None) or pd.isna(denominator):
        return None
    return float(numerator) / float(denominator)


def sum_col(df: pd.DataFrame, col: str) -> float:
    if col not in df.columns:
        return 0.0
    return float(pd.to_numeric(df[col], errors="coerce").fillna(0).sum())


def current_month(df: pd.DataFrame) -> int | None:
    if "Joblist_Month" in df.columns:
        months = pd.to_numeric(df["Joblist_Month"], errors="coerce")
        months = months[(months >= 1) & (months <= 12)]
        if not months.empty:
            return int(months.mode().iloc[0])
    if "Joblist_Date" in df.columns:
        dates = pd.to_datetime(df["Joblist_Date"], errors="coerce").dropna()
        if not dates.empty:
            return int(dates.max().month)
    return None


def page_header(title: str, subtitle: str, data_path: Path, df: pd.DataFrame) -> None:
    left, right = st.columns([0.72, 0.28], vertical_alignment="center")
    with left:
        st.title(title)
        st.markdown(f"<div class='subtle'>{subtitle}</div>", unsafe_allow_html=True)
    with right:
        month = current_month(df)
        month_text = f"{month}月" if month else "N/A"
        st.markdown(
            f"<div class='subtle' style='text-align:right;'>Data file<br><b>{data_path.name}</b><br>"
            f"<span class='status-pill'>{len(df):,} rows · {month_text}</span></div>",
            unsafe_allow_html=True,
        )


def render_kpi_cards(cards: list[dict[str, str]]) -> None:
    html = [f"<div class='kpi-grid kpi-count-{len(cards)}'>"]
    for card in cards:
        label = escape(str(card.get("label", "")))
        value = escape(str(card.get("value", "")))
        note = escape(str(card.get("note", "")))
        icon = escape(str(card.get("icon", "$")))
        icon_type = str(card.get("icon_type", "")).strip()
        note_html = f"<div class='kpi-note'>{note}</div>" if note else ""
        if card.get("kind") == "gauge":
            ratio = card.get("ratio")
            has_value = ratio is not None and not pd.isna(ratio)
            clamped = min(max(float(ratio or 0), 0), 1)
            gauge_angle = clamped * 180
            needle_angle = -90 + gauge_angle
            empty_class = "" if has_value else " is-empty"
            html.append(
                f"<div class='kpi-card gauge-kpi{empty_class}'>"
                f"<div class='gauge-meter' style='--gauge-angle: {gauge_angle:.1f}deg; --needle-angle: {needle_angle:.1f}deg;'>"
                "<div class='gauge-arc'></div>"
                "<div class='gauge-needle'></div>"
                "<div class='gauge-hub'></div>"
                "<div class='gauge-scale'><span>0</span><span>100%</span></div>"
                "</div>"
                "<div>"
                f"<div class='kpi-label'>{label}</div>"
                f"<div class='kpi-value'>{value}</div>"
                f"{note_html}"
                "</div>"
                "</div>"
            )
        else:
            if icon_type == "gm":
                src = image_data_uri(GM_ICON_PATH)
                if src:
                    icon_html = f"<div class='kpi-icon image-icon gm-image-icon'><img src='{src}' alt='GM percentage icon'></div>"
                else:
                    icon_html = "<div class='kpi-icon gm-icon'><span>GM</span></div>"
            elif icon_type == "jobs":
                icon_html = "<div class='kpi-icon jobs-icon'></div>"
            elif icon_type == "month":
                icon_html = f"<div class='kpi-icon month-icon'><span>{icon}</span></div>"
            elif icon_type == "revenue":
                src = image_data_uri(REVENUE_ICON_PATH)
                if src:
                    icon_html = f"<div class='kpi-icon image-icon'><img src='{src}' alt='Revenue icon'></div>"
                else:
                    icon_html = "<div class='kpi-icon'><span>$</span></div>"
            else:
                icon_html = f"<div class='kpi-icon'><span>{icon}</span></div>"
            html.append(
                "<div class='kpi-card'>"
                f"{icon_html}"
                "<div>"
                f"<div class='kpi-label'>{label}</div>"
                f"<div class='kpi-value'>{value}</div>"
                f"{note_html}"
                "</div>"
                "</div>"
            )
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def login_view() -> None:
    col1, col2, col3 = st.columns([1, 1.05, 1])
    with col2:
        if LOGO_PATH.exists():
            st.image(str(LOGO_PATH), width=140)
        st.markdown("### HEC Team Financial Dashboard")
        st.markdown("<div class='subtle'>Use your team account or Finance BP account to continue.</div>", unsafe_allow_html=True)
        with st.form("login_form"):
            username = st.text_input("Username", placeholder="bp or Team_Name_Official")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Sign in", width="stretch")
        if submitted:
            user = authenticate(username, password)
            if user:
                st.session_state["user"] = user
                st.session_state["selected_team"] = user["team"]
                st.session_state.pop("scope_choice", None)
                st.session_state.pop("page_choice", None)
                st.session_state.pop("page_tabs", None)
                st.rerun()
            st.error("Invalid username or password.")


def sidebar(df: pd.DataFrame) -> str | None:
    user = st.session_state["user"]
    if LOGO_PATH.exists():
        st.sidebar.image(str(LOGO_PATH), width=112)
    st.sidebar.markdown(f"**{user['display_name']}**")
    st.sidebar.caption("Finance BP" if user["role"] == "bp" else "Team access")

    teams = sorted([str(x) for x in df.get("Team_Name_Official", pd.Series(dtype=str)).dropna().unique()])
    selected_team = user["team"]
    if user["role"] == "bp":
        options = ["All Teams"] + teams
        if st.session_state.get("scope_choice") not in options:
            st.session_state["scope_choice"] = "All Teams"
        choice = st.sidebar.selectbox("Team scope", options, key="scope_choice")
        selected_team = None if choice == "All Teams" else choice
    else:
        st.sidebar.info(f"Current team: {selected_team}")

    if user["role"] == "bp":
        st.sidebar.divider()
        st.sidebar.markdown("**Navigation**")
        if st.sidebar.button(
            "设置",
            type="primary" if st.session_state.get("page_choice") == "设置" else "secondary",
            width="stretch",
        ):
            st.session_state["page_choice"] = "设置"

    st.sidebar.divider()
    if st.sidebar.button("Log out", width="stretch"):
        logout()
    return selected_team


def sync_page_from_tabs() -> None:
    selected = st.session_state.get("page_tabs")
    if selected:
        st.session_state["page_choice"] = selected


def page_tabs() -> str:
    pages = ["财务信息预览", "项目信息预览"]
    if st.session_state.get("page_tabs") not in pages:
        st.session_state["page_tabs"] = pages[0]
    if st.session_state.get("page_choice") not in pages + ["设置"]:
        st.session_state["page_choice"] = st.session_state["page_tabs"]
    page = st.segmented_control(
        "Navigation",
        pages,
        key="page_tabs",
        required=True,
        label_visibility="collapsed",
        width="content",
        on_change=sync_page_from_tabs,
    )
    if st.session_state.get("page_choice") != "设置":
        st.session_state["page_choice"] = str(page or pages[0])
    return str(st.session_state.get("page_choice") or pages[0])


def scoped_data(df: pd.DataFrame, selected_team: str | None) -> pd.DataFrame:
    if selected_team and "Team_Name_Official" in df.columns:
        return df[df["Team_Name_Official"].astype(str) == selected_team].copy()
    return df.copy()


def target_for_scope(targets: dict[str, float], df: pd.DataFrame, selected_team: str | None) -> float:
    if selected_team:
        return float(targets.get(selected_team, 0))
    teams = df.get("Team_Name_Official", pd.Series(dtype=str)).dropna().astype(str).unique()
    return float(sum(targets.get(team, 0) for team in teams))


def financial_page(df: pd.DataFrame, data_path: Path, selected_team: str | None) -> None:
    page_header(
        "HEC Team财务信息预览",
        "AOT, AOGM%, target completion and portfolio mix by team.",
        data_path,
        df,
    )

    missing = validate_columns(df, ["YTD_AOT", "YTD_AOGM", "Job category", "Parent_Client_Name", "Product_Name"])
    if missing:
        st.error(f"Missing required columns: {', '.join(missing)}")
        return

    targets = load_targets()
    ytd_aot = sum_col(df, "YTD_AOT")
    ytd_aogm = sum_col(df, "YTD_AOGM")
    target = target_for_scope(targets, df, selected_team)
    month = current_month(df)
    month_label = MONTH_LABELS.get(month or 0)
    month_aot_col = f"{month_label}_AOT" if month_label else ""

    total_jobs = int(df[df["Job category"].isin(["本年执行中项目", "本年结束的项目"])].shape[0])
    target_completion = safe_divide(ytd_aot, target)
    month_signed_aot = sum_col(df, month_aot_col)
    month_aot_label = f"{month_label.upper()} AOT" if month_label else "Monthly AOT"

    render_kpi_cards(
        [
            {"label": "Team YTD AOT", "value": format_currency(ytd_aot), "note": "Current scope", "icon_type": "revenue"},
            {
                "label": "Team YTD AOGM%",
                "value": format_percent(safe_divide(ytd_aogm, ytd_aot)),
                "note": "YTD_AOGM / YTD_AOT",
                "icon_type": "gm",
            },
            {
                "label": "Target completion%",
                "value": format_percent(target_completion),
                "note": "YTD AOT / Target AOT",
                "kind": "gauge",
                "ratio": target_completion,
            },
            {"label": "Total jobs this year", "value": f"{total_jobs:,}", "note": "Active and closed this year", "icon_type": "jobs"},
            {
                "label": month_aot_label,
                "value": format_currency(month_signed_aot),
                "note": "Monthly new signed AOT",
                "icon": month_label.upper() if month_label else "MTH",
                "icon_type": "month",
            },
        ]
    )

    if target <= 0:
        st.info("Team Target AOT is not set for this scope. Finance BP can maintain it in Settings.")

    with st.container(border=True):
        st.markdown("<div class='section-title'>AOT monthly movement</div>", unsafe_allow_html=True)
        monthly_values = pd.DataFrame(
            {
                "Period": ["C/F"] + list(MONTH_LABELS.values()) + ["C/O"],
                "AOT": [sum_col(df, col) for col in MONTHLY_AOT_COLUMNS],
            }
        )
        fig = px.bar(monthly_values, x="Period", y="AOT", text_auto=".2s", color_discrete_sequence=["#ff8a00"])
        fig.update_traces(width=0.50, marker_line_width=0, textposition="outside")
        fig.update_layout(
            height=340,
            margin=dict(l=8, r=8, t=18, b=4),
            yaxis_title="AOT",
            xaxis_title=None,
            plot_bgcolor="white",
            paper_bgcolor="white",
            font=dict(color="#4b5563", size=12),
            bargap=0.36,
        )
        fig.update_yaxes(gridcolor="#edf0f4", zerolinecolor="#dfe3ea")
        fig.update_xaxes(tickfont=dict(color="#5f6675"))
        st.plotly_chart(fig, width="stretch")

    left, right = st.columns([0.58, 0.42])
    with left:
        with st.container(border=True):
            st.markdown("<div class='section-title'>YTD AOT by Parent client · Top 10</div>", unsafe_allow_html=True)
            client = (
                df.assign(Parent_Client_Name=df["Parent_Client_Name"].fillna("Unknown").astype(str))
                .groupby("Parent_Client_Name", as_index=False)["YTD_AOT"]
                .sum()
                .sort_values("YTD_AOT", ascending=False)
                .head(10)
            )
            fig_client = px.bar(
                client.sort_values("YTD_AOT"),
                x="YTD_AOT",
                y="Parent_Client_Name",
                orientation="h",
                color_discrete_sequence=["#0796a8"],
                text_auto=".2s",
            )
            fig_client.update_layout(
                height=420,
                margin=dict(l=8, r=8, t=20, b=8),
                xaxis_title="YTD AOT",
                yaxis_title=None,
                plot_bgcolor="white",
                paper_bgcolor="white",
                font=dict(color="#4b5563", size=12),
            )
            fig_client.update_xaxes(gridcolor="#edf0f4", zerolinecolor="#dfe3ea")
            fig_client.update_yaxes(tickfont=dict(color="#7a8190"))
            st.plotly_chart(fig_client, width="stretch")
    with right:
        with st.container(border=True):
            st.markdown("<div class='section-title'>YTD AOT by Product</div>", unsafe_allow_html=True)
            product = (
                df.assign(Product_Name=df["Product_Name"].fillna("Unknown").astype(str))
                .groupby("Product_Name", as_index=False)["YTD_AOT"]
                .sum()
                .query("YTD_AOT > 0")
                .sort_values("YTD_AOT", ascending=False)
            )
            if product.empty:
                st.info("No positive YTD AOT product mix available.")
            else:
                total_product_aot = product["YTD_AOT"].sum()
                product["Share"] = product["YTD_AOT"] / total_product_aot
                top_products = product[product["Share"] >= 0.05].copy()
                other = product[product["Share"] < 0.05]["YTD_AOT"].sum()
                if other > 0:
                    top_products = pd.concat(
                        [top_products, pd.DataFrame([{"Product_Name": "Others", "YTD_AOT": other}])],
                        ignore_index=True,
                    )
                top_products = top_products.sort_values("YTD_AOT", ascending=False)
                fig_product = px.pie(
                    top_products,
                    names="Product_Name",
                    values="YTD_AOT",
                    hole=0.58,
                    color_discrete_sequence=["#0796a8", "#ff8a00", "#cf1f5a", "#f4bd16", "#d8d8d8"],
                )
                fig_product.update_traces(textposition="outside", textinfo="percent", marker=dict(line=dict(color="#ffffff", width=2)))
                fig_product.update_layout(
                    height=420,
                    margin=dict(l=8, r=8, t=20, b=8),
                    paper_bgcolor="white",
                    font=dict(color="#4b5563", size=12),
                    legend=dict(orientation="v", yanchor="middle", y=0.52, xanchor="left", x=1.02),
                )
                st.plotly_chart(fig_product, width="stretch")


def format_project_table(df: pd.DataFrame) -> pd.DataFrame:
    table = df[[col for col in PROJECT_COLUMNS if col in df.columns]].copy()
    for col in ["Job_Start_Date", "Job_End_Date"]:
        if col in table.columns:
            table[col] = pd.to_datetime(table[col], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    for col in ["GM_Forecast_Percentage", "YTD AOGM%"]:
        if col in table.columns:
            table[col] = pd.to_numeric(table[col], errors="coerce").map(format_percent_cell)
    return table


def alert_message_style(value: str) -> str:
    if value == "需要尽快GTJ":
        return "color: #d87920; font-weight: 800;"
    if value == "需检查是否有首付款未开":
        return "color: #c94f7c; font-weight: 800;"
    return ""


def project_page(df: pd.DataFrame, data_path: Path) -> None:
    page_header(
        "项目信息预览",
        "Turnover confirmation, searchable project list and control checks.",
        data_path,
        df,
    )

    missing = validate_columns(df, ["YTD_TO", "YTD_AOT", "Job category", "Record_Detail", "Parent_Client_Name"])
    if missing:
        st.error(f"Missing required columns: {', '.join(missing)}")
        return

    month = current_month(df)
    month_label = MONTH_LABELS.get(month or 0)
    month_to_col = f"{month_label}_TO" if month_label else ""
    this_month_to = sum_col(df, month_to_col)
    ytd_to = sum_col(df, "YTD_TO")
    ytd_aot = sum_col(df, "YTD_AOT")

    kpi_slot = st.empty()
    st.markdown("<div class='section-title'>Project list</div>", unsafe_allow_html=True)
    base_all = df[df["Job category"].astype(str) != "子项目"].copy()
    default_project_categories = ["本年执行中项目"]
    if st.session_state.get("project_filter_version") != PROJECT_FILTER_VERSION:
        st.session_state["project_category_filter"] = default_project_categories
        st.session_state["project_filter_version"] = PROJECT_FILTER_VERSION
    f1, f2 = st.columns([0.32, 0.68])
    with f1:
        selected_categories = st.multiselect(
            "Job category",
            PROJECT_CATEGORIES,
            key="project_category_filter",
        )
    with f2:
        search = st.text_input("Search project ID, project name or client", placeholder="Type to search...")

    if selected_categories:
        category_filtered = base_all[base_all["Job category"].isin(selected_categories)].copy()
    else:
        category_filtered = base_all.iloc[0:0].copy()

    selected_project_count = int(category_filtered.shape[0])
    base = category_filtered.copy()
    with kpi_slot.container():
        render_kpi_cards(
            [
                {"label": "选中项目数", "value": f"{selected_project_count:,}", "note": "由 Job category 选择决定", "icon_type": "jobs"},
                {"label": "This month TO", "value": format_currency(this_month_to), "note": month_to_col or "No month detected", "icon_type": "revenue"},
                {"label": "YTD TO", "value": format_currency(ytd_to), "note": "Confirmed revenue YTD", "icon_type": "revenue"},
                {"label": "YTD TO 确认%", "value": format_percent(safe_divide(ytd_to, ytd_aot)), "note": "YTD_TO / YTD_AOT", "icon": "%"},
            ]
        )

    if search:
        haystack = (
            base.get("Record_Detail", pd.Series("", index=base.index)).astype(str)
            + " "
            + base.get("Job_ID", pd.Series("", index=base.index)).astype(str)
            + " "
            + base.get("Parent_Client_Name", pd.Series("", index=base.index)).astype(str)
        )
        base = base[haystack.str.contains(search, case=False, na=False)]

    if "Job_End_Date" in base.columns:
        base = base.assign(_sort_job_end_date=pd.to_datetime(base["Job_End_Date"], errors="coerce"))
        base = base.sort_values("_sort_job_end_date", ascending=True, na_position="last").drop(columns=["_sort_job_end_date"])

    st.dataframe(
        format_project_table(base),
        width="stretch",
        hide_index=True,
        column_config={
            "Revenue_Forecast": st.column_config.NumberColumn("Revenue_Forecast", format="¥ %.0f"),
            "YTD_AOT": st.column_config.NumberColumn("YTD_AOT", format="¥ %.0f"),
            "YTD_TO": st.column_config.NumberColumn("YTD_TO", format="¥ %.0f"),
            "GM_Forecast_Percentage": st.column_config.TextColumn("GM Forecast %"),
            "Job Progress%": st.column_config.ProgressColumn(
                "Job Progress%",
                format="percent",
                min_value=0,
                max_value=1,
                color="#7DA693",
            ),
            "YTD AOGM%": st.column_config.TextColumn("YTD AOGM%"),
        },
    )

    st.markdown("<div class='section-title'>项目信控和合规检查</div>", unsafe_allow_html=True)
    alerts = []
    alert_source = df[pd.to_numeric(df["YTD_AOT"], errors="coerce").fillna(0) != 0].copy()
    if {"GTJ_Status", "Job category"}.issubset(alert_source.columns):
        gtj = alert_source[
            (alert_source["GTJ_Status"].astype(str) == "未完成")
            & (alert_source["Job category"].astype(str) == "本年执行中项目")
        ].copy()
        gtj["提示信息"] = "需要尽快GTJ"
        alerts.append(gtj)
    if {"Job Progress%", "Billing Progress"}.issubset(alert_source.columns):
        billing = alert_source[
            (pd.to_numeric(alert_source["Job Progress%"], errors="coerce").fillna(0) > 0.5)
            & (pd.to_numeric(alert_source["Billing Progress"], errors="coerce").fillna(0) == 0)
        ].copy()
        billing["提示信息"] = "需检查是否有首付款未开"
        alerts.append(billing)

    if alerts:
        alert_df = pd.concat(alerts, ignore_index=True)
        alert_df = alert_df.drop_duplicates(subset=["Record_Detail", "提示信息"])
        cols = [
            "提示信息",
            "Record_Detail",
            "Parent_Client_Name",
            "Job_Manager",
            "Job category",
            "GTJ_Status",
            "Job Progress%",
            "Billing Progress",
            "YTD_AOT",
        ]
        alert_view = alert_df[[col for col in cols if col in alert_df.columns]]
        if "提示信息" in alert_view.columns:
            alert_view = alert_view.style.map(alert_message_style, subset=["提示信息"])
        st.dataframe(
            alert_view,
            width="stretch",
            hide_index=True,
            column_config={
                "Job Progress%": st.column_config.ProgressColumn(
                    "Job Progress%",
                    format="percent",
                    min_value=0,
                    max_value=1,
                    color="#7DA693",
                ),
                "Billing Progress": st.column_config.ProgressColumn(
                    "Billing Progress",
                    format="percent",
                    min_value=0,
                    max_value=1,
                    color="#7DA693",
                ),
                "YTD_AOT": st.column_config.NumberColumn("YTD_AOT", format="¥ %.0f"),
            },
        )
    else:
        st.success("No control alerts under the current scope.")


def settings_page(df: pd.DataFrame, data_path: Path) -> None:
    page_header(
        "设置",
        "Finance BP tools for monthly uploads and Team Target AOT maintenance.",
        data_path,
        df,
    )

    st.markdown("<div class='section-title'>Upload monthly joblist</div>", unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload a new HEC AOT by team Excel file", type=["xlsx"])
    if uploaded is not None:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = f"HEC_AOT_by_team_{stamp}.xlsx"
        target = UPLOAD_DIR / safe_name
        with target.open("wb") as f:
            shutil.copyfileobj(uploaded, f)
        st.cache_data.clear()
        st.success(f"Uploaded and activated: {safe_name}")
        st.rerun()

    st.markdown("<div class='section-title'>Team Target AOT</div>", unsafe_allow_html=True)
    teams = sorted([str(x) for x in df.get("Team_Name_Official", pd.Series(dtype=str)).dropna().unique()])
    targets = load_targets()
    for team in teams:
        targets.setdefault(team, 0)

    editor_df = pd.DataFrame(
        [{"Team": team, "Team Target AOT": float(targets.get(team, 0))} for team in sorted(targets)]
    )
    edited = st.data_editor(
        editor_df,
        width="stretch",
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "Team": st.column_config.TextColumn("Team", required=True),
            "Team Target AOT": st.column_config.NumberColumn("Team Target AOT", min_value=0, step=10000, format="¥ %.0f"),
        },
    )
    if st.button("Save Team Targets", type="primary"):
        cleaned = {}
        for _, row in edited.dropna(subset=["Team"]).iterrows():
            team = str(row["Team"]).strip()
            if team:
                cleaned[team] = float(row.get("Team Target AOT") or 0)
        save_targets(cleaned)
        st.success("Team targets saved.")

    with st.expander("Current login configuration"):
        st.caption("Edit config/auth.json to change passwords. Default BP password is bp2026.")
        auth = load_auth()
        team_rows = [
            {"Team": item.get("team"), "Username": item.get("username"), "Password": item.get("password")}
            for item in auth.get("teams", [])
        ]
        st.dataframe(pd.DataFrame(team_rows), width="stretch", hide_index=True)


def main() -> None:
    ensure_runtime_dirs()
    inject_css()
    df, data_path = get_data()

    if "user" not in st.session_state:
        login_view()
        return

    selected_team = sidebar(df)
    page = page_tabs()

    if page == "设置" and st.session_state["user"]["role"] == "bp":
        settings_page(df, data_path)
        return

    view_df = scoped_data(df, selected_team)
    if view_df.empty:
        st.warning("No records available for the selected scope.")
        return

    if page == "财务信息预览":
        financial_page(view_df, data_path, selected_team)
    elif page == "项目信息预览":
        project_page(view_df, data_path)


if __name__ == "__main__":
    main()
