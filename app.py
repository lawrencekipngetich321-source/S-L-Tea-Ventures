"""
S&L TEA VENTURES — Tea Leaf Business Dashboard
Run locally: streamlit run app.py
"""

from __future__ import annotations

import calendar
from datetime import datetime
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    from docx import Document
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False


# ---------------------------------------------------------------------
# APP CONFIGURATION & BRAND PALETTE
# ---------------------------------------------------------------------

st.set_page_config(
    page_title="S&L Tea Ventures | Business Dashboard",
    page_icon="🍃",
    layout="wide",
    initial_sidebar_state="expanded",
)

BRAND = "S&L TEA VENTURES"
NAVY = "#102A4C"
NAVY_DARK = "#071B34"
BLUE = "#2474C6"
SKY = "#47A9E8"
TEAL = "#19A7A0"
GREEN = "#35A853"
ORANGE = "#F39C35"
INK = "#20334D"
MUTED = "#6B7C93"
BG = "#F3F6FA"
BORDER = "#E3EAF2"
WHITE = "#FFFFFF"

MONTHS = list(calendar.month_name)[1:]
MONTH_INDEX = {month: i + 1 for i, month in enumerate(MONTHS)}
DATA_PATH = Path(__file__).with_name("sl_tea_records.csv")

COLUMNS = [
    "Year",
    "Month",
    "Buying Price (KES/kg)",
    "Selling Price (KES/kg)",
    "Total Kgs",
    "Expenses (KES)",
    "Net Profit Target (KES)",
]

NUMERIC_COLS = [
    "Buying Price (KES/kg)",
    "Selling Price (KES/kg)",
    "Total Kgs",
    "Expenses (KES)",
    "Net Profit Target (KES)",
]

STARTER_DATA = pd.DataFrame(
    [
        {
            "Year": 2026,
            "Month": "September",
            "Buying Price (KES/kg)": 21.00,
            "Selling Price (KES/kg)": 24.00,
            "Total Kgs": 1388.00,
            "Expenses (KES)": 1500.00,
            "Net Profit Target (KES)": 3000.00,
        },
        {
            "Year": 2026,
            "Month": "October",
            "Buying Price (KES/kg)": 20.50,
            "Selling Price (KES/kg)": 24.00,
            "Total Kgs": 2536.60,
            "Expenses (KES)": 0.00,
            "Net Profit Target (KES)": 8000.00,
        },
        {
            "Year": 2026,
            "Month": "November",
            "Buying Price (KES/kg)": 20.50,
            "Selling Price (KES/kg)": 24.00,
            "Total Kgs": 3360.00,
            "Expenses (KES)": 0.00,
            "Net Profit Target (KES)": 11760.00,
        },
    ]
)


# ---------------------------------------------------------------------
# CUSTOM STYLING
# ---------------------------------------------------------------------

st.markdown(
    f"""
    <style>
    @import url('https://googleapis.com');

    html, body, [class*="css"] {{
        font-family: 'DM Sans', sans-serif;
    }}
    .stApp {{
        background: {BG};
        color: {INK};
    }}
    .block-container {{
        max-width: 1650px;
        padding-top: 1.1rem;
        padding-bottom: 2.3rem;
    }}
    [data-testid="stSidebar"] {{
        background: linear-gradient(180deg, {NAVY_DARK} 0%, {NAVY} 58%, #174B78 100%);
        border-right: 1px solid rgba(255,255,255,.08);
    }}
    [data-testid="stSidebar"] * {{
        color: #F4F8FF !important;
    }}
    [data-testid="stSidebar"] [data-testid="stRadio"] label {{
        padding: 7px 10px;
        border-radius: 8px;
        font-weight: 600;
    }}
    [data-testid="stSidebar"] [data-testid="stRadio"] label:hover {{
        background: rgba(255,255,255,.10);
    }}
    [data-testid="stSidebar"] hr {{
        border-color: rgba(255,255,255,.20);
    }}
    .brand-panel {{
        text-align: center;
        padding: 13px 5px 18px 5px;
    }}
    .brand-logo {{
        width: 66px;
        height: 66px;
        border-radius: 50%;
        margin: 0 auto 9px auto;
        display: flex;
        align-items: center;
        justify-content: center;
        background: rgba(255,255,255,.12);
        border: 1px solid rgba(255,255,255,.35);
        font-size: 38px;
    }}
    .brand-name {{
        font-size: 18px;
        font-weight: 800;
        letter-spacing: .6px;
        color: white;
    }}
    .brand-tagline {{
        color: #C6D8ED;
        font-size: 9px;
        letter-spacing: 1.1px;
        margin-top: 4px;
    }}
    .page-heading {{
        background: linear-gradient(110deg, {NAVY_DARK}, {BLUE});
        border-radius: 15px;
        padding: 19px 23px;
        color: white;
        margin-bottom: 16px;
        box-shadow: 0 7px 22px rgba(16,42,76,.13);
    }}
    .page-heading h1 {{
        color: white !important;
        font-size: 25px;
        font-weight: 800;
        margin: 0;
        letter-spacing: .2px;
    }}
    .page-heading p {{
        color: #D9E9FA;
        font-size: 12px;
        margin: 5px 0 0 0;
    }}
    .section-heading {{
        color: {NAVY};
        font-size: 14px;
        font-weight: 800;
        letter-spacing: .35px;
        padding: 10px 13px;
        background: white;
        border: 1px solid {BORDER};
        border-left: 4px solid {BLUE};
        border-radius: 8px;
        margin: 15px 0 10px 0;
        box-shadow: 0 2px 7px rgba(16,42,76,.035);
    }}
    .kpi-card {{
        background: white;
        border: 1px solid {BORDER};
        border-top: 4px solid var(--accent, {BLUE});
        border-radius: 11px;
        padding: 12px 13px 11px 13px;
        min-height: 102px;
        box-shadow: 0 3px 12px rgba(16,42,76,.045);
        margin-bottom: 8px;
    }}
    .kpi-label {{
        color: {MUTED};
        font-size: 10px;
        font-weight: 700;
        letter-spacing: .35px;
        text-transform: uppercase;
        line-height: 1.3;
    }}
    .kpi-value {{
        color: {NAVY};
        font-size: 20px;
        font-weight: 800;
        margin-top: 8px;
        line-height: 1.15;
        overflow-wrap: anywhere;
    }}
    .kpi-note {{
        color: {MUTED};
        font-size: 10px;
        margin-top: 5px;
    }}
    .home-hero {{
        position: relative;
        overflow: hidden;
        background: linear-gradient(115deg, #071B34 0%, #123F70 56%, #2474C6 100%);
        border-radius: 18px;
        padding: 25px 28px 23px 28px;
        margin: 2px 0 16px 0;
        color: #FFFFFF;
        box-shadow: 0 10px 26px rgba(16,42,76,.16);
    }}
    .home-hero:after {{
        content: "🍃";
        position: absolute;
        right: 30px;
        top: -20px;
        font-size: 118px;
        opacity: .12;
        transform: rotate(-18deg);
    }}
    .home-hero-kicker {{
        color: #BBD9F5;
        font-size: 10px;
        font-weight: 800;
        letter-spacing: 2px;
        text-transform: uppercase;
        margin-bottom: 7px;
    }}
    .home-hero-title {{
        color: #FFFFFF !important;
        font-size: 27px;
        line-height: 1.18;
        font-weight: 800;
        margin: 0;
    }}
    .home-hero-subtitle {{
        color: #DCEBFA;
        font-size: 12px;
        margin: 7px 0 0 0;
        max-width: 760px;
    }}
    .home-hero-pills {{
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin-top: 17px;
    }}
    .home-hero-pill {{
        display: inline-block;
        padding: 5px 12px;
        border-radius: 999px;
        background: rgba(255,255,255,0.14);
        border: 1px solid rgba(255,255,255,0.28);
        color: #FFFFFF;
        font-size: 11px;
        font-weight: 700;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------
# DATA SANITIZATION, PERSISTENCE & METRIC CALCULATIONS
# ---------------------------------------------------------------------


def sanitize_records(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    for col in COLUMNS:
        if col not in cleaned.columns:
            cleaned[col] = 0.0 if col in NUMERIC_COLS else ""

    cleaned = cleaned[COLUMNS].copy()
    cleaned["Year"] = pd.to_numeric(cleaned["Year"], errors="coerce").fillna(2026).astype(int)
    cleaned["Month"] = cleaned["Month"].astype(str).str.strip()
    cleaned = cleaned[cleaned["Month"].isin(MONTHS)].copy()

    for col in NUMERIC_COLS:
        cleaned[col] = pd.to_numeric(cleaned[col], errors="coerce").fillna(0.0).clip(lower=0.0)

    cleaned = cleaned.drop_duplicates(subset=["Year", "Month"], keep="last").reset_index(drop=True)
    return cleaned


def load_records() -> pd.DataFrame:
    if DATA_PATH.exists():
        try:
            df = pd.read_csv(DATA_PATH)
            cleaned = sanitize_records(df)
            if not cleaned.empty:
                return cleaned
        except Exception:
            pass
    STARTER_DATA.to_csv(DATA_PATH, index=False)
    return STARTER_DATA.copy()


def save_records(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = sanitize_records(df)
    cleaned.to_csv(DATA_PATH, index=False)
    return cleaned


def enrich_metrics(df: pd.DataFrame) -> pd.DataFrame:
    data = sanitize_records(df)
    data["Month_Num"] = data["Month"].map(MONTH_INDEX).fillna(0).astype(int)
    data = data.sort_values(["Year", "Month_Num"]).reset_index(drop=True)
    data["Period"] = data["Month"].str[:3] + " " + data["Year"].astype(str)

    data["Buying Cost (KES)"] = data["Buying Price (KES/kg)"] * data["Total Kgs"]
    data["Gross Revenue (KES)"] = data["Selling Price (KES/kg)"] * data["Total Kgs"]
    data["Gross Profit (KES)"] = data["Gross Revenue (KES)"] - data["Buying Cost (KES)"]
    data["Total Outflow (KES)"] = data["Buying Cost (KES)"] + data["Expenses (KES)"]
    data["Net Profit (KES)"] = data["Gross Profit (KES)"] - data["Expenses (KES)"]
    data["Target Variance (KES)"] = data["Net Profit (KES)"] - data["Net Profit Target (KES)"]
    data["Spread (KES/kg)"] = data["Selling Price (KES/kg)"] - data["Buying Price (KES/kg)"]
