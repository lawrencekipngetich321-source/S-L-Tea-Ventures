"""
S&L TEA VENTURES — Tea Leaf Business Dashboard
Run locally: streamlit run app.py

Required packages:
    streamlit
    pandas
    numpy
    plotly
    python-docx

The app stores records in sl_tea_records.csv beside this file.
For Streamlit Community Cloud, note that local files may be reset when
the app is redeployed or its container restarts. Use a database or
persistent cloud storage for long-term multi-user production storage.
"""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
import calendar
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False


# ---------------------------------------------------------------------
# APP CONFIGURATION & BRAND
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
RED = "#E05260"
PURPLE = "#8662C7"
PINK = "#D65B9E"
GOLD = "#D6A52B"
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

# The initial values below are taken from the supplied 2026 report.
# September's buying price is derived from buying cost / total kg:
# 29,148 / 1,388 = 21.00 KES/kg.
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
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap');

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
    .status-strip {{
        background: #EAF3FC;
        color: {NAVY};
        border: 1px solid #D6E6F6;
        border-radius: 9px;
        padding: 10px 13px;
        font-size: 12px;
        font-weight: 600;
        margin: 8px 0 12px 0;
    }}
    .insight-card {{
        background: white;
        border: 1px solid {BORDER};
        border-radius: 10px;
        padding: 13px 14px;
        min-height: 90px;
        margin-bottom: 8px;
    }}
    .insight-label {{
        font-size: 10px;
        color: {MUTED};
        font-weight: 700;
        text-transform: uppercase;
    }}
    .insight-value {{
        font-size: 15px;
        color: {NAVY};
        font-weight: 800;
        margin-top: 7px;
    }}
    div[data-testid="stMetric"] {{
        background: white;
        border: 1px solid {BORDER};
        border-left: 4px solid {BLUE};
        border-radius: 10px;
        padding: 12px 13px;
        box-shadow: 0 2px 8px rgba(16,42,76,.035);
    }}
    div[data-testid="stMetricLabel"] {{
        color: {MUTED};
        font-size: 11px;
    }}
    div[data-testid="stMetricValue"] {{
        color: {NAVY};
        font-weight: 800;
    }}
    .stButton > button, .stDownloadButton > button {{
        background: {BLUE};
        color: white;
        border: 0;
        border-radius: 8px;
        font-weight: 700;
        min-height: 40px;
    }}
    .stButton > button:hover, .stDownloadButton > button:hover {{
        background: {NAVY};
        color: white;
        border: 0;
    }}
    div[data-testid="stDataFrame"] {{
        background: white;
        border-radius: 10px;
    }}
    .footer {{
        color: #8090A5;
        text-align: center;
        font-size: 10px;
        padding: 18px 0 3px 0;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------
# DATA FUNCTIONS
# ---------------------------------------------------------------------

def load_records() -> pd.DataFrame:
    """Load saved records; seed the supplied report figures on first run."""
    if not DATA_PATH.exists():
        STARTER_DATA.to_csv(DATA_PATH, index=False)
        return STARTER_DATA.copy()

    try:
        data = pd.read_csv(DATA_PATH)
        for column in COLUMNS:
            if column not in data.columns:
                data[column] = np.nan

        data = data[COLUMNS].copy()
        data["Year"] = pd.to_numeric(data["Year"], errors="coerce")
        for column in COLUMNS[2:]:
            data[column] = pd.to_numeric(data[column], errors="coerce")

        data["Month"] = data["Month"].astype("string").str.strip()
        data = data.dropna(subset=["Year", "Month"])
        data["Year"] = data["Year"].astype(int)
        data = data[data["Month"].isin(MONTHS)]
        data = data.drop_duplicates(["Year", "Month"], keep="last")

        for column in COLUMNS[2:]:
            data[column] = data[column].fillna(0.0)

        return data.sort_values(
            ["Year", "Month"],
            key=lambda s: s.map(MONTH_INDEX) if s.name == "Month" else s,
        ).reset_index(drop=True)
    except Exception as exc:
        st.error(f"Could not read the saved CSV file: {exc}")
        return STARTER_DATA.copy()


def save_records(data: pd.DataFrame) -> None:
    """Persist clean monthly input records."""
    clean = data[COLUMNS].copy()
    clean.to_csv(DATA_PATH, index=False)


def calculate(data: pd.DataFrame) -> pd.DataFrame:
    """Calculate all financial values; net profit is never manually entered."""
    result = data.copy()

    numeric_columns = [
        "Buying Price (KES/kg)",
        "Selling Price (KES/kg)",
        "Total Kgs",
        "Expenses (KES)",
        "Net Profit Target (KES)",
    ]
    for column in numeric_columns:
        result[column] = pd.to_numeric(
            result[column], errors="coerce"
        ).fillna(0.0)

    result["Buying Cost (KES)"] = (
        result["Buying Price (KES/kg)"] * result["Total Kgs"]
    )
    result["Selling Value (KES)"] = (
        result["Selling Price (KES/kg)"] * result["Total Kgs"]
    )
    result["Gross Profit (KES)"] = (
        result["Selling Value (KES)"] - result["Buying Cost (KES)"]
    )
    result["Net Profit (KES)"] = (
        result["Gross Profit (KES)"] - result["Expenses (KES)"]
    )
    result["Net Margin (%)"] = np.where(
        result["Selling Value (KES)"] > 0,
        result["Net Profit (KES)"] / result["Selling Value (KES)"] * 100,
        0.0,
    )
    result["Profit Spread (KES/kg)"] = (
        result["Selling Price (KES/kg)"]
        - result["Buying Price (KES/kg)"]
    )
    result["Target Variance (KES)"] = (
        result["Net Profit (KES)"] - result["Net Profit Target (KES)"]
    )
    result["Target Variance (%)"] = np.where(
        result["Net Profit Target (KES)"] > 0,
        result["Target Variance (KES)"]
        / result["Net Profit Target (KES)"] * 100,
        0.0,
    )
    result["Target Status"] = np.select(
        [
            result["Net Profit (KES)"] > result["Net Profit Target (KES)"],
            np.isclose(
                result["Net Profit (KES)"],
                result["Net Profit Target (KES)"],
            ),
            result["Net Profit (KES)"] < result["Net Profit Target (KES)"],
        ],
        ["Target surpassed", "Target met", "Target not met"],
        default="Not recorded",
    )
    result["Month Number"] = result["Month"].map(MONTH_INDEX)
    return result


def fmt_kes(value: float) -> str:
    return f"KES {value:,.2f}"


def fmt_num(value: float, decimals: int = 2) -> str:
    return f"{value:,.{decimals}f}"


def get_year_summary(calculated: pd.DataFrame, year: int) -> dict:
    """Build annual summary from recorded months only."""
    year_data = calculated[
        (calculated["Year"] == year) & (calculated["Total Kgs"] > 0)
    ].copy()
    year_data = year_data.sort_values("Month Number")

    totals = {
        "months": len(year_data),
        "kgs": float(year_data["Total Kgs"].sum()),
        "buying": float(year_data["Buying Cost (KES)"].sum()),
        "selling": float(year_data["Selling Value (KES)"].sum()),
        "gross": float(year_data["Gross Profit (KES)"].sum()),
        "expenses": float(year_data["Expenses (KES)"].sum()),
        "net": float(year_data["Net Profit (KES)"].sum()),
        "target": float(year_data["Net Profit Target (KES)"].sum()),
    }
    totals["margin"] = (
        totals["net"] / totals["selling"] * 100
        if totals["selling"] else 0.0
    )
    totals["avg_buy"] = (
        totals["buying"] / totals["kgs"] if totals["kgs"] else 0.0
    )
    totals["avg_sell"] = (
        totals["selling"] / totals["kgs"] if totals["kgs"] else 0.0
    )
    totals["avg_kgs"] = (
        totals["kgs"] / totals["months"] if totals["months"] else 0.0
    )
    totals["avg_net"] = (
        totals["net"] / totals["months"] if totals["months"] else 0.0
    )
    totals["target_variance"] = totals["net"] - totals["target"]
    totals["target_variance_pct"] = (
        totals["target_variance"] / totals["target"] * 100
        if totals["target"] else 0.0
    )
    totals["recorded"] = year_data
    return totals


def monthly_report_table(calculated: pd.DataFrame, year: int) -> pd.DataFrame:
    """Create a 12-month report, leaving unrecorded financial values blank."""
    year_data = calculated[calculated["Year"] == year].copy()
    year_data = year_data.set_index("Month")

    rows = []
    for month in MONTHS:
        if month in year_data.index:
            row = year_data.loc[month]
            if isinstance(row, pd.DataFrame):
                row = row.iloc[-1]
            rows.append({
                "Month": month,
                "Total Kgs": row["Total Kgs"],
                "Buying Cost (KES)": row["Buying Cost (KES)"],
                "Selling Value (KES)": row["Selling Value (KES)"],
                "Gross Profit (KES)": row["Gross Profit (KES)"],
                "Expenses (KES)": row["Expenses (KES)"],
                "Net Profit (KES)": row["Net Profit (KES)"],
                "Net Margin (%)": row["Net Margin (%)"],
                "Net Profit Target (KES)": row["Net Profit Target (KES)"],
                "Target Variance (%)": row["Target Variance (%)"],
                "Target Status": row["Target Status"],
            })
        else:
            rows.append({
                "Month": month,
                "Total Kgs": 0.0,
                "Buying Cost (KES)": np.nan,
                "Selling Value (KES)": np.nan,
                "Gross Profit (KES)": np.nan,
                "Expenses (KES)": np.nan,
                "Net Profit (KES)": np.nan,
                "Net Margin (%)": np.nan,
                "Net Profit Target (KES)": np.nan,
                "Target Variance (%)": np.nan,
                "Target Status": "Not recorded",
            })
    return pd.DataFrame(rows)


def chart_layout(fig, height=310):
    fig.update_layout(
        template="plotly_white",
        height=height,
        margin=dict(l=12, r=12, t=48, b=12),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(family="DM Sans, sans-serif", color=INK, size=11),
        title_font=dict(size=14, color=NAVY),
        legend_title_text="",
        hoverlabel=dict(bgcolor=NAVY, font_color="white"),
    )
    fig.update_xaxes(showgrid=False, linecolor=BORDER)
    fig.update_yaxes(gridcolor="#EDF1F6", zerolinecolor=BORDER)
    return fig


# ---------------------------------------------------------------------
# WORD REPORT GENERATION
# ---------------------------------------------------------------------

def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill.replace("#", ""))
    tc_pr.append(shd)


def set_cell_text(cell, text: str, *, bold=False, color="20334D", size=8):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(2)
    run = paragraph.add_run(str(text))
    run.bold = bold
    run.font.name = "Aptos"
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color.replace("#", ""))
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def set_cell_margins(cell, top=80, start=90, bottom=80, end=90):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in [
        ("top", top), ("start", start), ("bottom", bottom), ("end", end)
    ]:
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def add_word_table(document, headers, rows, widths=None, font_size=7):
    table = document.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    for idx, header in enumerate(headers):
        cell = table.rows[0].cells[idx]
        shade_cell(cell, NAVY)
        set_cell_text(cell, header, bold=True, color="FFFFFF", size=font_size)
        set_cell_margins(cell)

    for row_index, row in enumerate(rows):
        cells = table.add_row().cells
        for idx, value in enumerate(row):
            if row_index % 2 == 1:
                shade_cell(cells[idx], "F1F5FA")
            set_cell_text(cells[idx], value, size=font_size)
            set_cell_margins(cells[idx])

    if widths:
        for row in table.rows:
            for idx, width in enumerate(widths):
                row.cells[idx].width = Inches(width)
    document.add_paragraph().paragraph_format.space_after = Pt(1)
    return table


def add_word_heading(document, text, level=1):
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(9)
    paragraph.paragraph_format.space_after = Pt(5)
    run = paragraph.add_run(text)
    run.bold = True
    run.font.name = "Aptos Display"
    run.font.size = Pt(15 if level == 1 else 11)
    run.font.color.rgb = RGBColor.from_string(NAVY.replace("#", ""))
    return paragraph


def create_word_report(calculated: pd.DataFrame, selected_year: int) -> bytes:
    """Generate a printable, styled Word report for one selected year."""
    if not DOCX_AVAILABLE:
        raise RuntimeError("python-docx is not installed.")

    summary = get_year_summary(calculated, selected_year)
    recorded = summary["recorded"]
    report = monthly_report_table(calculated, selected_year)

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)

    styles = doc.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(9)
    styles["Normal"].font.color.rgb = RGBColor.from_string(INK.replace("#", ""))

    # Header and footer
    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hr = header.add_run("S&L TEA VENTURES  |  BUSINESS PERFORMANCE")
    hr.bold = True
    hr.font.size = Pt(8)
    hr.font.color.rgb = RGBColor.from_string(BLUE.replace("#", ""))

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = footer.add_run(
        f"Confidential business report  •  Generated {datetime.now():%d %B %Y}"
    )
    fr.font.size = Pt(8)
    fr.font.color.rgb = RGBColor.from_string(MUTED.replace("#", ""))

    # Cover title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(22)
    r = p.add_run("🍃  S&L TEA VENTURES")
    r.bold = True
    r.font.name = "Aptos Display"
    r.font.size = Pt(24)
    r.font.color.rgb = RGBColor.from_string(NAVY.replace("#", ""))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"MONTHLY & YEARLY PERFORMANCE REPORT — {selected_year}")
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor.from_string(BLUE.replace("#", ""))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(
        f"Prepared on {datetime.now():%d %B %Y}  |  Currency: Kenyan Shillings (KES)"
    )
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor.from_string(MUTED.replace("#", ""))

    add_word_heading(doc, "1. YEAR-TO-DATE EXECUTIVE SUMMARY")

    kpi_rows = [
        ["Months reported", str(summary["months"]),
         "Total KGs supplied", fmt_num(summary["kgs"])],
        ["Total buying cost", fmt_kes(summary["buying"]),
         "Total selling value", fmt_kes(summary["selling"])],
        ["Gross profit", fmt_kes(summary["gross"]),
         "Total expenses", fmt_kes(summary["expenses"])],
        ["NET PROFIT", fmt_kes(summary["net"]),
         "Net profit margin", f'{summary["margin"]:.1f}%'],
        ["Average buying price/kg", fmt_kes(summary["avg_buy"]),
         "Average selling price/kg", fmt_kes(summary["avg_sell"])],
        ["Average KGs/month", fmt_num(summary["avg_kgs"]),
         "Average net profit/month", fmt_kes(summary["avg_net"])],
        ["Profit spread/kg", fmt_kes(summary["avg_sell"] - summary["avg_buy"]),
         "Combined monthly target", fmt_kes(summary["target"])],
    ]
    add_word_table(
        doc,
        ["Performance measure", "Result", "Performance measure", "Result"],
        kpi_rows,
        font_size=8,
    )

    add_word_heading(doc, "2. MONTH-BY-MONTH PERFORMANCE")
    monthly_rows = []
    for _, row in report.iterrows():
        def val(col, suffix=""):
            value = row[col]
            if pd.isna(value):
                return "—"
            return f"{value:,.2f}{suffix}"

        monthly_rows.append([
            row["Month"],
            val("Total Kgs"),
            val("Buying Cost (KES)"),
            val("Selling Value (KES)"),
            val("Gross Profit (KES)"),
            val("Expenses (KES)"),
            val("Net Profit (KES)"),
            val("Net Margin (%)", "%"),
            row["Target Status"],
        ])

    add_word_table(
        doc,
        [
            "Month", "Kgs", "Buying cost", "Selling value",
            "Gross profit", "Expenses", "Net profit", "Margin", "Target status"
        ],
        monthly_rows,
        font_size=6,
    )

    add_word_heading(doc, "3. TARGET PERFORMANCE")
    met = int((recorded["Net Profit (KES)"] == recorded["Net Profit Target (KES)"]).sum())
    surpassed = int((recorded["Net Profit (KES)"] > recorded["Net Profit Target (KES)"]).sum())
    missed = int((recorded["Net Profit (KES)"] < recorded["Net Profit Target (KES)"]).sum())
    variance_pct = summary["target_variance_pct"]
    target_status = (
        f"Overall target surpassed by {variance_pct:.1f}%"
        if variance_pct > 0
        else f"Overall target shortfall of {abs(variance_pct):.1f}%"
        if variance_pct < 0
        else "Overall target met"
    )
    target_rows = [
        ["Months with targets", str(summary["months"])],
        ["Targets met", str(met)],
        ["Targets surpassed", str(surpassed)],
        ["Targets not met", str(missed)],
        ["Combined target", fmt_kes(summary["target"])],
        ["Actual net profit", fmt_kes(summary["net"])],
        ["Variance against target", fmt_kes(summary["target_variance"])],
        ["Overall target position", target_status],
    ]
    add_word_table(doc, ["Target measure", "Result"], target_rows, font_size=8)

    add_word_heading(doc, "4. BUSINESS HIGHLIGHTS")
    if not recorded.empty:
        best = recorded.loc[recorded["Net Profit (KES)"].idxmax()]
        lowest = recorded.loc[recorded["Net Profit (KES)"].idxmin()]
        peak_kg = recorded.loc[recorded["Total Kgs"].idxmax()]
        peak_margin = recorded.loc[recorded["Net Margin (%)"].idxmax()]
        highlight_rows = [
            ["Best month by net profit", f'{best["Month"]} — {fmt_kes(best["Net Profit (KES)"])}'],
            ["Lowest month by net profit", f'{lowest["Month"]} — {fmt_kes(lowest["Net Profit (KES)"])}'],
            ["Highest KGs supplied", f'{peak_kg["Month"]} — {peak_kg["Total Kgs"]:,.2f} kg'],
            ["Highest net margin", f'{peak_margin["Month"]} — {peak_margin["Net Margin (%)"]:.1f}%'],
        ]
    else:
        highlight_rows = [["No recorded trading months", "Enter monthly data to generate highlights."]]
    add_word_table(doc, ["Highlight", "Result"], highlight_rows, font_size=8)

    add_word_heading(doc, "5. YEAR-END OUTLOOK")
    latest_month_no = int(recorded["Month Number"].max()) if not recorded.empty else 0
    remaining = max(0, 12 - latest_month_no)
    projection = summary["net"] + summary["avg_net"] * remaining
    projection_rows = [
        ["Months remaining in year", str(remaining)],
        ["Actual net profit to date", fmt_kes(summary["net"])],
        ["Average net profit per recorded month", fmt_kes(summary["avg_net"])],
        ["Projected year-end net profit", fmt_kes(projection)],
        ["Projection method", "Actual to date + average monthly net profit × months remaining"],
        ["Important note", "Projection is indicative and assumes the recorded-month average continues."],
    ]
    add_word_table(doc, ["Outlook measure", "Result"], projection_rows, font_size=8)

    add_word_heading(doc, "6. MANAGEMENT COMMENTARY")
    if summary["months"]:
        latest_row = recorded.iloc[-1]
        commentary = (
            f"At the end of {latest_row['Month']} {selected_year}, "
            f"{summary['months']} month(s) have been recorded, with "
            f"{summary['kgs']:,.2f} kg supplied. Total selling value is "
            f"{fmt_kes(summary['selling'])}, compared with buying costs of "
            f"{fmt_kes(summary['buying'])}. Gross profit is "
            f"{fmt_kes(summary['gross'])}. After expenses of "
            f"{fmt_kes(summary['expenses'])}, net profit stands at "
            f"{fmt_kes(summary['net'])}, representing a net margin of "
            f"{summary['margin']:.1f}%. The year-end net profit projection "
            f"is {fmt_kes(projection)}. This projection is indicative."
        )
    else:
        commentary = (
            "No monthly trading records have been entered for this year. "
            "Add monthly data in the dashboard to populate this report."
        )
    doc.add_paragraph(commentary)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("S&L TEA VENTURES  •  END OF REPORT")
    r.bold = True
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor.from_string(BLUE.replace("#", ""))

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()


# ---------------------------------------------------------------------
# SESSION STATE & SIDEBAR NAVIGATION
# ---------------------------------------------------------------------

if "records" not in st.session_state:
    st.session_state.records = load_records()

raw_records = st.session_state.records.copy()
calculated = calculate(raw_records)

with st.sidebar:
    st.markdown(
        """
        <div class="brand-panel">
            <div class="brand-logo">🍃</div>
            <div class="brand-name">S&L TEA VENTURES</div>
            <div class="brand-tagline">TEA LEAF BUSINESS MANAGEMENT</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.divider()
    st.markdown("**MAIN MENU**")

    page = st.radio(
        "Dashboard navigation",
        [
            "Home",
            "Data Entry",
            "Monthly Performance",
            "Yearly Performance",
            "Monthly Report",
        ],
        label_visibility="collapsed",
        key="main_navigation",
    )

    st.divider()
    available_years = sorted(
        set(range(2026, 2031)) | set(raw_records["Year"].astype(int).tolist())
    )
    selected_year = st.selectbox(
        "REPORTING YEAR",
        available_years,
        index=available_years.index(2026) if 2026 in available_years else 0,
    )
    st.divider()
    st.caption("◉  LIVE BUSINESS OVERVIEW")
    st.caption(f"Data last loaded: {datetime.now():%d %b %Y, %I:%M %p}")
    st.caption("Currency: KES")


# ---------------------------------------------------------------------
# SELECTED-YEAR DATA & COMMON CALCULATIONS
# ---------------------------------------------------------------------

year_data = calculated[
    (calculated["Year"] == selected_year)
    & (calculated["Total Kgs"] > 0)
].copy()
year_data = year_data.sort_values("Month Number")
summary = get_year_summary(calculated, selected_year)
report_table = monthly_report_table(calculated, selected_year)

months_reported = summary["months"]
months_remaining = (
    max(0, 12 - int(year_data["Month Number"].max()))
    if not year_data.empty else 12
)
projected_year_profit = (
    summary["net"] + summary["avg_net"] * months_remaining
)
projected_variance_pct = (
    (projected_year_profit - summary["target"]) / summary["target"] * 100
    if summary["target"] else 0.0
)

if not year_data.empty:
    latest = year_data.iloc[-1]
    previous = year_data.iloc[-2] if len(year_data) > 1 else None
else:
    latest = None
    previous = None


# ---------------------------------------------------------------------
# REUSABLE PAGE HEADER & KPI CARD
# ---------------------------------------------------------------------

def page_header(title: str, subtitle: str) -> None:
    st.markdown(
        f"""
        <div class="page-heading">
            <h1>🍃 {title}</h1>
            <p>{subtitle}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section(title: str) -> None:
    st.markdown(
        f'<div class="section-heading">{title}</div>',
        unsafe_allow_html=True,
    )


def kpi(label: str, value: str, accent: str = BLUE, note: str = "") -> None:
    st.markdown(
        f"""
        <div class="kpi-card" style="--accent:{accent}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def insight(label: str, value: str) -> None:
    st.markdown(
        f"""
        <div class="insight-card">
            <div class="insight-label">{label}</div>
            <div class="insight-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def report_download_button(key: str = "word_report") -> None:
    if not DOCX_AVAILABLE:
        st.warning(
            "Word export needs python-docx. Add `python-docx` to requirements.txt."
        )
        return
    try:
        word_bytes = create_word_report(calculated, selected_year)
        st.download_button(
            "⬇️ DOWNLOAD PRINT-READY WORD REPORT",
            data=word_bytes,
            file_name=f"SL_Tea_Ventures_Performance_Report_{selected_year}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key=key,
        )
    except Exception as exc:
        st.error(f"Could not create the Word report: {exc}")


# ---------------------------------------------------------------------
# HOME — ONE-PAGE EXECUTIVE DASHBOARD
# ---------------------------------------------------------------------

if page == "Home":
    page_header(
        "BUSINESS PERFORMANCE DASHBOARD",
        f"Executive overview of tea leaf supply, buying, selling, expenses and profit • {selected_year}",
    )

    if months_remaining > 0:
        st.markdown(
            f'<div class="status-strip">● REPORTING YEAR IN PROGRESS — '
            f'{months_remaining} month(s) remaining. The annual report updates '
            f'automatically as monthly records are saved.</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="status-strip">✓ YEAR-END REPORTING PERIOD COMPLETE</div>',
            unsafe_allow_html=True,
        )

    section("YEAR-TO-DATE SNAPSHOT")

    row1 = st.columns(4)
    kpi("TOTAL KGS SUPPLIED", f'{summary["kgs"]:,.2f} kg', BLUE, "Recorded months")
    kpi("TOTAL BUYING COST", fmt_kes(summary["buying"]), PURPLE, "Cost of tea leaf purchases")
    kpi("TOTAL SELLING VALUE", fmt_kes(summary["selling"]), TEAL, "Gross sales value")
    kpi("NET PROFIT", fmt_kes(summary["net"]), GREEN, "After recorded expenses")

    row2 = st.columns(4)
    kpi("GROSS PROFIT", fmt_kes(summary["gross"]), SKY, "Selling value less buying cost")
    kpi("TOTAL EXPENSES", fmt_kes(summary["expenses"]), ORANGE, "Operating / transport / other costs")
    kpi("NET PROFIT MARGIN", f'{summary["margin"]:.1f}%', PINK, "Net profit ÷ selling value")
    kpi("MONTHS REPORTED", str(months_reported), GOLD, "Months with supplied kilograms")

    row3 = st.columns(4)
    kpi("AVG BUYING PRICE / KG", f'KES {summary["avg_buy"]:,.2f}', PURPLE)
    kpi("AVG SELLING PRICE / KG", f'KES {summary["avg_sell"]:,.2f}', TEAL)
    kpi("AVG KGS PER MONTH", f'{summary["avg_kgs"]:,.2f}', BLUE)
    kpi("AVG NET PROFIT / MONTH", fmt_kes(summary["avg_net"]), GREEN)

    # Latest month vs previous month comparison
    section("LATEST MONTH VS PREVIOUS MONTH")
    if latest is not None and previous is not None:
        comp_rows = []
        compare_fields = [
            ("Total KGs", "Total Kgs", "kg"),
            ("Buying price per kg", "Buying Price (KES/kg)", "KES"),
            ("Selling price per kg", "Selling Price (KES/kg)", "KES"),
            ("Gross profit", "Gross Profit (KES)", "KES"),
            ("Net profit", "Net Profit (KES)", "KES"),
        ]
        for label, field, unit in compare_fields:
            current_value = float(latest[field])
            previous_value = float(previous[field])
            change = current_value - previous_value
            change_pct = (
                change / abs(previous_value) * 100
                if previous_value else np.nan
            )
            comp_rows.append({
                "Metric": label,
                str(latest["Month"]): current_value,
                str(previous["Month"]): previous_value,
                "Change": change,
                "Change %": change_pct,
            })
        comp_df = pd.DataFrame(comp_rows)
        st.dataframe(
            comp_df.style.format({
                str(latest["Month"]): "{:,.2f}",
                str(previous["Month"]): "{:,.2f}",
                "Change": "{:+,.2f}",
                "Change %": lambda x: "—" if pd.isna(x) else f"{x:+.1f}%",
            }),
            use_container_width=True,
            hide_index=True,
        )
    elif latest is not None:
        st.info("A second recorded month is needed to show month-on-month changes.")
    else:
        st.info("Enter monthly records to display the comparison.")

    # Target performance
    section("TARGET PERFORMANCE")
    if not year_data.empty:
        target_met = int(
            (year_data["Net Profit (KES)"] == year_data["Net Profit Target (KES)"]).sum()
        )
        target_surpassed = int(
            (year_data["Net Profit (KES)"] > year_data["Net Profit Target (KES)"]).sum()
        )
        target_missed = int(
            (year_data["Net Profit (KES)"] < year_data["Net Profit Target (KES)"]).sum()
        )
    else:
        target_met = target_surpassed = target_missed = 0

    target_cols = st.columns(4)
    with target_cols[0]:
        st.metric("Months with targets", months_reported)
    with target_cols[1]:
        st.metric("Targets met", target_met)
    with target_cols[2]:
        st.metric("Targets surpassed", target_surpassed)
    with target_cols[3]:
        st.metric("Targets not met", target_missed)

    target_a, target_b, target_c = st.columns(3)
    with target_a:
        insight("COMBINED TARGET", fmt_kes(summary["target"]))
    with target_b:
        insight("ACTUAL NET PROFIT", fmt_kes(summary["net"]))
    with target_c:
        insight(
            "OVERALL TARGET POSITION",
            f'{summary["target_variance_pct"]:+.1f}%'
            if summary["target"] else "No targets entered",
        )

    # Highlights
    section("BUSINESS HIGHLIGHTS")
    if not year_data.empty:
        best = year_data.loc[year_data["Net Profit (KES)"].idxmax()]
        lowest = year_data.loc[year_data["Net Profit (KES)"].idxmin()]
        max_kgs = year_data.loc[year_data["Total Kgs"].idxmax()]
        max_margin = year_data.loc[year_data["Net Margin (%)"].idxmax()]
        hi = st.columns(4)
        with hi[0]:
            insight("BEST MONTH — NET PROFIT", f'{best["Month"]} · {fmt_kes(best["Net Profit (KES)"])}')
        with hi[1]:
            insight("LOWEST MONTH — NET PROFIT", f'{lowest["Month"]} · {fmt_kes(lowest["Net Profit (KES)"])}')
        with hi[2]:
            insight("HIGHEST SUPPLY", f'{max_kgs["Month"]} · {max_kgs["Total Kgs"]:,.2f} kg')
        with hi[3]:
            insight("HIGHEST NET MARGIN", f'{max_margin["Month"]} · {max_margin["Net Margin (%)"]:.1f}%')
    else:
        st.info("Highlights will appear after monthly records are entered.")

    # Outlook
    section("YEAR-END OUTLOOK")
    outlook = st.columns(3)
    with outlook[0]:
        insight("PROJECTED YEAR-END NET PROFIT", fmt_kes(projected_year_profit))
    with outlook[1]:
        insight("MONTHS REMAINING", str(months_remaining))
    with outlook[2]:
        insight("PROJECTED VS COMBINED TARGET", f"{projected_variance_pct:+.1f}%" if summary["target"] else "No target entered")
    st.caption(
        "Projection method: actual net profit to date + average net profit per recorded month × months remaining. "
        "This is an indicative estimate, not a guarantee."
    )

    # Charts arranged in dashboard grid
    section("PERFORMANCE ANALYTICS")
    if not year_data.empty:
        left, middle, right = st.columns(3)

        with left:
            fig = px.bar(
                year_data,
                x="Month",
                y="Total Kgs",
                color="Month",
                text="Total Kgs",
                color_discrete_sequence=[BLUE, SKY, TEAL, GREEN, ORANGE, PURPLE],
                title="Tea Leaf Supply by Month",
            )
            fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
            fig.update_layout(showlegend=False, xaxis_title="", yaxis_title="Kilograms")
            st.plotly_chart(chart_layout(fig, 310), use_container_width=True)

        with middle:
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=year_data["Month"], y=year_data["Net Profit (KES)"],
                name="Net profit", marker_color=GREEN,
            ))
            fig.add_trace(go.Scatter(
                x=year_data["Month"], y=year_data["Net Profit Target (KES)"],
                name="Target", mode="lines+markers",
                line=dict(color=ORANGE, width=3, dash="dash"),
            ))
            fig.update_layout(title="Net Profit vs Target", xaxis_title="", yaxis_title="KES")
            st.plotly_chart(chart_layout(fig, 310), use_container_width=True)

        with right:
            pie_values = [
                max(summary["buying"], 0),
                max(summary["expenses"], 0),
                max(summary["net"], 0),
            ]
            if sum(pie_values) > 0:
                fig = go.Figure(go.Pie(
                    labels=["Buying cost", "Expenses", "Net profit"],
                    values=pie_values,
                    hole=0.58,
                    marker=dict(colors=[BLUE, ORANGE, GREEN]),
                    textinfo="label+percent",
                ))
                fig.update_layout(title="Selling Value Allocation")
                st.plotly_chart(chart_layout(fig, 310), use_container_width=True)
            else:
                st.info("No positive values available for the allocation chart.")

        left2, middle2, right2 = st.columns(3)

        with left2:
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=year_data["Month"], y=year_data["Buying Price (KES/kg)"],
                mode="lines+markers", name="Buying price",
                line=dict(color=PURPLE, width=3),
            ))
            fig.add_trace(go.Scatter(
                x=year_data["Month"], y=year_data["Selling Price (KES/kg)"],
                mode="lines+markers", name="Selling price",
                line=dict(color=TEAL, width=3),
            ))
            fig.update_layout(title="Buying vs Selling Price / KG", xaxis_title="", yaxis_title="KES/kg")
            st.plotly_chart(chart_layout(fig, 310), use_container_width=True)

        with middle2:
            fig = px.bar(
                year_data,
                x="Month",
                y=["Buying Cost (KES)", "Selling Value (KES)"],
                barmode="group",
                title="Buying Cost vs Selling Value",
                color_discrete_sequence=[BLUE, TEAL],
            )
            fig.update_layout(xaxis_title="", yaxis_title="KES")
            st.plotly_chart(chart_layout(fig, 310), use_container_width=True)

        with right2:
            fig = px.line(
                year_data,
                x="Month",
                y="Net Margin (%)",
                markers=True,
                title="Net Profit Margin Trend",
                color_discrete_sequence=[PINK],
            )
            fig.update_traces(line_width=3, marker_size=8)
            fig.update_layout(xaxis_title="", yaxis_title="Margin (%)")
            st.plotly_chart(chart_layout(fig, 310), use_container_width=True)
    else:
        st.info("Charts will appear when monthly data is entered.")

    section("MONTH-BY-MONTH SUMMARY")
    st.dataframe(
        report_table,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Total Kgs": st.column_config.NumberColumn(format="%.2f"),
            "Buying Cost (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Selling Value (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Gross Profit (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Expenses (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Net Profit (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Net Margin (%)": st.column_config.NumberColumn(format="%.1f%%"),
            "Net Profit Target (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Target Variance (%)": st.column_config.NumberColumn(format="%+.1f%%"),
        },
    )

    section("DOWNLOAD YOUR PRINTABLE REPORT")
    report_download_button("home_word_report")


# ---------------------------------------------------------------------
# DATA ENTRY
# ---------------------------------------------------------------------

elif page == "Data Entry":
    page_header(
        "MONTHLY DATA ENTRY",
        "Enter buying price, selling price, total kilograms, expenses and target. Profit is calculated automatically.",
    )

    st.markdown(
        '<div class="status-strip">IMPORTANT: Net profit is calculated by the system. '
        'You do not need to enter net profit manually.</div>',
        unsafe_allow_html=True,
    )

    with st.form("monthly_record_form", clear_on_submit=False):
        st.subheader("Monthly trading information")
        col1, col2, col3 = st.columns(3)

        with col1:
            entry_year = st.selectbox(
                "Financial year",
                list(range(2026, 2036)),
                index=0,
            )
            entry_month = st.selectbox("Month", MONTHS)
            buying_price = st.number_input(
                "Buying price per kg (KES)",
                min_value=0.0,
                value=20.50,
                step=0.50,
                format="%.2f",
            )

        with col2:
            selling_price = st.number_input(
                "Selling price per kg (KES)",
                min_value=0.0,
                value=24.00,
                step=0.50,
                format="%.2f",
            )
            total_kgs = st.number_input(
                "Total kilograms supplied",
                min_value=0.0,
                value=0.0,
                step=10.0,
                format="%.2f",
            )

        with col3:
            expenses = st.number_input(
                "Total monthly expenses (KES)",
                min_value=0.0,
                value=0.0,
                step=100.0,
                format="%.2f",
            )
            profit_target = st.number_input(
                "Monthly net profit target (KES)",
                min_value=0.0,
                value=0.0,
                step=500.0,
                format="%.2f",
                help="Optional benchmark used to compare actual calculated net profit.",
            )

        submitted = st.form_submit_button(
            "💾 SAVE MONTHLY RECORD",
            use_container_width=True,
        )

    if submitted:
        if buying_price <= 0 or selling_price <= 0 or total_kgs <= 0:
            st.error("Buying price, selling price and total kilograms must be greater than zero.")
        else:
            new_row = {
                "Year": int(entry_year),
                "Month": entry_month,
                "Buying Price (KES/kg)": float(buying_price),
                "Selling Price (KES/kg)": float(selling_price),
                "Total Kgs": float(total_kgs),
                "Expenses (KES)": float(expenses),
                "Net Profit Target (KES)": float(profit_target),
            }

            current = st.session_state.records.copy()
            duplicate = (
                (current["Year"] == int(entry_year))
                & (current["Month"] == entry_month)
            )
            current = current.loc[~duplicate].copy()
            current = pd.concat([current, pd.DataFrame([new_row])], ignore_index=True)
            current = current[COLUMNS].sort_values(
                ["Year", "Month"],
                key=lambda s: s.map(MONTH_INDEX) if s.name == "Month" else s,
            ).reset_index(drop=True)

            st.session_state.records = current
            save_records(current)

            gross = (selling_price - buying_price) * total_kgs
            net = gross - expenses
            st.success(
                f"{entry_month} {entry_year} saved. "
                f"Calculated gross profit: {fmt_kes(gross)} | "
                f"Calculated net profit: {fmt_kes(net)}"
            )
            st.rerun()

    section("SAVED MONTHLY RECORDS")
    current_calculated = calculate(st.session_state.records)
    if current_calculated.empty:
        st.info("No records saved yet.")
    else:
        visible = current_calculated.sort_values(
            ["Year", "Month Number"], ascending=[False, False]
        )
        st.dataframe(
            visible[
                [
                    "Year", "Month", "Buying Price (KES/kg)",
                    "Selling Price (KES/kg)", "Total Kgs",
                    "Buying Cost (KES)", "Selling Value (KES)",
                    "Expenses (KES)", "Gross Profit (KES)",
                    "Net Profit (KES)", "Net Profit Target (KES)",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )

        with st.expander("Edit or delete a saved monthly record"):
            if not visible.empty:
                record_options = [
                    f'{int(row["Year"])} — {row["Month"]}'
                    for _, row in visible.iterrows()
                ]
                chosen = st.selectbox("Choose record", record_options)
                chosen_year_text, chosen_month = chosen.split(" — ", 1)
                chosen_year = int(chosen_year_text)
                selected_row = visible[
                    (visible["Year"] == chosen_year)
                    & (visible["Month"] == chosen_month)
                ].iloc[0]

                st.write(
                    f"Current net profit: **{fmt_kes(selected_row['Net Profit (KES)'])}**"
                )
                if st.button("Delete selected month", type="secondary"):
                    updated = st.session_state.records[
                        ~(
                            (st.session_state.records["Year"] == chosen_year)
                            & (st.session_state.records["Month"] == chosen_month)
                        )
                    ].copy()
                    st.session_state.records = updated.reset_index(drop=True)
                    save_records(st.session_state.records)
                    st.success(f"{chosen_month} {chosen_year} deleted.")
                    st.rerun()

    st.caption(
        "Saving a month that already exists for the same year replaces that month's previous record."
    )


# ---------------------------------------------------------------------
# MONTHLY PERFORMANCE
# ---------------------------------------------------------------------

elif page == "Monthly Performance":
    page_header(
        "MONTHLY PERFORMANCE",
        f"Explore month-by-month supply, sales value, costs, expenses and calculated profit • {selected_year}",
    )

    if year_data.empty:
        st.info("No records found for the selected year. Add data from the Data Entry tab.")
    else:
        chosen_month = st.selectbox(
            "Select month",
            year_data["Month"].tolist(),
            index=len(year_data) - 1,
        )
        month_row = year_data[year_data["Month"] == chosen_month].iloc[0]

        section(f"{chosen_month.upper()} PERFORMANCE SNAPSHOT")
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.metric("Total KGs", f'{month_row["Total Kgs"]:,.2f} kg')
        with m2:
            st.metric("Buying cost", fmt_kes(month_row["Buying Cost (KES)"]))
        with m3:
            st.metric("Selling value", fmt_kes(month_row["Selling Value (KES)"]))
        with m4:
            st.metric("Net profit", fmt_kes(month_row["Net Profit (KES)"]))

        m5, m6, m7, m8 = st.columns(4)
        with m5:
            st.metric("Gross profit", fmt_kes(month_row["Gross Profit (KES)"]))
        with m6:
            st.metric("Expenses", fmt_kes(month_row["Expenses (KES)"]))
        with m7:
            st.metric("Net margin", f'{month_row["Net Margin (%)"]:.1f}%')
        with m8:
            st.metric("Profit spread/kg", f'KES {month_row["Profit Spread (KES/kg)"]:.2f}')

        section("MONTHLY FINANCIAL BREAKDOWN")
        breakdown = pd.DataFrame({
            "Financial item": ["Buying cost", "Expenses", "Net profit"],
            "Amount (KES)": [
                month_row["Buying Cost (KES)"],
                month_row["Expenses (KES)"],
                month_row["Net Profit (KES)"],
            ],
        })
        fig = px.bar(
            breakdown,
            x="Financial item",
            y="Amount (KES)",
            color="Financial item",
            text="Amount (KES)",
            color_discrete_sequence=[BLUE, ORANGE, GREEN],
            title=f"{chosen_month} — Cost and Profit Breakdown",
        )
        fig.update_traces(texttemplate="KES %{text:,.0f}", textposition="outside")
        fig.update_layout(showlegend=False, xaxis_title="")
        st.plotly_chart(chart_layout(fig, 330), use_container_width=True)

        section("MONTHLY TARGET ASSESSMENT")
        target = float(month_row["Net Profit Target (KES)"])
        actual = float(month_row["Net Profit (KES)"])
        if target > 0:
            variance = actual - target
            variance_pct = variance / target * 100
            t1, t2, t3 = st.columns(3)
            with t1:
                st.metric("Net profit target", fmt_kes(target))
            with t2:
                st.metric("Actual net profit", fmt_kes(actual))
            with t3:
                st.metric("Variance", fmt_kes(variance), f"{variance_pct:+.1f}%")
            st.progress(min(max(actual / target, 0.0), 1.0))
            if actual > target:
                st.success(f"Target surpassed by {variance_pct:.1f}%.")
            elif np.isclose(actual, target):
                st.success("Target met.")
            else:
                st.warning(f"Target not met. Shortfall: {abs(variance_pct):.1f}%.")
        else:
            st.info("No target was entered for this month.")

        section("MONTHLY RECORD DETAILS")
        st.dataframe(
            month_row.drop(labels=["Month Number"]).to_frame("Value"),
            use_container_width=True,
        )


# ---------------------------------------------------------------------
# YEARLY PERFORMANCE
# ---------------------------------------------------------------------

elif page == "Yearly Performance":
    page_header(
        "YEARLY PERFORMANCE",
        "Annual financial overview, trends, target performance and year-end outlook.",
    )

    all_data = calculate(st.session_state.records)
    year_summary_rows = []
    for year in sorted(all_data["Year"].dropna().astype(int).unique()):
        ys = get_year_summary(all_data, int(year))
        year_summary_rows.append({
            "Year": int(year),
            "Months reported": ys["months"],
            "Total KGs": ys["kgs"],
            "Buying cost (KES)": ys["buying"],
            "Selling value (KES)": ys["selling"],
            "Expenses (KES)": ys["expenses"],
            "Gross profit (KES)": ys["gross"],
            "Net profit (KES)": ys["net"],
            "Net margin (%)": ys["margin"],
        })

    section("SELECTED YEAR — KEY FINANCIAL INDICATORS")
    y1, y2, y3, y4 = st.columns(4)
    with y1:
        st.metric("Total KGs", f'{summary["kgs"]:,.2f}')
    with y2:
        st.metric("Selling value", fmt_kes(summary["selling"]))
    with y3:
        st.metric("Gross profit", fmt_kes(summary["gross"]))
    with y4:
        st.metric("Net profit", fmt_kes(summary["net"]))

    section("ANNUAL TREND CHARTS")
    if not year_data.empty:
        a, b = st.columns(2)
        with a:
            fig = px.bar(
                year_data,
                x="Month",
                y=["Buying Cost (KES)", "Selling Value (KES)"],
                barmode="group",
                color_discrete_sequence=[BLUE, TEAL],
                title=f"{selected_year} — Buying Cost vs Selling Value",
            )
            fig.update_layout(xaxis_title="", yaxis_title="KES")
            st.plotly_chart(chart_layout(fig, 360), use_container_width=True)

        with b:
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=year_data["Month"],
                y=year_data["Net Profit (KES)"],
                name="Net profit",
                marker_color=GREEN,
            ))
            fig.add_trace(go.Scatter(
                x=year_data["Month"],
                y=year_data["Net Profit Target (KES)"],
                name="Target",
                mode="lines+markers",
                line=dict(color=ORANGE, dash="dash", width=3),
            ))
            fig.update_layout(
                title=f"{selected_year} — Net Profit vs Target",
                xaxis_title="",
                yaxis_title="KES",
            )
            st.plotly_chart(chart_layout(fig, 360), use_container_width=True)

        c, d = st.columns(2)
        with c:
            fig = px.area(
                year_data,
                x="Month",
                y="Total Kgs",
                markers=True,
                title=f"{selected_year} — Monthly Supply Trend",
                color_discrete_sequence=[SKY],
            )
            st.plotly_chart(chart_layout(fig, 330), use_container_width=True)
        with d:
            fig = px.line(
                year_data,
                x="Month",
                y="Net Margin (%)",
                markers=True,
                title=f"{selected_year} — Net Margin Trend",
                color_discrete_sequence=[PINK],
            )
            fig.update_traces(line_width=3, marker_size=8)
            st.plotly_chart(chart_layout(fig, 330), use_container_width=True)
    else:
        st.info("No monthly data exists for this year.")

    section("YEARLY COMPARISON")
    if year_summary_rows:
        yearly_df = pd.DataFrame(year_summary_rows)
        st.dataframe(
            yearly_df.style.format({
                "Total KGs": "{:,.2f}",
                "Buying cost (KES)": "KES {:,.2f}",
                "Selling value (KES)": "KES {:,.2f}",
                "Expenses (KES)": "KES {:,.2f}",
                "Gross profit (KES)": "KES {:,.2f}",
                "Net profit (KES)": "KES {:,.2f}",
                "Net margin (%)": "{:.1f}%",
            }),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("No yearly records to compare.")

    section("YEAR-END OUTLOOK")
    o1, o2, o3 = st.columns(3)
    with o1:
        st.metric("Actual net profit to date", fmt_kes(summary["net"]))
    with o2:
        st.metric("Projected year-end net profit", fmt_kes(projected_year_profit))
    with o3:
        st.metric(
            "Projected vs recorded-month target",
            f"{projected_variance_pct:+.1f}%" if summary["target"] else "No target",
        )
    st.caption(
        "The projection uses actual net profit plus the average net profit of recorded months "
        "multiplied by the remaining calendar months. It is an estimate only."
    )

    section("DOWNLOAD YEARLY WORD REPORT")
    report_download_button("yearly_word_report")


# ---------------------------------------------------------------------
# MONTHLY REPORT
# ---------------------------------------------------------------------

elif page == "Monthly Report":
    page_header(
        "MONTHLY REPORT & PRINTABLE SUMMARY",
        f"Full January–December ledger, target assessment and downloadable Word report • {selected_year}",
    )

    section("MONTH-BY-MONTH FINANCIAL REPORT")
    st.dataframe(
        report_table,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Total Kgs": st.column_config.NumberColumn(format="%.2f"),
            "Buying Cost (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Selling Value (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Gross Profit (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Expenses (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Net Profit (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Net Margin (%)": st.column_config.NumberColumn(format="%.1f%%"),
            "Net Profit Target (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Target Variance (%)": st.column_config.NumberColumn(format="%+.1f%%"),
        },
    )

    section("YEAR-TO-DATE TOTALS")
    totals = st.columns(4)
    with totals[0]:
        st.metric("Total KGs", f'{summary["kgs"]:,.2f}')
    with totals[1]:
        st.metric("Buying cost", fmt_kes(summary["buying"]))
    with totals[2]:
        st.metric("Selling value", fmt_kes(summary["selling"]))
    with totals[3]:
        st.metric("Net profit", fmt_kes(summary["net"]))

    section("AUTOMATIC MANAGEMENT COMMENTARY")
    if summary["months"]:
        latest_month = year_data.iloc[-1]["Month"]
        commentary = (
            f"At the end of {latest_month} {selected_year}, "
            f"{summary['months']} month(s) have been recorded. "
            f"The business supplied {summary['kgs']:,.2f} kg of tea leaf, "
            f"with a total selling value of {fmt_kes(summary['selling'])} "
            f"against buying costs of {fmt_kes(summary['buying'])}. "
            f"Gross profit is {fmt_kes(summary['gross'])}. After expenses "
            f"of {fmt_kes(summary['expenses'])}, net profit is "
            f"{fmt_kes(summary['net'])}, giving a net margin of "
            f"{summary['margin']:.1f}%. Average buying price is "
            f"KES {summary['avg_buy']:.2f}/kg and average selling price is "
            f"KES {summary['avg_sell']:.2f}/kg."
        )
        st.write(commentary)
    else:
        st.info("No records are available to create commentary.")

    section("DOWNLOAD REPORTS")
    report_download_button("monthly_page_word_report")

    csv_data = report_table.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ DOWNLOAD MONTHLY REPORT AS CSV",
        data=csv_data,
        file_name=f"SL_Tea_Ventures_Monthly_Report_{selected_year}.csv",
        mime="text/csv",
        use_container_width=True,
        key="monthly_csv",
    )

    full_csv = calculated.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ DOWNLOAD COMPLETE BUSINESS LEDGER AS CSV",
        data=full_csv,
        file_name="SL_Tea_Ventures_Complete_Ledger.csv",
        mime="text/csv",
        use_container_width=True,
        key="full_ledger_csv",
    )


# ---------------------------------------------------------------------
# FOOTER
# ---------------------------------------------------------------------

st.markdown(
    """
    <div class="footer">
        🍃 <b>S&L TEA VENTURES</b><br>
        Tea Leaf Business Intelligence & Financial Performance Dashboard<br>
        All financial amounts are shown in Kenyan Shillings (KES).
    </div>
    """,
    unsafe_allow_html=True,
)
