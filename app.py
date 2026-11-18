import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import date

# ---------------------------------------------------------------------
# MONTHLY TRADING DASHBOARD
# Layout tailored to a buying-and-selling business that trades by weight.
# Demo figures are illustrative. Replace the sample records with your data.
# Run: streamlit run app.py
# ---------------------------------------------------------------------

st.set_page_config(
    page_title="Monthly Trading | Business Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ------------------------------- COLOUR THEME ---------------------------
NAVY = "#243F63"
BLUE = "#2F6698"
PALE_BLUE = "#EAF2F8"
GREEN = "#239B68"
PALE_GREEN = "#E8F5EE"
INK = "#203044"
MUTED = "#718096"
WHITE = "#FFFFFF"
BORDER = "#DCE5EC"

st.markdown(
    f"""
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
      html, body, [class*="css"] {{ font-family: Inter, sans-serif; }}
      .stApp {{ background: #F3F6F9; color: {INK}; }}
      [data-testid="stSidebar"] {{
        background: linear-gradient(180deg, {NAVY} 0%, #31577A 100%);
      }}
      [data-testid="stSidebar"] * {{ color: #FFFFFF !important; }}
      [data-testid="stSidebar"] [data-testid="stRadio"] label {{
        background: rgba(255,255,255,.07);
        border: 1px solid rgba(255,255,255,.10);
        border-radius: 10px; padding: 9px 12px; margin: 4px 0;
      }}
      [data-testid="stSidebar"] [data-testid="stRadio"] label:hover {{
        background: rgba(255,255,255,.16);
      }}
      [data-testid="stSidebar"] [data-testid="stRadio"] [aria-checked="true"] + div {{
        color: white;
      }}
      .brand-wrap {{ text-align:center; padding: 12px 4px 22px; }}
      .brand-mark {{ font-size:34px; line-height:1.2; }}
      .brand-name {{ color:white; font-size:20px; font-weight:800; letter-spacing:2px; }}
      .brand-sub {{ color:#D8E7F3; font-size:10px; letter-spacing:1.4px; }}
      .header {{
        background: linear-gradient(110deg, {NAVY}, #3D6D91);
        color:white; padding:18px 22px; border-radius:16px; margin-bottom:14px;
        box-shadow: 0 5px 18px rgba(36,63,99,.12);
      }}
      .header-title {{ font-size:25px; font-weight:800; line-height:1.25; }}
      .header-sub {{ font-size:12px; color:#DCEAF4; margin-top:3px; }}
      .section-kicker {{ color:{BLUE}; font-size:11px; font-weight:800; letter-spacing:1.1px;
                         text-transform:uppercase; margin:5px 0 9px; }}
      .metric-card {{
        background:{WHITE}; border:1px solid {BORDER}; border-radius:14px;
        padding:14px 15px; min-height:118px; box-shadow:0 2px 8px rgba(32,48,68,.035);
      }}
      .metric-top {{ display:flex; align-items:center; justify-content:space-between; }}
      .metric-icon {{
        width:38px; height:38px; display:flex; align-items:center; justify-content:center;
        border-radius:11px; font-size:19px; background:{PALE_BLUE};
      }}
      .metric-label {{ color:{MUTED}; font-size:10px; font-weight:800; letter-spacing:.6px; }}
      .metric-value {{ color:{INK}; font-size:23px; font-weight:800; margin-top:7px; }}
      .metric-note {{ color:{GREEN}; font-size:10px; margin-top:2px; }}
      .panel-title {{ color:{INK}; font-size:13px; font-weight:800; margin:0 0 7px 1px; }}
      .panel-caption {{ color:{MUTED}; font-size:10px; margin-top:-5px; margin-bottom:5px; }}
      .insight-row {{ display:flex; gap:10px; align-items:flex-start; padding:9px 3px;
                      border-bottom:1px solid #EDF1F4; color:#526174; font-size:12px; }}
      .insight-row:last-child {{ border-bottom:0; }}
      .insight-icon {{ font-size:17px; }}
      .insight-row b {{ color:{INK}; }}
      .soft-note {{ background:{PALE_BLUE}; border:1px solid #D7E6F2; border-radius:11px;
                    padding:11px 13px; color:{NAVY}; font-size:12px; }}
      div[data-testid="stMetric"] {{ background:white; border:1px solid {BORDER};
          border-radius:12px; padding:12px; }}
      div[data-testid="stDataFrame"] {{ border:1px solid {BORDER}; border-radius:12px; }}
      .stButton button {{ border-radius:9px; }}
      footer {{visibility:hidden;}}
    </style>
    """,
    unsafe_allow_html=True,
)

# ------------------------------- SAMPLE TRADING DATA --------------------
# The values below are sample display data only.
# Your business measures:
# Buying cost = buying price per kg x total kg
# Sales revenue = selling price per kg x total kg
# Gross profit = sales revenue - buying cost
# Net profit = gross profit - expenses

trading = pd.DataFrame({
    "Month": ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    "Total Kgs": [4200, 4650, 5100, 4900, 5600, 5850, 6100, 5700, 5300, 6250, 6600, 6900],
    "Buying Price/kg": [48, 49, 50, 51, 50, 52, 53, 52, 54, 55, 54, 56],
    "Selling Price/kg": [58, 60, 61, 62, 61, 64, 65, 64, 66, 67, 67, 69],
    "Expenses": [18000, 19500, 21000, 20500, 22000, 23500, 24000, 23000, 22500, 25000, 26000, 27000],
})
trading["Total Buying Cost"] = trading["Total Kgs"] * trading["Buying Price/kg"]
trading["Total Sales Revenue"] = trading["Total Kgs"] * trading["Selling Price/kg"]
trading["Gross Profit"] = trading["Total Sales Revenue"] - trading["Total Buying Cost"]
trading["Net Profit"] = trading["Gross Profit"] - trading["Expenses"]
trading["Profit Margin"] = trading["Net Profit"] / trading["Total Sales Revenue"] * 100

total_kgs = int(trading["Total Kgs"].sum())
total_buying = trading["Total Buying Cost"].sum()
total_revenue = trading["Total Sales Revenue"].sum()
total_gross = trading["Gross Profit"].sum()
total_expenses = trading["Expenses"].sum()
total_net = trading["Net Profit"].sum()
margin = total_net / total_revenue * 100 if total_revenue else 0

def money(value):
    return f"KSh {value:,.0f}"

def panel_heading(title, caption=None):
    st.markdown(f'<div class="panel-title">{title}</div>', unsafe_allow_html=True)
    if caption:
        st.markdown(f'<div class="panel-caption">{caption}</div>', unsafe_allow_html=True)

def chart_style(fig, height=250):
    fig.update_layout(
        height=height, margin=dict(l=8, r=12, t=12, b=8),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", size=10, color=MUTED),
        legend=dict(font=dict(size=9)),
    )
    fig.update_xaxes(showgrid=False, zeroline=False, title=None)
    fig.update_yaxes(showgrid=True, gridcolor="#EDF1F4", zeroline=False, title=None)
    return fig

# -------------------------------- SIDEBAR -------------------------------
with st.sidebar:
    st.markdown(
        '<div class="brand-wrap"><div class="brand-mark">◈</div>'
        '<div class="brand-name">VEDAKARNA</div>'
        '<div class="brand-sub">MONTHLY TRADING</div></div>',
        unsafe_allow_html=True,
    )
    st.markdown("**WORKSPACE MENU**")
    page = st.radio(
        "Navigation",
        [
            "🏠  Home",
            "📊  Trading Dashboard",
            "🛒  Purchases",
            "🚚  Sales",
            "⚖️  Stock & Volume",
            "💸  Expenses",
            "📈  Profit Analysis",
            "📅  Monthly Reports",
            "🧾  Data Input",
        ],
        index=0,
        label_visibility="collapsed",
    )
    st.markdown("---")
    st.markdown("**REPORTING PERIOD**")
    selected_year = st.selectbox("Year", [2026, 2025, 2024], label_visibility="collapsed")
    st.markdown(
        '<div style="font-size:10px;color:#D8E7F3;padding:10px 2px;">'
        'Blue • White • Green<br>Business performance workspace</div>',
        unsafe_allow_html=True,
    )

# -------------------------------- HEADER --------------------------------
st.markdown(
    f'<div class="header"><div class="header-title">📊 &nbsp; {page.split("  ",1)[-1]}</div>'
    f'<div class="header-sub">Monthly buying, selling, expenses and profit overview'
    f' &nbsp; • &nbsp; {selected_year}</div></div>',
    unsafe_allow_html=True,
)

# Month filter is presented consistently across the workspace.
filter_col, info_col = st.columns([1.25, 4])
with filter_col:
    selected_month = st.selectbox("Select month", ["All months"] + trading["Month"].tolist())
with info_col:
    st.markdown(
        f'<div style="text-align:right;padding-top:30px;color:{MUTED};font-size:11px;">'
        f'Last updated: {date.today().strftime("%d %b %Y")}</div>',
        unsafe_allow_html=True,
    )

view = trading.copy()
if selected_month != "All months":
    view = view[view["Month"] == selected_month]

# Aggregate selected period values for cards.
kgs = int(view["Total Kgs"].sum())
buying = view["Total Buying Cost"].sum()
sales = view["Total Sales Revenue"].sum()
gross = view["Gross Profit"].sum()
expenses = view["Expenses"].sum()
net = view["Net Profit"].sum()
selected_margin = (net / sales * 100) if sales else 0

# -------------------------------- HOME ----------------------------------
if page.startswith("🏠") or page.startswith("📊"):
    st.markdown('<div class="section-kicker">Business at a glance</div>', unsafe_allow_html=True)
    cards = [
        ("⚖️", "TOTAL QUANTITY TRADED", f"{kgs:,.0f} kg", "Volume purchased and sold", PALE_BLUE),
        ("🛒", "TOTAL BUYING COST", money(buying), "Buying price × total kilograms", PALE_BLUE),
        ("💰", "TOTAL SALES REVENUE", money(sales), "Selling price × total kilograms", PALE_GREEN),
        ("📊", "GROSS PROFIT", money(gross), "Revenue less buying cost", PALE_GREEN),
        ("✅", "NET PROFIT", money(net), f"Net margin: {selected_margin:.1f}%", PALE_GREEN),
    ]
    metric_cols = st.columns(5, gap="small")
    for col, (icon, label, value, note, tint) in zip(metric_cols, cards):
        with col:
            st.markdown(
                f'<div class="metric-card"><div class="metric-top">'
                f'<div class="metric-label">{label}</div>'
                f'<div class="metric-icon" style="background:{tint}">{icon}</div></div>'
                f'<div class="metric-value">{value}</div>'
                f'<div class="metric-note">{note}</div></div>',
                unsafe_allow_html=True,
            )

    st.write("")
    left, right = st.columns([1.55, 1], gap="small")
    with left:
        with st.container(border=True):
            panel_heading("📈  Monthly Sales Revenue", "Total value of goods sold each month")
            fig = px.area(view, x="Month", y="Total Sales Revenue")
            fig.update_traces(line_color=BLUE, fillcolor="rgba(47,102,152,.16)")
            fig.update_yaxes(tickprefix="KSh ", tickformat=",.0f")
            st.plotly_chart(chart_style(fig, 260), use_container_width=True, config={"displayModeBar": False})
    with right:
        with st.container(border=True):
            panel_heading("💹  Monthly Net Profit", "Profit remaining after expenses")
            fig = px.bar(view, x="Month", y="Net Profit")
            fig.update_traces(marker_color=GREEN, marker_line_width=0)
            fig.update_yaxes(tickprefix="KSh ", tickformat=",.0f")
            st.plotly_chart(chart_style(fig, 260), use_container_width=True, config={"displayModeBar": False})

    a, b = st.columns([1.2, 1.35], gap="small")
    with a:
        with st.container(border=True):
            panel_heading("⚖️  Trading Volume", "Kilograms handled per month")
            fig = px.bar(view, x="Month", y="Total Kgs")
            fig.update_traces(marker_color="#75A9C9")
            fig.update_yaxes(ticksuffix=" kg")
            st.plotly_chart(chart_style(fig, 235), use_container_width=True, config={"displayModeBar": False})
    with b:
        with st.container(border=True):
            panel_heading("🧮  Buying Cost vs Sales Revenue", "Compare the value of purchases and sales")
            comparison = view.melt(
                id_vars="Month",
                value_vars=["Total Buying Cost", "Total Sales Revenue"],
                var_name="Trading Activity", value_name="Amount",
            )
            fig = px.bar(comparison, x="Month", y="Amount", color="Trading Activity",
                         barmode="group",
                         color_discrete_map={"Total Buying Cost": BLUE, "Total Sales Revenue": GREEN})
            fig.update_yaxes(tickprefix="KSh ", tickformat=",.0f")
            st.plotly_chart(chart_style(fig, 235), use_container_width=True, config={"displayModeBar": False})

    c1, c2 = st.columns([1.15, 1], gap="small")
    with c1:
        with st.container(border=True):
            panel_heading("🧾  Monthly Trading Summary")
            summary = view[["Month", "Total Kgs", "Buying Price/kg", "Selling Price/kg",
                            "Total Buying Cost", "Total Sales Revenue", "Expenses",
                            "Gross Profit", "Net Profit"]].copy()
            st.dataframe(
                summary.style.format({
                    "Total Kgs": "{:,.0f}",
                    "Buying Price/kg": "KSh {:,.2f}",
                    "Selling Price/kg": "KSh {:,.2f}",
                    "Total Buying Cost": "KSh {:,.0f}",
                    "Total Sales Revenue": "KSh {:,.0f}",
                    "Expenses": "KSh {:,.0f}",
                    "Gross Profit": "KSh {:,.0f}",
                    "Net Profit": "KSh {:,.0f}",
                }),
                use_container_width=True, hide_index=True,
            )
    with c2:
        with st.container(border=True):
            panel_heading("✨  Business Highlights")
            insights = [
                ("📦", f"Total quantity traded: <b>{kgs:,.0f} kg</b>."),
                ("🛒", f"Total buying investment: <b>{money(buying)}</b>."),
                ("💰", f"Sales generated: <b>{money(sales)}</b>."),
                ("📈", f"Gross profit: <b>{money(gross)}</b>."),
                ("✅", f"Net profit after expenses: <b>{money(net)}</b>."),
            ]
            for icon, message in insights:
                st.markdown(
                    f'<div class="insight-row"><span class="insight-icon">{icon}</span>'
                    f'<span>{message}</span></div>', unsafe_allow_html=True
                )

# ---------------------------- PURCHASES ---------------------------------
elif page.startswith("🛒"):
    st.markdown('<div class="section-kicker">Purchase management</div>', unsafe_allow_html=True)
    x1, x2, x3 = st.columns(3)
    x1.metric("Total kilograms purchased", f"{kgs:,.0f} kg")
    x2.metric("Total buying cost", money(buying))
    x3.metric("Average buying price/kg", f"KSh {view['Buying Price/kg'].mean():,.2f}" if len(view) else "—")
    with st.container(border=True):
        panel_heading("🛒  Purchase Records")
        st.dataframe(view[["Month", "Total Kgs", "Buying Price/kg", "Total Buying Cost"]], use_container_width=True, hide_index=True)

# -------------------------------- SALES ---------------------------------
elif page.startswith("🚚"):
    st.markdown('<div class="section-kicker">Sales management</div>', unsafe_allow_html=True)
    x1, x2, x3 = st.columns(3)
    x1.metric("Total kilograms sold", f"{kgs:,.0f} kg")
    x2.metric("Total sales revenue", money(sales))
    x3.metric("Average selling price/kg", f"KSh {view['Selling Price/kg'].mean():,.2f}" if len(view) else "—")
    with st.container(border=True):
        panel_heading("🚚  Sales Records")
        st.dataframe(view[["Month", "Total Kgs", "Selling Price/kg", "Total Sales Revenue"]], use_container_width=True, hide_index=True)

# --------------------------- STOCK & VOLUME -----------------------------
elif page.startswith("⚖️"):
    st.markdown('<div class="section-kicker">Quantity and volume overview</div>', unsafe_allow_html=True)
    st.metric("Total kilograms traded", f"{kgs:,.0f} kg")
    fig = px.bar(view, x="Month", y="Total Kgs", text="Total Kgs")
    fig.update_traces(marker_color=BLUE, textposition="outside")
    st.plotly_chart(chart_style(fig, 340), use_container_width=True, config={"displayModeBar": False})
    with st.container(border=True):
        panel_heading("⚖️  Monthly Quantity Records")
        st.dataframe(view[["Month", "Total Kgs"]], use_container_width=True, hide_index=True)

# ------------------------------- EXPENSES -------------------------------
elif page.startswith("💸"):
    st.markdown('<div class="section-kicker">Operating expenses</div>', unsafe_allow_html=True)
    x1, x2 = st.columns(2)
    x1.metric("Total expenses", money(expenses))
    x2.metric("Expenses as % of sales", f"{expenses / sales * 100:.1f}%" if sales else "—")
    fig = px.bar(view, x="Month", y="Expenses")
    fig.update_traces(marker_color="#75A9C9")
    fig.update_yaxes(tickprefix="KSh ", tickformat=",.0f")
    st.plotly_chart(chart_style(fig, 320), use_container_width=True, config={"displayModeBar": False})
    with st.container(border=True):
        panel_heading("💸  Monthly Expense Records")
        st.dataframe(view[["Month", "Expenses"]], use_container_width=True, hide_index=True)

# ---------------------------- PROFIT ANALYSIS ---------------------------
elif page.startswith("📈"):
    st.markdown('<div class="section-kicker">Profitability overview</div>', unsafe_allow_html=True)
    x1, x2, x3 = st.columns(3)
    x1.metric("Gross profit", money(gross))
    x2.metric("Net profit", money(net))
    x3.metric("Net profit margin", f"{selected_margin:.2f}%")
    profit_data = view.melt(id_vars="Month", value_vars=["Gross Profit", "Net Profit"],
                            var_name="Profit Type", value_name="Amount")
    fig = px.bar(profit_data, x="Month", y="Amount", color="Profit Type", barmode="group",
                 color_discrete_map={"Gross Profit": BLUE, "Net Profit": GREEN})
    fig.update_yaxes(tickprefix="KSh ", tickformat=",.0f")
    st.plotly_chart(chart_style(fig, 330), use_container_width=True, config={"displayModeBar": False})
    with st.container(border=True):
        panel_heading("📈  Profit Breakdown")
        st.dataframe(view[["Month", "Total Sales Revenue", "Total Buying Cost", "Expenses",
                           "Gross Profit", "Net Profit", "Profit Margin"]].style.format({
            "Total Sales Revenue": "KSh {:,.0f}", "Total Buying Cost": "KSh {:,.0f}",
            "Expenses": "KSh {:,.0f}", "Gross Profit": "KSh {:,.0f}",
            "Net Profit": "KSh {:,.0f}", "Profit Margin": "{:.2f}%"
        }), use_container_width=True, hide_index=True)

# ---------------------------- MONTHLY REPORTS ---------------------------
elif page.startswith("📅"):
    st.markdown('<div class="section-kicker">Monthly performance reports</div>', unsafe_allow_html=True)
    report = view[["Month", "Total Kgs", "Total Buying Cost", "Total Sales Revenue",
                   "Expenses", "Gross Profit", "Net Profit"]]
    st.dataframe(report, use_container_width=True, hide_index=True)
    st.download_button(
        "⬇️  Download monthly report (CSV)",
        data=report.to_csv(index=False).encode("utf-8"),
        file_name=f"monthly_trading_report_{selected_year}.csv",
        mime="text/csv",
    )

# -------------------------------- DATA INPUT ----------------------------
elif page.startswith("🧾"):
    st.markdown('<div class="section-kicker">Trading data workspace</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="soft-note">This section is arranged around your trading records: '
        '<b>year, month, buying price per kg, total kilograms, selling price per kg and expenses.</b> '
        'The dashboard design uses these fields to present buying cost, sales revenue, gross profit and net profit.</div>',
        unsafe_allow_html=True,
    )
    st.write("")
    with st.container(border=True):
        panel_heading("🧾  Upload Trading Records")
        uploaded = st.file_uploader("Choose a CSV file", type=["csv"])
        if uploaded is not None:
            uploaded_data = pd.read_csv(uploaded)
            st.success(f"File loaded: {len(uploaded_data):,} records")
            st.dataframe(uploaded_data, use_container_width=True)
    with st.container(border=True):
        panel_heading("📋  Expected Trading Fields")
        st.dataframe(pd.DataFrame({
            "Field": ["Year", "Month", "Buying Price/kg", "Total Kgs", "Selling Price/kg", "Expenses"],
            "Purpose": ["Reporting year", "Trading month", "Purchase rate per kilogram",
                        "Quantity traded", "Selling rate per kilogram", "Operating costs"],
        }), use_container_width=True, hide_index=True)

st.markdown(
    f'<div style="text-align:center;color:{MUTED};font-size:10px;padding:18px 0 6px;">'
    'VEDAKARNA • MONTHLY TRADING DASHBOARD • SAMPLE FIGURES FOR LAYOUT PREVIEW</div>',
    unsafe_allow_html=True,
)
