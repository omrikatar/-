"""

        FINledger  Advanced Financial Dashboard         
        Built with Streamlit + Plotly + Pandas            


pip install streamlit pandas plotly openpyxl numpy
Run: streamlit run financial_dashboard.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import re

# 
#  PAGE CONFIG
# 
st.set_page_config(
    page_title="FinLedger Dashboard",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 
#  CUSTOM CSS  (dark financial aesthetic)
# 
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=Syne:wght@400;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Syne', sans-serif;
    background-color: #0d0f14;
    color: #e8eaf0;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #12151d;
    border-right: 1px solid #1e2330;
}
section[data-testid="stSidebar"] * { font-family: 'IBM Plex Mono', monospace; font-size: 13px; }

/* Metric cards */
[data-testid="metric-container"] {
    background: #161923;
    border: 1px solid #1e2a40;
    border-radius: 12px;
    padding: 16px 20px;
}
[data-testid="stMetricLabel"]  { color: #7a8299 !important; font-size: 12px; letter-spacing: 1px; text-transform: uppercase; }
[data-testid="stMetricValue"]  { color: #e8eaf0 !important; font-family: 'IBM Plex Mono', monospace; font-size: 28px; }
[data-testid="stMetricDelta"]  { font-size: 13px; }

/* Headers */
h1 { font-family: 'Syne', sans-serif; font-weight: 800; font-size: 2.4rem; color: #e8eaf0; }
h2, h3 { font-family: 'Syne', sans-serif; font-weight: 700; color: #c8cfe0; }

/* Dividers */
hr { border-color: #1e2330; }

/* Tables */
[data-testid="stDataFrame"] { border: 1px solid #1e2330; border-radius: 8px; }

/* Inputs */
.stTextInput > div > input, .stNumberInput > div > input, .stSelectbox > div > div {
    background: #1a1e2a !important;
    border: 1px solid #2a3050 !important;
    color: #e8eaf0 !important;
    font-family: 'IBM Plex Mono', monospace;
}

/* Buttons */
.stButton > button {
    background: #2355f5;
    color: white;
    border: none;
    border-radius: 8px;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
    letter-spacing: 0.5px;
}
.stButton > button:hover { background: #3d6bff; }

/* Upload area */
[data-testid="stFileUploader"] {
    background: #161923;
    border: 1px dashed #2a3050;
    border-radius: 12px;
}

/* Section separators */
.section-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: #2355f5;
    margin-bottom: 16px;
    padding-bottom: 8px;
    border-bottom: 1px solid #1e2330;
}

.kpi-row { gap: 16px; }

.amort-table { font-family: 'IBM Plex Mono', monospace; font-size: 12px; }

.tag-positive { color: #22d98a; }
.tag-negative { color: #f05454; }
</style>
""", unsafe_allow_html=True)


# 
#  KEYWORD CATEGORIZATION DICTIONARY
# 
CATEGORY_KEYWORDS = {
    " Food & Dining":     ["restaurant", "cafe", "coffee", "mcdonald", "pizza", "sushi",
                              "grocery", "supermarket", "uber eats", "deliveroo", "food",
                              "lunch", "dinner", "breakfast", "bakery", "bar"],
    " Housing":           ["rent", "mortgage", "property", "landlord", "maintenance",
                              "electric", "electricity", "water bill", "gas bill", "council tax",
                              "insurance home", "hoa"],
    " Transport":         ["uber", "lyft", "taxi", "bus", "metro", "train", "fuel",
                              "petrol", "parking", "toll", "transport", "flight", "airline",
                              "car lease", "auto"],
    " Salary & Income":   ["salary", "payroll", "income", "wage", "bonus", "freelance",
                              "consulting fee", "dividend", "interest earned", "refund",
                              "payment received", "transfer in"],
    " Shopping":          ["amazon", "ebay", "zara", "h&m", "clothing", "shoes",
                              "electronics", "apple", "samsung", "ikea", "shop", "store", "retail"],
    " Health":            ["pharmacy", "doctor", "dentist", "hospital", "clinic",
                              "prescription", "medicine", "gym", "fitness", "health", "optical"],
    " Utilities & Tech":  ["netflix", "spotify", "apple subscription", "google",
                              "microsoft", "internet", "phone bill", "mobile", "software",
                              "subscription", "saas"],
    " Education":         ["university", "course", "udemy", "coursera", "school",
                              "tuition", "books", "training", "workshop", "seminar"],
    " Travel":            ["hotel", "airbnb", "booking.com", "expedia", "holiday",
                              "vacation", "resort", "travel", "visa"],
    " Tax & Fees":        ["tax", "vat", "levy", "penalty", "fine", "fee", "charge",
                              "commission", "bank fee", "atm fee"],
    " Business Expense":  ["office", "equipment", "supplies", "advertising", "marketing",
                              "printing", "postage", "accountant", "legal", "professional"],
}

VAT_ELIGIBLE_CATEGORIES = [" Business Expense", " Utilities & Tech", " Education"]


def categorize(description: str) -> str:
    if not isinstance(description, str):
        return " Uncategorized"
    desc_lower = description.lower()
    for category, keywords in CATEGORY_KEYWORDS.items():
        if any(kw in desc_lower for kw in keywords):
            return category
    return " Uncategorized"


# 
#  DATA LOADING & CLEANING
# 
def load_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
    return df


def detect_columns(df: pd.DataFrame):
    """Heuristically detect Date, Description, Amount columns."""
    date_col = desc_col = amount_col = None

    for col in df.columns:
        cl = col.lower().strip()
        if date_col is None and any(k in cl for k in ["date", "time", "period", "when"]):
            date_col = col
        if desc_col is None and any(k in cl for k in ["desc", "narr", "memo", "note", "payee", "merchant", "name"]):
            desc_col = col
        if amount_col is None and any(k in cl for k in ["amount", "value", "sum", "debit", "credit", "price", "total"]):
            amount_col = col

    # fallback: try dtypes
    if amount_col is None:
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                amount_col = col
                break
    if date_col is None:
        for col in df.columns:
            try:
                pd.to_datetime(df[col], errors="raise")
                date_col = col
                break
            except Exception:
                continue
    if desc_col is None:
        remaining = [c for c in df.columns if c not in [date_col, amount_col]]
        if remaining:
            desc_col = remaining[0]

    return date_col, desc_col, amount_col


def clean_data(df: pd.DataFrame, date_col: str, desc_col: str, amount_col: str) -> pd.DataFrame:
    out = df[[date_col, desc_col, amount_col]].copy()
    out.columns = ["Date", "Description", "Amount"]
    out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
    out.dropna(subset=["Date", "Amount"], inplace=True)
    # Parse amounts  strip currency symbols, commas
    if out["Amount"].dtype == object:
        out["Amount"] = out["Amount"].astype(str).str.replace(r"[^\d.\-\+]", "", regex=True)
        out["Amount"] = pd.to_numeric(out["Amount"], errors="coerce")
    out.dropna(subset=["Amount"], inplace=True)
    out["Description"] = out["Description"].fillna("Unknown")
    out["Category"] = out["Description"].apply(categorize)
    out["Month"] = out["Date"].dt.to_period("M").astype(str)
    out["Type"] = out["Amount"].apply(lambda x: "Income" if x > 0 else "Expense")
    return out.sort_values("Date")


# 
#  AMORTIZATION CALCULATOR
# 
def amortization_schedule(principal: float, annual_rate: float, months: int) -> pd.DataFrame:
    if annual_rate == 0:
        monthly_payment = principal / months
        rows = [{"Month": i + 1, "Payment": monthly_payment, "Principal": monthly_payment,
                 "Interest": 0.0, "Balance": principal - monthly_payment * (i + 1)} for i in range(months)]
        return pd.DataFrame(rows), monthly_payment

    r = annual_rate / 100 / 12
    monthly_payment = principal * r * (1 + r) ** months / ((1 + r) ** months - 1)
    balance = principal
    rows = []
    for i in range(1, months + 1):
        interest = balance * r
        principal_part = monthly_payment - interest
        balance -= principal_part
        rows.append({
            "Month": i,
            "Payment": round(monthly_payment, 2),
            "Principal": round(principal_part, 2),
            "Interest": round(interest, 2),
            "Balance": round(max(balance, 0), 2),
        })
    return pd.DataFrame(rows), monthly_payment


# 
#  PLOTLY THEME
# 
PLOTLY_LAYOUT = dict(
    paper_bgcolor="#0d0f14",
    plot_bgcolor="#0d0f14",
    font=dict(family="IBM Plex Mono, monospace", color="#c8cfe0", size=12),
    margin=dict(t=40, b=40, l=40, r=40),
    legend=dict(bgcolor="#12151d", bordercolor="#1e2330", borderwidth=1),
)

CATEGORY_COLORS = [
    "#2355f5", "#22d98a", "#f5a623", "#f05454", "#a855f7",
    "#06b6d4", "#f97316", "#84cc16", "#ec4899", "#14b8a6",
    "#6366f1", "#64748b",
]


# 
#  SIDEBAR
# 
with st.sidebar:
    st.markdown("##  FinLedger")
    st.markdown("---")

    # File Upload
    st.markdown('<div class="section-title"> Data Source</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Upload Excel / CSV", type=["csv", "xlsx", "xls"])

    st.markdown("---")

    # Tax Settings
    st.markdown('<div class="section-title"> Tax Settings</div>', unsafe_allow_html=True)
    apply_tax = st.toggle("Estimate Income Tax", value=False)
    tax_rate = st.slider("Tax Rate (%)", min_value=0, max_value=60, value=25, step=1)
    apply_vat = st.toggle("Show VAT-Deductible Table (\")", value=False)
    vat_rate = st.slider("VAT Rate (%)", min_value=0, max_value=30, value=17, step=1)

    st.markdown("---")

    # Mortgage / Loan Calculator
    st.markdown('<div class="section-title"> Mortgage / Loan</div>', unsafe_allow_html=True)
    loan_amount = st.number_input("Loan Amount (/$)", min_value=0.0, value=500_000.0, step=10_000.0, format="%.0f")
    loan_rate = st.number_input("Annual Interest Rate (%)", min_value=0.0, max_value=30.0, value=4.5, step=0.1)
    loan_term = st.number_input("Loan Term (months)", min_value=1, max_value=360, value=240, step=12)
    show_amort = st.toggle("Show Full Amortization Table", value=False)

    st.markdown("---")
    st.caption("FinLedger v1.0 · Senior Finance Dashboard")


# 
#  MAIN HEADER
# 
st.markdown("#  FinLedger")
st.markdown("**Advanced Financial Intelligence Dashboard** · Upload your transactions to begin")
st.markdown("---")


# 
#  DEMO DATA (when no file uploaded)
# 
def generate_demo_data() -> pd.DataFrame:
    np.random.seed(42)
    dates = pd.date_range("2024-01-01", "2024-12-31", freq="3D")
    descriptions_expenses = [
        "Restaurant Le Marais", "Uber Eats delivery", "Spotify Premium",
        "Amazon order", "Rent Payment", "Electric Bill", "Uber ride",
        "Gym membership", "Netflix subscription", "Office supplies",
        "Pharmacy purchase", "Zara clothing", "Petrol station",
        "Google Workspace", "Course on Udemy",
    ]
    descriptions_income = [
        "Salary payroll", "Freelance consulting fee", "Bonus payment",
        "Dividend income", "Payment received",
    ]
    rows = []
    for d in dates:
        if np.random.rand() < 0.15:
            rows.append({"Date": d, "Description": np.random.choice(descriptions_income),
                         "Amount": round(np.random.uniform(3000, 8000), 2)})
        else:
            rows.append({"Date": d, "Description": np.random.choice(descriptions_expenses),
                         "Amount": -round(np.random.uniform(20, 600), 2)})
    df = pd.DataFrame(rows)
    df["Category"] = df["Description"].apply(categorize)
    df["Month"] = df["Date"].dt.to_period("M").astype(str)
    df["Type"] = df["Amount"].apply(lambda x: "Income" if x > 0 else "Expense")
    return df


# 
#  LOAD DATA
# 
if uploaded_file:
    try:
        raw_df = load_file(uploaded_file)
        date_col, desc_col, amount_col = detect_columns(raw_df)

        with st.expander(" Column Mapping  Review & Override", expanded=False):
            cols = raw_df.columns.tolist()
            c1, c2, c3 = st.columns(3)
            date_col = c1.selectbox("Date Column", cols, index=cols.index(date_col) if date_col else 0)
            desc_col = c2.selectbox("Description Column", cols, index=cols.index(desc_col) if desc_col else 0)
            amount_col = c3.selectbox("Amount Column", cols, index=cols.index(amount_col) if amount_col else 0)

        df = clean_data(raw_df, date_col, desc_col, amount_col)
        st.success(f" Loaded **{len(df):,}** transactions from `{uploaded_file.name}`")
    except Exception as e:
        st.error(f" Error loading file: {e}")
        df = generate_demo_data()
        st.info("Showing demo data instead.")
else:
    df = generate_demo_data()
    st.info(" No file uploaded  showing **demo data**. Upload your own CSV/Excel via the sidebar.")


# 
#  COMPUTED METRICS
# 
total_income = df[df["Type"] == "Income"]["Amount"].sum()
total_expenses = abs(df[df["Type"] == "Expense"]["Amount"].sum())
net_balance = total_income - total_expenses

# Tax adjustment
if apply_tax:
    tax_amount = total_income * (tax_rate / 100)
    net_income_after_tax = total_income - tax_amount
else:
    tax_amount = 0
    net_income_after_tax = total_income

# Mortgage monthly
amort_df, monthly_payment = amortization_schedule(loan_amount, loan_rate, int(loan_term))
annual_debt_service = monthly_payment * 12

# Monthly averages
months_count = df["Month"].nunique()
monthly_burn = total_expenses / months_count if months_count else 0
savings_rate = ((net_income_after_tax - total_expenses) / net_income_after_tax * 100) if net_income_after_tax > 0 else 0
debt_to_income = (annual_debt_service / net_income_after_tax * 100) if net_income_after_tax > 0 else 0


# 
#  KPI ROW 1
# 
st.markdown('<div class="section-title"> Key Performance Indicators</div>', unsafe_allow_html=True)

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric(" Total Income", f"${total_income:,.0f}",
          f"-{tax_rate}% tax = ${net_income_after_tax:,.0f}" if apply_tax else None)
k2.metric(" Total Expenses", f"${total_expenses:,.0f}", f"${monthly_burn:,.0f}/mo avg")
k3.metric(" Net Balance", f"${net_balance:,.0f}",
          f"{'' if net_balance > 0 else ''} {abs(net_balance / total_income * 100):.1f}% of income" if total_income else None,
          delta_color="normal")
k4.metric(" Monthly Burn Rate", f"${monthly_burn:,.0f}", "avg per month")
k5.metric(" Savings Rate", f"{savings_rate:.1f}%",
          "healthy " if savings_rate > 20 else "below 20% ",
          delta_color="normal" if savings_rate > 20 else "inverse")

st.markdown("---")

# KPI ROW 2  Mortgage
k6, k7, k8, k9, _ = st.columns(5)
k6.metric(" Monthly Repayment", f"${monthly_payment:,.0f}", f"{loan_term} mo term")
k7.metric(" Annual Debt Service", f"${annual_debt_service:,.0f}", "")
k8.metric(" Debt-to-Income", f"{debt_to_income:.1f}%",
          "manageable " if debt_to_income < 36 else "high ",
          delta_color="normal" if debt_to_income < 36 else "inverse")
k9.metric(" Net After Debt", f"${net_income_after_tax - annual_debt_service:,.0f}", "annual")

st.markdown("---")


# 
#  CHARTS ROW 1    Donut + Trend
# 
st.markdown('<div class="section-title"> Visualizations</div>', unsafe_allow_html=True)

col_donut, col_trend = st.columns([1, 2])

#  Donut Chart 
with col_donut:
    st.markdown("#### Expense Breakdown")
    expense_df = df[df["Type"] == "Expense"].copy()
    expense_df["Amount"] = expense_df["Amount"].abs()
    cat_summary = expense_df.groupby("Category")["Amount"].sum().reset_index()
    cat_summary = cat_summary.sort_values("Amount", ascending=False)

    fig_donut = px.pie(
        cat_summary,
        names="Category",
        values="Amount",
        hole=0.55,
        color_discrete_sequence=CATEGORY_COLORS,
    )
    fig_donut.update_traces(
        textposition="outside",
        textinfo="percent+label",
        textfont=dict(size=11, family="IBM Plex Mono"),
        hovertemplate="<b>%{label}</b><br>$%{value:,.0f}<br>%{percent}<extra></extra>",
    )
    fig_donut.update_layout(
        **PLOTLY_LAYOUT,
        showlegend=False,
        annotations=[dict(
            text=f"<b>${total_expenses:,.0f}</b><br>Total",
            x=0.5, y=0.5, font=dict(size=14, family="IBM Plex Mono", color="#e8eaf0"),
            showarrow=False
        )],
        height=380,
    )
    st.plotly_chart(fig_donut, use_container_width=True)

#  Trend Line 
with col_trend:
    st.markdown("#### Income vs Expenses  Monthly Trend")
    monthly = df.groupby(["Month", "Type"])["Amount"].apply(
        lambda x: x[x > 0].sum() if x.name[1] == "Income" else x[x < 0].abs().sum()
    ).reset_index()
    monthly.columns = ["Month", "Type", "Amount"]

    fig_trend = px.line(
        monthly, x="Month", y="Amount", color="Type",
        markers=True,
        color_discrete_map={"Income": "#22d98a", "Expense": "#f05454"},
    )
    fig_trend.update_traces(line_width=2.5, marker_size=6)
    fig_trend.update_layout(
        **PLOTLY_LAYOUT,
        height=380,
        xaxis=dict(showgrid=False, tickangle=-30),
        yaxis=dict(showgrid=True, gridcolor="#1e2330", tickprefix="$"),
        hovermode="x unified",
    )
    st.plotly_chart(fig_trend, use_container_width=True)

#  Waterfall Chart 
st.markdown("#### Net Balance Waterfall")
expense_cats = expense_df.groupby("Category")["Amount"].sum().reset_index()
expense_cats = expense_cats.sort_values("Amount", ascending=False)

wf_measure = ["absolute"] + ["relative"] * len(expense_cats) + ["total"]
wf_x = ["Gross Income"] + expense_cats["Category"].tolist() + ["Net Balance"]
wf_y = [total_income] + [-v for v in expense_cats["Amount"]] + [net_balance]
wf_text = [f"${abs(v):,.0f}" for v in wf_y]

fig_waterfall = go.Figure(go.Waterfall(
    measure=wf_measure,
    x=wf_x,
    y=wf_y,
    text=wf_text,
    textposition="outside",
    textfont=dict(family="IBM Plex Mono", size=11, color="#c8cfe0"),
    connector=dict(line=dict(color="#2a3050", width=1, dash="dot")),
    increasing=dict(marker=dict(color="#22d98a", line=dict(width=0))),
    decreasing=dict(marker=dict(color="#f05454", line=dict(width=0))),
    totals=dict(marker=dict(color="#2355f5", line=dict(width=0))),
    hovertemplate="<b>%{x}</b><br>$%{y:,.0f}<extra></extra>",
))
fig_waterfall.update_layout(
    **PLOTLY_LAYOUT,
    height=450,
    xaxis=dict(tickangle=-30, showgrid=False),
    yaxis=dict(showgrid=True, gridcolor="#1e2330", tickprefix="$"),
)
st.plotly_chart(fig_waterfall, use_container_width=True)

st.markdown("---")


# 
#  MORTGAGE AMORTIZATION
# 
st.markdown('<div class="section-title"> Mortgage & Debt Service Analysis</div>', unsafe_allow_html=True)

c_loan1, c_loan2 = st.columns(2)

with c_loan1:
    total_interest = amort_df["Interest"].sum()
    total_paid = monthly_payment * int(loan_term)
    st.markdown(f"""
    | Detail | Amount |
    |---|---|
    | Loan Principal | `${loan_amount:,.0f}` |
    | Monthly Payment | `${monthly_payment:,.2f}` |
    | Total Interest Paid | `${total_interest:,.0f}` |
    | Total Cost of Loan | `${total_paid:,.0f}` |
    | Interest Overhead | `{total_interest/loan_amount*100:.1f}%` |
    """)

with c_loan2:
    # Principal vs Interest over time (area chart)
    amort_monthly = amort_df.copy()
    amort_monthly["Year"] = ((amort_monthly["Month"] - 1) // 12) + 1
    amort_yearly = amort_monthly.groupby("Year")[["Principal", "Interest"]].sum().reset_index()

    fig_amort = go.Figure()
    fig_amort.add_trace(go.Bar(name="Principal", x=amort_yearly["Year"], y=amort_yearly["Principal"],
                                marker_color="#22d98a", hovertemplate="Year %{x}<br>Principal: $%{y:,.0f}<extra></extra>"))
    fig_amort.add_trace(go.Bar(name="Interest", x=amort_yearly["Year"], y=amort_yearly["Interest"],
                                marker_color="#f05454", hovertemplate="Year %{x}<br>Interest: $%{y:,.0f}<extra></extra>"))
    fig_amort.update_layout(
        **PLOTLY_LAYOUT, barmode="stack", height=280,
        xaxis=dict(title="Year", showgrid=False),
        yaxis=dict(showgrid=True, gridcolor="#1e2330", tickprefix="$"),
        title="Annual Principal vs Interest Split",
    )
    st.plotly_chart(fig_amort, use_container_width=True)

if show_amort:
    st.markdown("**Full Amortization Schedule**")
    st.dataframe(
        amort_df.style
        .format({"Payment": "${:,.2f}", "Principal": "${:,.2f}",
                 "Interest": "${:,.2f}", "Balance": "${:,.2f}"})
        .background_gradient(subset=["Balance"], cmap="Blues_r"),
        height=300,
        use_container_width=True,
    )

st.markdown("---")


# 
#  TAX & VAT SECTION
# 
if apply_tax or apply_vat:
    st.markdown('<div class="section-title"> Tax & VAT Analysis</div>', unsafe_allow_html=True)
    t1, t2 = st.columns(2)

    if apply_tax:
        with t1:
            st.markdown("#### Income Tax Summary")
            income_monthly = df[df["Type"] == "Income"].groupby("Month")["Amount"].sum().reset_index()
            income_monthly["Tax"] = income_monthly["Amount"] * (tax_rate / 100)
            income_monthly["Net"] = income_monthly["Amount"] - income_monthly["Tax"]
            fig_tax = go.Figure()
            fig_tax.add_trace(go.Bar(name="Net Income", x=income_monthly["Month"], y=income_monthly["Net"],
                                      marker_color="#22d98a"))
            fig_tax.add_trace(go.Bar(name=f"Tax ({tax_rate}%)", x=income_monthly["Month"], y=income_monthly["Tax"],
                                      marker_color="#f5a623"))
            fig_tax.update_layout(**PLOTLY_LAYOUT, barmode="stack", height=320,
                                   xaxis=dict(tickangle=-30, showgrid=False),
                                   yaxis=dict(gridcolor="#1e2330", tickprefix="$"),
                                   title="Monthly Income: Net vs Tax")
            st.plotly_chart(fig_tax, use_container_width=True)

    if apply_vat:
        with t2:
            st.markdown("#### VAT-Deductible Expenses (\")")
            vat_df = df[df["Category"].isin(VAT_ELIGIBLE_CATEGORIES) & (df["Type"] == "Expense")].copy()
            vat_df["Amount"] = vat_df["Amount"].abs()
            vat_df["VAT Amount"] = (vat_df["Amount"] * vat_rate / 100).round(2)
            vat_df["Net (Ex-VAT)"] = (vat_df["Amount"] - vat_df["VAT Amount"]).round(2)

            st.dataframe(
                vat_df[["Date", "Description", "Category", "Amount", "VAT Amount", "Net (Ex-VAT)"]]
                .rename(columns={"Amount": "Gross Amount"})
                .style.format({
                    "Gross Amount": "${:.2f}", "VAT Amount": "${:.2f}", "Net (Ex-VAT)": "${:.2f}"
                }),
                height=300,
                use_container_width=True,
            )
            total_vat = vat_df["VAT Amount"].sum()
            st.metric("Total Reclaimable VAT", f"${total_vat:,.2f}", f"@ {vat_rate}%")

    st.markdown("---")


# 
#  TRANSACTIONS TABLE
# 
st.markdown('<div class="section-title"> Transaction Ledger</div>', unsafe_allow_html=True)

# Filters
col_f1, col_f2, col_f3 = st.columns(3)
all_cats = sorted(df["Category"].unique().tolist())
selected_cats = col_f1.multiselect("Filter by Category", all_cats, default=all_cats)
selected_type = col_f2.selectbox("Transaction Type", ["All", "Income", "Expense"])
selected_month = col_f3.selectbox("Month", ["All"] + sorted(df["Month"].unique().tolist()))

filtered = df.copy()
filtered = filtered[filtered["Category"].isin(selected_cats)]
if selected_type != "All":
    filtered = filtered[filtered["Type"] == selected_type]
if selected_month != "All":
    filtered = filtered[filtered["Month"] == selected_month]

def color_amount(val):
    color = "#22d98a" if val > 0 else "#f05454"
    return f"color: {color}; font-family: IBM Plex Mono; font-weight: 600"

display_df = filtered[["Date", "Description", "Category", "Amount", "Type", "Month"]].copy()
display_df["Date"] = display_df["Date"].dt.strftime("%Y-%m-%d")

st.dataframe(
    display_df.style
    .applymap(color_amount, subset=["Amount"])
    .format({"Amount": "${:,.2f}"}),
    height=400,
    use_container_width=True,
)

# Export
csv_out = filtered.to_csv(index=False).encode("utf-8")
st.download_button(" Export Filtered Data (CSV)", csv_out, "finledger_export.csv", "text/csv")

st.markdown("---")
st.caption(" FinLedger · Built with Streamlit + Plotly + Pandas · For educational & personal finance use only.")
