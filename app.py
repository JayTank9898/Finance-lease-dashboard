"""
Finance Lease Calculation Dashboard
====================================
Supports two accounting frameworks:
  1. Ind AS 116 (Right-of-Use asset / Lease liability model)
  2. AS 19 (Old Indian GAAP - Finance Lease model)

Outputs:
  - Lease Liability amortization schedule (finance cost + closing liability)
  - ROU Asset / Leased Asset depreciation schedule (with closing WDV)
  - Interest-free Security Deposit unwinding schedule (income + closing balance)
  - Downloadable Excel workbook with all schedules

Run with:
    pip install streamlit pandas numpy xlsxwriter python-dateutil
    streamlit run lease_dashboard.py
"""

import io
from datetime import date

import numpy as np
import pandas as pd
import streamlit as st
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Finance Lease Dashboard", layout="wide")

# ----------------------------------------------------------------------
# Helper functions
# ----------------------------------------------------------------------

def annual_to_monthly_rate(annual_rate_pct: float) -> float:
    """Effective annual rate -> effective monthly rate."""
    r = annual_rate_pct / 100
    return (1 + r) ** (1 / 12) - 1


def build_payment_schedule(base_rent, term_months, escalation_pct, escalation_freq_months):
    """Builds a list of monthly lease payments with periodic escalation."""
    payments = []
    rent = base_rent
    for m in range(1, term_months + 1):
        if escalation_freq_months > 0 and m > 1 and (m - 1) % escalation_freq_months == 0:
            rent = rent * (1 + escalation_pct / 100)
        payments.append(rent)
    return payments


def pv_of_payments(payments, monthly_rate, timing):
    """Present value of a payment stream, advance or arrears."""
    pv = 0.0
    for i, pmt in enumerate(payments, start=1):
        exponent = (i - 1) if timing == "Advance (start of month)" else i
        pv += pmt / ((1 + monthly_rate) ** exponent)
    return pv


def lease_liability_schedule(payments, monthly_rate, timing, start_date):
    """Builds month-wise lease liability amortization table."""
    opening_bal = pv_of_payments(payments, monthly_rate, timing)
    rows = []
    bal = opening_bal
    for i, pmt in enumerate(payments, start=1):
        period_date = start_date + relativedelta(months=i - 1)
        opening = bal
        if timing == "Advance (start of month)":
            after_payment = opening - pmt
            interest = after_payment * monthly_rate
            closing = after_payment + interest
        else:  # Arrears
            interest = opening * monthly_rate
            closing = opening + interest - pmt
        rows.append({
            "Month": i,
            "Date": period_date.strftime("%b-%Y"),
            "Opening Liability": round(opening, 2),
            "Interest / Finance Cost": round(interest, 2),
            "Lease Payment": round(pmt, 2),
            "Closing Liability": round(closing, 2),
        })
        bal = closing
    df = pd.DataFrame(rows)
    # Fix rounding drift on final closing balance
    if not df.empty:
        df.loc[df.index[-1], "Closing Liability"] = 0.0
    return df, round(opening_bal, 2)


def rou_depreciation_schedule(rou_initial, term_months, start_date, useful_life_months=None):
    """Straight-line depreciation of ROU asset / leased asset."""
    dep_period = term_months if not useful_life_months else min(term_months, useful_life_months)
    monthly_dep = rou_initial / dep_period if dep_period else 0
    rows = []
    bal = rou_initial
    for i in range(1, term_months + 1):
        period_date = start_date + relativedelta(months=i - 1)
        opening = bal
        dep = monthly_dep if i <= dep_period else 0
        closing = max(opening - dep, 0)
        rows.append({
            "Month": i,
            "Date": period_date.strftime("%b-%Y"),
            "Opening WDV": round(opening, 2),
            "Depreciation": round(dep, 2),
            "Closing WDV": round(closing, 2),
        })
        bal = closing
    return pd.DataFrame(rows)


def deposit_unwinding_schedule(deposit_amount, market_rate_pct, deposit_term_months, start_date):
    """Unwinding of discount on interest-free refundable security deposit."""
    monthly_rate = annual_to_monthly_rate(market_rate_pct)
    pv_deposit = deposit_amount / ((1 + monthly_rate) ** deposit_term_months)
    day1_diff = deposit_amount - pv_deposit  # treated as prepaid rent / ROU addition

    rows = []
    bal = pv_deposit
    for i in range(1, deposit_term_months + 1):
        period_date = start_date + relativedelta(months=i - 1)
        opening = bal
        income = opening * monthly_rate
        closing = opening + income
        rows.append({
            "Month": i,
            "Date": period_date.strftime("%b-%Y"),
            "Opening Deposit (Amortised Cost)": round(opening, 2),
            "Unwinding / Notional Interest Income": round(income, 2),
            "Closing Deposit (Amortised Cost)": round(closing, 2),
        })
        bal = closing
    df = pd.DataFrame(rows)
    if not df.empty:
        df.loc[df.index[-1], "Closing Deposit (Amortised Cost)"] = deposit_amount
    return df, round(pv_deposit, 2), round(day1_diff, 2)


def to_excel(sheets: dict) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    return output.getvalue()


# ----------------------------------------------------------------------
# Sidebar - Inputs
# ----------------------------------------------------------------------

st.title("📊 Finance Lease Calculation Dashboard")
st.caption("Ind AS 116 (ROU model) / AS 19 (Finance Lease model)")

with st.sidebar:
    st.header("1. Accounting Framework")
    standard = st.radio(
        "Select applicable standard",
        ["Ind AS 116", "AS 19"],
        help="Ind AS 116: all leases capitalised as Right-of-Use asset & liability.\n"
             "AS 19: only leases classified as 'finance lease' are capitalised.",
    )

    st.header("2. Lease Terms")
    start_date = st.date_input("Lease commencement date", value=date.today())
    term_years = st.number_input("Lease term (years)", min_value=0.0, value=5.0, step=0.5)
    term_months = int(round(term_years * 12))

    base_rent = st.number_input("Monthly lease rent (₹)", min_value=0.0, value=100000.0, step=1000.0)
    escalation_pct = st.number_input("Rent escalation (%)", min_value=0.0, value=5.0, step=0.5)
    escalation_freq_years = st.number_input("Escalation frequency (years)", min_value=0.0, value=1.0, step=1.0)
    escalation_freq_months = int(round(escalation_freq_years * 12))

    timing = st.selectbox("Payment timing", ["Arrears (end of month)", "Advance (start of month)"])

    rate_label = "Incremental borrowing rate (%)" if standard == "Ind AS 116" \
        else "Interest rate implicit in lease / IBR (%)"
    discount_rate = st.number_input(rate_label, min_value=0.0, value=10.0, step=0.25)

    st.header("3. Initial Costs / Adjustments")
    idc = st.number_input("Initial direct costs (₹)", min_value=0.0, value=0.0, step=1000.0)
    incentives = st.number_input("Lease incentives received (₹)", min_value=0.0, value=0.0, step=1000.0)

    useful_life_years = st.number_input(
        "Useful life of asset (years) — leave 0 to use lease term",
        min_value=0.0, value=0.0, step=0.5
    )
    useful_life_months = int(round(useful_life_years * 12)) if useful_life_years > 0 else None

    st.header("4. Interest-Free Security Deposit")
    has_deposit = st.checkbox("Include refundable interest-free security deposit", value=True)
    deposit_amount = 0.0
    deposit_term_years = 0.0
    market_rate = discount_rate
    fair_value_deposit = standard == "Ind AS 116"
    if has_deposit:
        deposit_amount = st.number_input("Security deposit amount (₹)", min_value=0.0, value=600000.0, step=10000.0)
        deposit_term_years = st.number_input("Deposit refund term (years)", min_value=0.0, value=term_years, step=0.5)
        market_rate = st.number_input("Market rate of interest for discounting deposit (%)", min_value=0.0,
                                       value=discount_rate, step=0.25)
        fair_value_deposit = st.checkbox(
            "Recognise deposit at present value (fair valuation, Ind AS 109 approach)",
            value=(standard == "Ind AS 116"),
            help="Mandatory in substance under Ind AS 116/109. Optional / not typically done under old AS 19 framework."
        )

# ----------------------------------------------------------------------
# Calculations
# ----------------------------------------------------------------------

payments = build_payment_schedule(base_rent, term_months, escalation_pct, escalation_freq_months)
monthly_rate = annual_to_monthly_rate(discount_rate)

liability_df, initial_liability = lease_liability_schedule(payments, monthly_rate, timing, start_date)

deposit_df = pd.DataFrame()
pv_deposit = deposit_amount
day1_diff = 0.0
if has_deposit and fair_value_deposit and deposit_amount > 0 and deposit_term_years > 0:
    deposit_df, pv_deposit, day1_diff = deposit_unwinding_schedule(
        deposit_amount, market_rate, int(round(deposit_term_years * 12)), start_date
    )

rou_initial = initial_liability + idc - incentives + day1_diff
rou_df = rou_depreciation_schedule(rou_initial, term_months, start_date, useful_life_months)

asset_label = "ROU Asset" if standard == "Ind AS 116" else "Leased Asset"

# ----------------------------------------------------------------------
# Summary
# ----------------------------------------------------------------------

st.subheader("Summary")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Initial Lease Liability", f"₹{initial_liability:,.0f}")
c2.metric(f"Initial {asset_label} Value", f"₹{rou_initial:,.0f}")
c3.metric("Total Finance Cost (life of lease)", f"₹{liability_df['Interest / Finance Cost'].sum():,.0f}")
if not deposit_df.empty:
    c4.metric("Total Unwinding Income (deposit)", f"₹{deposit_df['Unwinding / Notional Interest Income'].sum():,.0f}")
else:
    c4.metric("Total Unwinding Income (deposit)", "N/A")

if has_deposit and fair_value_deposit:
    st.info(
        f"Security deposit of ₹{deposit_amount:,.0f} recognised at present value ₹{pv_deposit:,.0f}. "
        f"Day-1 difference of ₹{day1_diff:,.0f} added to {asset_label} as prepaid rent, "
        f"to be recovered over the lease term through depreciation, while notional interest income "
        f"unwinds the deposit back to ₹{deposit_amount:,.0f} by the end of its term."
    )

st.divider()

# ----------------------------------------------------------------------
# Tabs
# ----------------------------------------------------------------------

tab1, tab2, tab3, tab4 = st.tabs([
    "📄 Lease Liability & Finance Cost",
    f"🏢 {asset_label} Depreciation",
    "💰 Security Deposit Unwinding",
    "⬇️ Export",
])

with tab1:
    st.markdown(f"**Discount rate used:** {discount_rate}% p.a. ({monthly_rate*100:.4f}% monthly)")
    st.dataframe(liability_df, use_container_width=True, hide_index=True)
    st.line_chart(liability_df.set_index("Date")[["Opening Liability", "Closing Liability"]])
    yearly = liability_df.copy()
    yearly["Year"] = ((yearly["Month"] - 1) // 12) + 1
    yearly_summary = yearly.groupby("Year").agg(
        {"Interest / Finance Cost": "sum", "Lease Payment": "sum", "Closing Liability": "last"}
    ).reset_index()
    st.markdown("**Year-wise Summary**")
    st.dataframe(yearly_summary, use_container_width=True, hide_index=True)

with tab2:
    st.markdown(f"**Initial {asset_label} value:** ₹{rou_initial:,.2f} "
                f"(Lease liability + IDC − incentives + deposit day-1 difference)")
    st.dataframe(rou_df, use_container_width=True, hide_index=True)
    st.line_chart(rou_df.set_index("Date")[["Opening WDV", "Closing WDV"]])

with tab3:
    if deposit_df.empty:
        st.warning("No security deposit unwinding schedule generated. "
                    "Enable the deposit and fair-valuation option in the sidebar to see this schedule.")
    else:
        st.dataframe(deposit_df, use_container_width=True, hide_index=True)
        st.line_chart(deposit_df.set_index("Date")[
            ["Opening Deposit (Amortised Cost)", "Closing Deposit (Amortised Cost)"]
        ])

with tab4:
    st.markdown("Download all schedules as a single Excel workbook.")
    sheets = {
        "Lease Liability": liability_df,
        f"{asset_label} Depreciation": rou_df,
    }
    if not deposit_df.empty:
        sheets["Deposit Unwinding"] = deposit_df
    excel_bytes = to_excel(sheets)
    st.download_button(
        "Download Excel Workbook",
        data=excel_bytes,
        file_name="finance_lease_schedules.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.divider()
st.caption(
    "Note: This tool provides an indicative computation for management/working purposes. "
    "Lease classification (finance vs operating under AS 19) and other judgemental aspects "
    "(e.g. discount rate determination, lease term including renewal/termination options) "
    "should be independently assessed as per the applicable standard."
)
