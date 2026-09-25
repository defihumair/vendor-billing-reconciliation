from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

import generate_data as gd
import reconcile as rc

BASE = Path(__file__).parent
SAMPLE = BASE / "sample_data"
SAMPLE_FILES = ["internal_volumes.csv", "vendor_invoice.csv", "rate_card.csv"]

st.set_page_config(page_title="Vendor Billing Reconciliation", page_icon="🧾", layout="wide")


@st.cache_data
def sample_inputs():
    paths = [SAMPLE / f for f in SAMPLE_FILES]
    if all(p.exists() for p in paths):
        return tuple(pd.read_csv(p) for p in paths)
    return gd.build()


def money(x, currency):
    return f"{currency} {x:,.0f}"


st.title("Vendor billing reconciliation")
st.caption(
    "Reconciles a vendor's weekly invoice against internal activity volumes, applies billing rules, "
    "requires manager approval, then generates the payable invoice. All data in the demo is synthetic."
)

with st.sidebar:
    st.header("Data")
    source = st.radio("Data source", ["Sample data", "Upload CSV files"])
    uploads = None
    if source == "Upload CSV files":
        f_int = st.file_uploader("Internal volumes", type="csv", help="Columns: year, week, facility, billing_head, qty")
        f_ven = st.file_uploader("Vendor invoice", type="csv", help="Columns: year, week, invoice_no, facility, billing_head, qty, amount")
        f_rate = st.file_uploader("Rate card", type="csv", help="Columns: billing_head, uom, rate")
        if f_int and f_ven and f_rate:
            uploads = (pd.read_csv(f_int), pd.read_csv(f_ven), pd.read_csv(f_rate))
    st.header("Invoice")
    vendor_name = st.text_input("Vendor", "Vendor A")
    bill_to = st.text_input("Bill to", "Demo Logistics Co.")
    currency = st.text_input("Currency", "PKR")

if source == "Upload CSV files" and uploads is None:
    st.info("Upload all three CSV files in the sidebar, or switch to Sample data to try the demo.")
    st.stop()

try:
    internal, vendor, rates = rc.validate(*(uploads if uploads else sample_inputs()))
except ValueError as err:
    st.error(str(err))
    st.stop()

tab_rec, tab_log, tab_how = st.tabs(["Reconcile a week", "Savings log", "How it works"])

with tab_rec:
    options = rc.periods(vendor)
    c1, c2, c3 = st.columns(3)
    period = c1.selectbox("Billing week", options, format_func=lambda p: f"{p[0]}, week {p[1]}")
    year, week = period
    invoice_date = c2.date_input("Invoice date", date.today())
    invoice_no = c3.text_input("Invoice number", rc.invoice_number(vendor, year, week), key=f"inv_{year}_{week}")

    lines = rc.reconcile(internal, vendor, rates, year, week)
    exceptions = int((lines["Status"] != "Match").sum())

    st.subheader("1. Review every line")
    st.caption(
        f"{len(lines)} lines, {exceptions} need attention. Exceptions are listed first. "
        "Change the Override column only where the default rule does not fit."
    )
    view = ["Line", "Billing head", "Facility", "UOM", "Rate", "Internal qty", "Vendor qty",
            "Qty variance", "Vendor amount", "Status", "Override", "Recommendation"]
    edited = st.data_editor(
        lines[view],
        hide_index=True,
        key=f"editor_{year}_{week}",
        disabled=[c for c in view if c != "Override"],
        column_config={
            "Rate": st.column_config.NumberColumn(format="%.2f"),
            "Internal qty": st.column_config.NumberColumn(format="%.0f"),
            "Vendor qty": st.column_config.NumberColumn(format="%.0f"),
            "Qty variance": st.column_config.NumberColumn(format="%.0f"),
            "Vendor amount": st.column_config.NumberColumn(format="%.2f"),
            "Recommendation": st.column_config.TextColumn(width="medium"),
            "Override": st.column_config.SelectboxColumn(options=rc.OVERRIDES, required=True, width="medium"),
        },
    )
    final = rc.apply_overrides(lines, edited["Override"])
    s = rc.summarize(final)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Internal value", money(s["internal_value"], currency))
    m2.metric("Vendor invoiced", money(s["vendor_total"], currency))
    m3.metric("Approved payable", money(s["payable"], currency), help="Before tax")
    m4.metric("Correction", money(s["correction"], currency), help="Vendor invoiced minus approved payable")

    with st.expander("Approved amounts by line"):
        st.dataframe(
            final[["Line", "Billing head", "Facility", "Override", "Approved qty", "Rate", "Payable", "Vendor amount", "Correction"]],
            hide_index=True,
            column_config={c: st.column_config.NumberColumn(format="%.2f") for c in ["Rate", "Payable", "Vendor amount", "Correction"]},
        )

    st.subheader("2. Approve and generate")
    a1, a2 = st.columns([1, 2])
    approver = a1.text_input("Approver name", key=f"approver_{year}_{week}")
    approved = a2.checkbox(
        "I have reviewed every line and approve this payable amount.", key=f"approve_{year}_{week}"
    )
    if s["overrides"]:
        st.caption(f"{s['overrides']} manager override(s) will be recorded in the audit trail.")
    ready = approved and approver.strip() != ""
    if st.button("Approve and generate payable invoice", type="primary", disabled=not ready):
        meta = rc.invoice_meta(vendor_name, bill_to, invoice_no, invoice_date, year, week, currency, approver.strip())
        st.session_state["invoice"] = (period, rc.build_invoice(final, meta), f"Payable_Invoice_{year}_W{week:02d}.xlsx")
    if not ready:
        st.caption("Enter your name and tick the approval box to generate the invoice.")
    saved = st.session_state.get("invoice")
    if saved and saved[0] == period:
        st.success("Payable invoice generated with manager overrides applied.")
        st.download_button(
            "Download payable invoice (.xlsx)", saved[1], saved[2],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with tab_log:
    log = rc.savings_log(internal, vendor, rates)
    total = float(log["Correction"].sum())
    per_week = total / len(log) if len(log) else 0.0
    k1, k2, k3 = st.columns(3)
    k1.metric("Total corrected", money(total, currency))
    k2.metric("Average per week", money(per_week, currency))
    k3.metric("Estimated per month", money(per_week * 52 / 12, currency))
    st.bar_chart(log.assign(Period=log["Year"].astype(str) + "-W" + log["Week"].astype(str).str.zfill(2)).set_index("Period")["Correction"])
    st.dataframe(
        log, hide_index=True,
        column_config={c: st.column_config.NumberColumn(format="%.2f") for c in ["Vendor invoiced", "Approved payable", "Correction"]},
    )
    st.caption("Calculated with the system default rules for every week. Manager overrides on individual weeks can change these figures.")
    st.download_button("Download savings log (.csv)", log.to_csv(index=False).encode(), "savings_log.csv", mime="text/csv")

with tab_how:
    st.markdown(
        """
**Flow:** source files → consolidated internal volumes → vendor invoice → line-by-line reconciliation → manager review → payable invoice.

| Situation | Default action | Why |
|---|---|---|
| Vendor quantity above internal records | Pay the internal quantity | The vendor billed more work than was recorded |
| Vendor quantity below internal records | Pay the vendor's quantity | The vendor billed less; pay what was invoiced |
| Quantities match | Pay as invoiced | No difference |
| Vendor amount ≠ quantity × contract rate | Contract rate applied | Protects against rate drift |
| Manager override | Force internal or vendor quantity for that line | Handles cases the rules cannot see |

Payable amount = approved quantity × contract rate. No invoice is generated until a named approver confirms the review, and every override is written to the audit trail sheet of the invoice file.
"""
    )
