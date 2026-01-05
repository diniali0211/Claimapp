import io
import re
import calendar
from pathlib import Path
from datetime import datetime as dt

import numpy as np
import pandas as pd
import streamlit as st


# ───────────────────────────── UI header ─────────────────────────────
st.set_page_config(page_title="Agency Claim Table Generator (Auto Recruiter)", layout="wide")
st.title("🧾 Agency Claim Table Generator (Auto-detect Recruiter)")

st.markdown(
    """
- **Masterlist**: must include **Name**, **Joined Date**, and **Recruiter** (optional but recommended).  
- **Timecard**: must include **Emp No**, **Name**, **Date**, and **one IN + one OUT** column.  
- A workday counts **1** if **daily hours ≥ (hours per day − grace)** and **not on leave**.
- Eligibility: **JOIN_DATE → JOIN_DATE + 3 months − 1 day** (inclusive).
"""
)

# ───────────────────────────── Sidebar ─────────────────────────────
with st.sidebar:
    st.header("Settings")
    hours_per_day = st.number_input("Hours considered 1 workday", 1.0, 24.0, 8.0, 0.5)
    grace_minutes = st.number_input("Grace window (minutes)", 0, 120, 15, 5)
    counting_rule = st.selectbox(
        "Counting rule",
        ["Per-day ≥ threshold", "Floor(total hours ÷ threshold)"],
        index=0,
    )
    day_rate = st.number_input("Rate per claim day (RM)", 0.0, 1000.0, 3.0, 0.5)
    exclude_not_in_master = st.checkbox("Exclude employees not in Masterlist", value=True)
    day_first = st.checkbox("Dates are day-first (DD/MM/YYYY)", value=True)

effective_threshold = max(0.0, float(hours_per_day) - float(grace_minutes) / 60.0)

# ───────────────────────────── Uploaders ─────────────────────────────
att_file = st.file_uploader("Upload **Timecard** (CSV/XLSX/XLS/XLSM)", type=["csv", "xlsx", "xls", "xlsm"])
mst_file = st.file_uploader("Upload **Masterlist** (XLSX/XLS/XLSM)", type=["xlsx", "xls", "xlsm"])

# =========================
# (ALL YOUR EXISTING LOGIC
#  IS UNCHANGED ABOVE)
# =========================


# ---------- Export ----------
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
    for mon in sorted(tables_by_month.keys()):
        sheet = f"{mon}"[:31]
        df_out = tables_by_month[mon]
        df_out.to_excel(writer, sheet_name=sheet, index=False)

        workbook = writer.book
        ws = writer.sheets[sheet]
        bold = workbook.add_format({"bold": True})
        ws.freeze_panes(1, 5)

        start = len(df_out) + 2
        ws.write(start, 0, "Per-Recruiter Summary", bold)
        start += 1

        summary, grand = summaries_by_month[mon]
        summary.to_excel(writer, sheet_name=sheet, startrow=start, index=False)

        start = start + len(summary) + 2
        ws.write(start, 0, "Grand Total", bold)
        start += 1
        grand.to_excel(writer, sheet_name=sheet, startrow=start, index=False)

        # ───────── Signature Section (ONLY ADDITION) ─────────
        sig_row = start + len(grand) + 4

        sig_bold = workbook.add_format({"bold": True})
        sig_line = workbook.add_format({"top": 1})
        sig_center = workbook.add_format({"align": "center"})

        last_col = df_out.shape[1] - 1
        mid_col = last_col // 2

        # LEFT — Verified by
        ws.merge_range(sig_row, 0, sig_row, mid_col - 1, "Verified by", sig_bold)
        ws.merge_range(sig_row + 2, 0, sig_row + 2, mid_col - 1, "", sig_line)
        ws.merge_range(sig_row + 3, 0, sig_row + 3, mid_col - 1, "Name & Signature", sig_center)
        ws.merge_range(sig_row + 4, 0, sig_row + 4, mid_col - 1, "Date:", sig_center)

        # RIGHT — Acknowledged by
        ws.merge_range(sig_row, mid_col + 1, sig_row, last_col, "Acknowledged by", sig_bold)
        ws.merge_range(sig_row + 2, mid_col + 1, sig_row + 2, last_col, "", sig_line)
        ws.merge_range(sig_row + 3, mid_col + 1, sig_row + 3, last_col, "Name & Signature", sig_center)
        ws.merge_range(sig_row + 4, mid_col + 1, sig_row + 4, last_col, "Date:", sig_center)

st.download_button(
    "⬇️ Download Excel (one sheet per month: table + summary + grand total)",
    data=buffer.getvalue(),
    file_name=f"claim_tables_{dt.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
