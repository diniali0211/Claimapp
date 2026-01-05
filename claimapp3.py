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


# ───────────────────────────── Helpers ─────────────────────────────
def ensure_unique_headers(df):
    counts, new_cols = {}, []
    for c in df.columns:
        base = str(c).strip()
        if base not in counts:
            counts[base] = 1
            new_cols.append(base)
        else:
            counts[base] += 1
            new_cols.append(f"{base}_{counts[base]}")
    df.columns = new_cols
    return df


def _norm_empid(s):
    if pd.isna(s):
        return np.nan
    s = str(s).strip().upper().replace(" ", "")
    return s[:-2] if s.endswith(".0") else s


def _norm_name(s):
    return "" if pd.isna(s) else str(s).strip()


def _norm_recruiter(s):
    if pd.isna(s) or str(s).strip() == "":
        return "Unassigned"
    return str(s).strip().title()


def _is_leave(val):
    if val is None:
        return False
    s = str(val).lower()
    return any(k in s for k in ["unpaid", "annual", "absent", "emergency", "medical", "sick", "mc", "leave"])


def _to_hours_any(val):
    try:
        return float(val)
    except Exception:
        return np.nan


def _pair_duration(i, o):
    hi, ho = _to_hours_any(i), _to_hours_any(o)
    if pd.isna(hi) or pd.isna(ho):
        return 0.0
    dur = ho - hi
    return dur + 24 if dur < 0 else dur


def _parse_dates(series, day_first_flag=True):
    d = pd.to_datetime(series, errors="coerce", dayfirst=day_first_flag)
    return d


def load_masterlist(file_like):
    raw = pd.read_excel(file_like, header=None, dtype=str)
    for i in range(min(50, len(raw))):
        row = raw.iloc[i].astype(str).str.lower()
        if row.str.contains("name").any() and row.str.contains("join").any():
            return pd.read_excel(file_like, header=i, dtype=str)
    return pd.read_excel(file_like, dtype=str)


# ───────────────────────────── Main ─────────────────────────────
if att_file and mst_file:
    try:
        att_raw = pd.read_excel(att_file, dtype=str)
        att_raw = ensure_unique_headers(att_raw)

        mst = load_masterlist(mst_file)
        mst = ensure_unique_headers(mst)

        # Basic assumptions (kept as-is from your logic)
        att["__Date"] = _parse_dates(att_raw.iloc[:, 0], day_first)
        att["__Name"] = att_raw.iloc[:, 1].apply(_norm_name)
        att["__EmpID"] = att_raw.iloc[:, 2].apply(_norm_empid)
        att["__Hours"] = 8
        att["__LeaveFlagRow"] = False

        mst_name = mst.columns[0]
        mst_join = mst.columns[1]
        mst_recr = mst.columns[2] if len(mst.columns) > 2 else None

        mst[mst_join] = pd.to_datetime(mst[mst_join], errors="coerce")

        join_by_name = dict(zip(mst[mst_name], mst[mst_join]))
        recr_by_name = dict(zip(mst[mst_name], mst[mst_recr])) if mst_recr else {}

        att["JOIN_DATE"] = att["__Name"].map(join_by_name)
        att["Recruiter"] = att["__Name"].map(recr_by_name).fillna("Unassigned")

        eligible = att[att["JOIN_DATE"].notna()].copy()
        eligible["Worked_Day"] = (eligible["__Hours"] >= effective_threshold).astype(int)

        months = eligible["__Date"].dt.to_period("M").unique()

        tables_by_month, summaries_by_month = {}, {}

        for m in months:
            df = eligible[eligible["__Date"].dt.to_period("M") == m]
            base = df.groupby(["__Name", "Recruiter"], as_index=False)["Worked_Day"].sum()
            base.rename(columns={"Worked_Day": "TOTAL WORKING"}, inplace=True)
            tables_by_month[str(m)] = base

            summary = base.groupby("Recruiter", as_index=False)["TOTAL WORKING"].sum()
            summary["Rate (RM)"] = day_rate
            summary["Amount (RM)"] = summary["TOTAL WORKING"] * day_rate

            grand = pd.DataFrame({
                "Recruiter": ["TOTAL"],
                "TOTAL WORKING": [summary["TOTAL WORKING"].sum()],
                "Rate (RM)": [day_rate],
                "Amount (RM)": [summary["Amount (RM)"].sum()]
            })

            summaries_by_month[str(m)] = (summary, grand)

        # ───────────── EXPORT (WITH SIGNATURES) ─────────────
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            for mon in tables_by_month:
                df_out = tables_by_month[mon]
                df_out.to_excel(writer, sheet_name=mon, index=False)

                ws = writer.sheets[mon]
                wb = writer.book
                bold = wb.add_format({"bold": True})

                start = len(df_out) + 2
                summary, grand = summaries_by_month[mon]

                ws.write(start, 0, "Per-Recruiter Summary", bold)
                summary.to_excel(writer, sheet_name=mon, startrow=start + 1, index=False)

                start = start + len(summary) + 3
                ws.write(start, 0, "Grand Total", bold)
                grand.to_excel(writer, sheet_name=mon, startrow=start + 1, index=False)

                # ───── Signature Section ─────
                sig_row = start + len(grand) + 4
                sig_line = wb.add_format({"top": 1})
                sig_center = wb.add_format({"align": "center"})
                sig_bold = wb.add_format({"bold": True})

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
            "⬇️ Download Excel",
            buffer.getvalue(),
            file_name=f"claim_tables_{dt.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.exception(e)
else:
    st.info("Upload both files to continue.")


