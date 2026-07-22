import streamlit as st
import pandas as pd
import numpy as np
import os
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Insurance Premium Calculator",
    page_icon="💰",
    layout="wide"
)

st.title("💰 Aviva GCL Insurance Premium Calculator")
st.markdown("Select plan details below")

FILE_MAP = {
    ("Level",    "Single Life", "Home Loan"): "Homeloan Single Life.xlsx",
    ("Level",    "Single Life", "LAP"):       "LAP Single Life.xlsx",
    ("Level",    "Joint Life",  "Home Loan"): "Homeloan Joint Life.xlsx",
    ("Level",    "Joint Life",  "LAP"):       "LAP Joint Life.xlsx",
    ("Reducing", "Joint Life",  "Home Loan"): "Reducing- Homeloan.xlsx",
    ("Reducing", "Joint Life",  "LAP"):       "Reducing- LAP.xlsx",
}

# GST is fixed and always applied on top of the Loader-adjusted rate.
GST_RATE_FIXED = 18.0

# ============================================
# LOADER + GST FORMULA (applied to every rate before premium is computed):
#   Rate After Loader = Base Rate / (1 - Loader% / 100)
#   Final Rate         = Rate After Loader x (1 + GST% / 100)
# ============================================
def apply_loader_and_gst(base_rate, loader_pct, gst_pct=GST_RATE_FIXED):
    after_loader = base_rate / (1 - (loader_pct / 100.0))
    final_rate = after_loader * (1 + (gst_pct / 100.0))
    return final_rate


def load_rate_table(cover_type, life_type, loan_type):
    fname = FILE_MAP[(cover_type, life_type, loan_type)]
    if not os.path.exists(fname):
        raise FileNotFoundError(
            f"File not found: '{fname}' — Please make sure this file is in the GitHub repo."
        )

    raw = pd.read_excel(fname, sheet_name="Sheet1", header=None)

    header_row = None
    for i, row in raw.iterrows():
        for val in row.values:
            if isinstance(val, str) and "AGE" in val.upper():
                header_row = i
                break
        if header_row is not None:
            break

    if header_row is None:
        raise ValueError("Could not find AGE/TERM header row in the file.")

    df = pd.read_excel(fname, sheet_name="Sheet1", header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    age_col = df.columns[0]
    df = df.dropna(subset=[age_col])
    df[age_col] = pd.to_numeric(df[age_col], errors='coerce')
    df = df.dropna(subset=[age_col])
    df[age_col] = df[age_col].astype(int)
    df = df.set_index(age_col)

    tenure_map = {}
    for col in df.columns:
        try:
            tenure_map[int(float(col))] = col
        except Exception:
            pass

    return df, tenure_map


def get_rate(df, tenure_map, age, tenure):
    if age not in df.index:
        raise ValueError(f"Age {age} not found in rate table.")
    if tenure not in tenure_map:
        raise ValueError(f"Tenure {tenure} yrs not found in rate table.")
    return float(df.loc[age, tenure_map[tenure]])


def find_column(df, target):
    """Exact (case/space-insensitive) match."""
    target_norm = target.strip().lower().replace(" ", "")
    for col in df.columns:
        col_norm = str(col).strip().lower().replace(" ", "")
        if col_norm == target_norm:
            return col
    return None


def find_sum_assured_column(df):
    """
    Flexibly detect a Sum Assured / Sum Insured column first — this takes
    priority whenever present. Falls back to Loan Outstanding only if no
    Sum Assured-type column is found.
    """
    for col in df.columns:
        norm = re.sub(r'[\s_\-/]+', '', str(col).lower())
        if 'sumassured' in norm or 'suminsured' in norm:
            return col
    for col in df.columns:
        norm = re.sub(r'[\s_\-/]+', '', str(col).lower())
        if ('loanoutstanding' in norm or 'outstandingamount' in norm or 'outstandingloan' in norm
                or norm == 'outstanding' or 'loanos' in norm or norm == 'os'):
            return col
    return None


# ============================================
# DROPDOWNS
# ============================================

col1, col2 = st.columns(2)
with col1:
    life_type = st.selectbox("Select Life Type", ["Single Life", "Joint Life"])
with col2:
    loan_type = st.selectbox("Select Loan Type", ["Home Loan", "LAP"])

if life_type == "Joint Life":
    cover_type = st.selectbox("Select Type of Cover", ["Level", "Reducing"])
else:
    cover_type = "Level"

# ============================================
# SHARED LOADER % — applied to every rate (Manual + Bulk)
# before computing premium. GST @ 18% is then added automatically on top.
# ============================================
st.subheader("⚙️ Loader Setting")
loader_pct_input = st.number_input(
    "Loader % (optional — applied to all rate lookups below; GST @ 18% is then added automatically)",
    min_value=0.0,
    max_value=99.99,
    value=None,
    step=1.0,
    placeholder="Enter loader % (optional, defaults to 0)",
    key="shared_loader_pct"
)
loader_pct = loader_pct_input if loader_pct_input is not None else 0.0

# ============================================
# SUM ASSURED RANGE — rates in the backend files are per ₹1,00,000
# ============================================
if loan_type == "Home Loan":
    sa_min, sa_max = 100000, 6000000
else:
    sa_min, sa_max = 100000, 4000000

st.divider()

# ============================================
# MANUAL SECTION
# ============================================

st.subheader("🔢 Manual Rate Lookup")

if loan_type == "Home Loan":
    min_tenure, max_tenure = 5, 25
else:
    min_tenure, max_tenure = 2, 10

col3, col4 = st.columns(2)
with col3:
    age = st.number_input("Enter Age", min_value=18, max_value=65, value=30, step=1)
with col4:
    tenure = st.number_input(
        "Enter Tenure",
        min_value=min_tenure,
        max_value=max_tenure,
        value=min_tenure,
        step=1
    )
    st.caption("📅 Tenure is in Years")

sum_assured_manual = st.number_input(
    "Select Sum Assured (₹)",
    min_value=0,
    value=sa_min,
    step=1,
    help="Enter the exact Sum Assured for this member.",
    key="sa_manual"
)

if st.button("Get Rate", type="primary"):
    try:
        df_rates, tenure_map = load_rate_table(cover_type, life_type, loan_type)
        base_rate = get_rate(df_rates, tenure_map, age, tenure)
        final_rate = apply_loader_and_gst(base_rate, loader_pct)
        premium = final_rate * (sum_assured_manual / 100000)
        st.success(
            f"✅ {life_type} | {loan_type} | {cover_type} Cover | Age {age} | Tenure {tenure} yrs | "
            f"Sum Assured ₹{sum_assured_manual:,} | Loader {loader_pct}% | GST {GST_RATE_FIXED}%"
        )
        st.metric("Premium", f"₹ {premium:,.2f}")
    except Exception as e:
        st.error(f"Error: {e}")

st.divider()

# ============================================
# EXCEL UPLOAD SECTION
# ============================================

st.subheader("📂 Upload Member Data for Bulk Rate Lookup")

st.markdown(
    "Your Excel must have at least: **Name**, **Age**, **Sum Assured** (or **Loan "
    "Outstanding** if Sum Assured isn't present), and a **Tenure** column (in years)."
)

st.caption(f"Each row uses its own Sum Assured from the Excel (must be between ₹{sa_min:,} and ₹{sa_max:,}).")
st.warning("⚠️ Please make sure you have selected **Life Type**, **Loan Type**, and **Type of Cover** above before uploading your Excel file.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]

        st.subheader("Uploaded Data Preview")
        st.dataframe(df.head())

        if loan_type == "Home Loan":
            min_t, max_t = 5, 25
        else:
            min_t, max_t = 2, 10

        df_rates, tenure_map = load_rate_table(cover_type, life_type, loan_type)

        name_col = find_column(df, "Name")
        age_col = find_column(df, "Age")
        tenure_col = find_column(df, "Tenure")
        sa_col = find_sum_assured_column(df)

        if not name_col or not age_col or not tenure_col:
            raise ValueError(
                "Excel must contain mandatory columns: Name, Age, and Tenure."
            )
        if not sa_col:
            raise ValueError("Excel must contain a Sum Assured column (e.g. 'Sum Assured', 'Sum Insured') or, failing that, a 'Loan Outstanding' column.")

        df[age_col] = pd.to_numeric(df[age_col], errors='coerce')
        df[tenure_col] = pd.to_numeric(df[tenure_col], errors='coerce')
        df[sa_col] = pd.to_numeric(df[sa_col], errors='coerce')

        if df[tenure_col].dropna().median() > 30:
            st.info("ℹ️ Tenure values look like months — auto-converting to years.")
            df[tenure_col] = (df[tenure_col] / 12).round(0).astype('Int64')
        else:
            df[tenure_col] = df[tenure_col].round(0).astype('Int64')

        df[age_col] = df[age_col].round(0).astype('Int64')

        premiums = []
        statuses = []
        for idx, row in df.iterrows():
            try:
                r_age = int(row[age_col]) if pd.notna(row[age_col]) else None
                r_tenure = int(row[tenure_col]) if pd.notna(row[tenure_col]) else None
                r_sa = float(row[sa_col])

                if r_age is None or r_age < 18 or r_age > 65:
                    raise ValueError("Age must be between 18 and 65")
                if r_tenure is None or r_tenure < min_t or r_tenure > max_t:
                    raise ValueError(f"Tenure must be between {min_t} and {max_t} yrs")

                r_base = get_rate(df_rates, tenure_map, r_age, r_tenure)
                r_final = apply_loader_and_gst(r_base, loader_pct)
                premium = round(r_final * (r_sa / 100000), 2)
                premiums.append(premium)
                statuses.append("✅")
            except Exception as e:
                premiums.append(None)
                statuses.append(f"❌ {e}")

        df["Premium"] = premiums
        df["Status"] = statuses

        core_cols = [name_col, age_col, tenure_col, sa_col, "Premium"]
        extra_cols = [c for c in df.columns if c not in core_cols]
        df = df[core_cols + extra_cols]

        total_premium = pd.to_numeric(pd.Series(premiums), errors='coerce').sum()
        st.metric("💰 Grand Total Premium", f"₹ {total_premium:,.2f}")

        st.subheader("Rate Lookup Output")
        st.dataframe(df, use_container_width=True)

        total_row = {c: "" for c in df.columns}
        total_row[name_col] = "TOTAL PREMIUM"
        total_row["Premium"] = round(total_premium, 2)
        df_out = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

        output_file = "Rate_Output.xlsx"
        df_out.to_excel(output_file, index=False)

        with open(output_file, "rb") as file:
            st.download_button(
                label="⬇ Download Output Excel",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
