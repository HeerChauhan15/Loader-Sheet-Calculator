import streamlit as st
import pandas as pd
import numpy as np
import os
import re

st.set_page_config(
    page_title="Loader Sheet Calculator",
    page_icon="💰",
    layout="wide"
)

st.title("💰 Loader Sheet Calculator")
st.markdown("Select Type of Life, Segment, and Loading % below")

GST_RATE = 0.18

# ============================================
# AGE LIMITS (ALL SEGMENTS / LIFE TYPES)
# ============================================
AGE_MIN = 18
AGE_MAX = 60

# ============================================
# LOADING % LIMITS
# ============================================
LOADING_MIN = 0
LOADING_MAX = 500  # generous upper bound, adjust if needed

# ============================================
# FILE MAP: (Segment, Type of Life) -> filename
# EDIT THESE FILENAMES to match your actual uploaded files exactly.
# ============================================
FILE_MAP = {
    ("Home Loan", "Single"): "Homeloan Single Life.xlsx",
    ("Home Loan", "Joint"): "Homeloan Joint Life.xlsx",
    ("Loan Against Property", "Single"): "LAP Single Life.xlsx",
    ("Loan Against Property", "Joint"): "LAP Joint Life.xlsx",
}

# ============================================
# SUM ASSURED & TENURE LIMITS PER SEGMENT
# (same limits used for Single and Joint life)
# ============================================
SEGMENT_LIMITS = {
    "Home Loan": {"sa_min": 500000, "sa_max": 6000000, "t_min": 5, "t_max": 25},
    "Loan Against Property": {"sa_min": 100000, "sa_max": 4000000, "t_min": 1, "t_max": 10},
}

SEGMENT_OPTIONS = list(SEGMENT_LIMITS.keys())
LIFE_OPTIONS = ["Single", "Joint"]

# ============================================
# RATE TABLE LOADER (sheet header row detected via "AGE" text, same as reference)
# ============================================
def load_rate_table(segment, life_type):
    fname = FILE_MAP[(segment, life_type)]
    if not os.path.exists(fname):
        raise FileNotFoundError(
            f"File not found: '{fname}' — Please make sure this file is in the GitHub repo "
            f"with this exact name (or update FILE_MAP in the code)."
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


def get_base_rate(df, tenure_map, age, tenure_years):
    """Tenure is in YEARS in both the user input and the rate table's columns —
    no conversion needed."""
    if age not in df.index:
        raise ValueError(f"Age {age} not found in rate table.")
    tenure_key = int(round(tenure_years))
    if tenure_key not in tenure_map:
        raise ValueError(
            f"Tenure {tenure_years} yrs not found in rate table."
        )
    return float(df.loc[age, tenure_map[tenure_key]])


def apply_loading(base_rate, loading_pct):
    """Loading % is applied on the sheet rate. GST is already included in the
    sheet's rate, so this loaded value IS the final rate — no further math needed."""
    return base_rate * (1 + (loading_pct / 100.0))


# ============================================
# DROPDOWNS (Type of Life, Segment) + LOADING %
# ============================================
col1, col2, col3 = st.columns(3)
with col1:
    life_type = st.selectbox("Type of Life", LIFE_OPTIONS)
with col2:
    segment = st.selectbox("Segment", SEGMENT_OPTIONS)
with col3:
    loading_pct = st.number_input(
        "Loading % (Header Loader)",
        min_value=float(LOADING_MIN),
        max_value=float(LOADING_MAX),
        value=0.0,
        step=1.0,
        help="Enter the loading percentage to apply on top of the base rate before GST."
    )

limits = SEGMENT_LIMITS[segment]
min_tenure, max_tenure = limits["t_min"], limits["t_max"]

st.divider()

# ============================================
# MANUAL SECTION
# ============================================
col4, col5 = st.columns(2)
with col4:
    age_input = st.number_input(
        "Enter Age",
        min_value=18,
        value=30,
        step=1
    )

    if age_input < AGE_MIN:
        st.warning(
            f"⚠️ Minimum Age is **{AGE_MIN} years**. Value adjusted to **{AGE_MIN} years**."
        )
        age = AGE_MIN
    elif age_input > AGE_MAX:
        st.warning(
            f"⚠️ Maximum Age is **{AGE_MAX} years**. Value adjusted to **{AGE_MAX} years**."
        )
        age = AGE_MAX
    else:
        age = age_input

with col5:
    tenure_input = st.number_input(
        "Enter Tenure",
        min_value=0,
        value=min_tenure,
        step=1
    )
    if tenure_input < min_tenure:
        st.warning(
            f"⚠️ Minimum Tenure for {segment} is {min_tenure} yrs. "
            f"Value adjusted to {min_tenure} yrs."
        )
        tenure = min_tenure
    elif tenure_input > max_tenure:
        st.warning(
            f"⚠️ Maximum Tenure for {segment} is {max_tenure} yrs. "
            f"Value adjusted to {max_tenure} yrs."
        )
        tenure = max_tenure
    else:
        tenure = tenure_input
    st.caption("📅 Tenure is in Years")

st.write("")
if st.button("Get Rate", type="primary", use_container_width=True):
    try:
        df_rates, tenure_map = load_rate_table(segment, life_type)
        base_rate = get_base_rate(df_rates, tenure_map, age, tenure)
        rate = apply_loading(base_rate, loading_pct)

        st.success(
            f"✅ {segment} | {life_type} Life | Age {age} | Tenure {tenure} yrs | "
            f"Loading {loading_pct}%"
        )

        st.metric("Rate", f"{rate:,.4f}")
        st.caption("Rate is per ₹1,00,000 Sum Assured, GST already included.")
    except Exception as e:
        st.error(f"Error: {e}")

st.divider()

# ============================================
# FULL LOADED RATE TABLE — GENERATE & DOWNLOAD
# (No upload needed — built entirely from the selections above)
# ============================================
st.subheader("📊 Generate Full Loaded Rate Table")

st.markdown(
    f"This builds the complete rate table (every Age × Tenure combination) for "
    f"**{segment} | {life_type} Life | Loading {loading_pct}%**, with GST applied, "
    f"ready to download."
)

if st.button("Generate Rate Table", type="primary", use_container_width=True):
    try:
        df_rates, tenure_map = load_rate_table(segment, life_type)

        # Ages: only those present in the sheet AND within the allowed 18-60 range
        valid_ages = sorted(a for a in df_rates.index if AGE_MIN <= a <= AGE_MAX)

        # Tenures: only those present in the sheet AND within the segment's allowed range
        valid_tenure_years = sorted(tenure_map.keys())
        valid_tenure_years = [
            yr for yr in valid_tenure_years
            if min_tenure <= yr <= max_tenure
        ]

        if not valid_ages or not valid_tenure_years:
            raise ValueError("No valid Age/Tenure combinations found within the allowed limits.")

        rows = []
        for age_v in valid_ages:
            for tenure_v in valid_tenure_years:
                try:
                    base_rate = get_base_rate(df_rates, tenure_map, age_v, tenure_v)
                    rate = apply_loading(base_rate, loading_pct)
                    rows.append({
                        "Age": age_v,
                        "Tenure (Yrs)": tenure_v,
                        "Loading %": loading_pct,
                        "Rate": round(rate, 4),
                    })
                except Exception:
                    continue

        df_table = pd.DataFrame(rows)

        st.success(
            f"✅ Generated {len(df_table)} rate combinations for "
            f"{segment} | {life_type} Life | Loading {loading_pct}%."
        )
        st.dataframe(df_table, use_container_width=True)

        output_file = "Loaded_Rate_Table.xlsx"
        df_table.to_excel(output_file, index=False)

        with open(output_file, "rb") as file:
            st.download_button(
                label="⬇ Download Rate Table Excel",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
