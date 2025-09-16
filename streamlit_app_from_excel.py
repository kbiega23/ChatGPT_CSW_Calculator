import math
import re
from typing import Dict, List, Optional, Tuple

import pandas as pd
import numpy as np
import streamlit as st

# =========================================================
# CONFIG
# =========================================================
WORKBOOK_FILENAME = "CSW Savings Calculator 2_0_0_Unlocked.xlsx"
WEATHER_SHEET = "Weather Information"
LOOKUP_SHEET = "Savings Lookup"   # exact name in your file
APP_TITLE = "Commercial Secondary Windows (CSW) Savings Calculator"

st.set_page_config(page_title=APP_TITLE, layout="centered")
st.title(APP_TITLE)
st.caption("Excel-driven prototype: uses your workbook’s Weather and Savings Lookup tables with hours interpolation.")


# =========================================================
# UTIL: robust column matcher (by substring, case-insensitive)
# =========================================================
def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """
    Return the first column in df whose name contains ANY of the candidates (case-insensitive).
    Candidates are substrings like 'Electric Savings' or 'kWh/SF'.
    """
    cols = list(df.columns)
    low_cols = [c.lower() for c in cols]
    for cand in candidates:
        lc = cand.lower()
        for i, col in enumerate(low_cols):
            if lc in col:
                return cols[i]
    return None


# =========================================================
# LOAD WORKBOOK
# =========================================================
@st.cache_data(show_spinner=True)
def load_workbook() -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(WORKBOOK_FILENAME)

    # Parse Weather Information: repeated 4-col blocks [City, HDD, CDD, State]
    wsheet = pd.read_excel(xls, sheet_name=WEATHER_SHEET, header=None)
    records = []
    max_rows, max_cols = wsheet.shape
    for c0 in range(0, max_cols, 4):
        # guard against ragged rows
        for r in range(max_rows):
            city = wsheet.iat[r, c0] if c0 < max_cols else None
            hdd  = wsheet.iat[r, c0+1] if c0+1 < max_cols else None
            cdd  = wsheet.iat[r, c0+2] if c0+2 < max_cols else None
            state= wsheet.iat[r, c0+3] if c0+3 < max_cols else None

            # filter header rows and empties
            if isinstance(city, str) and isinstance(state, str):
                if city.strip().lower() == "city" or state.strip().lower() == "state":
                    continue
                # HDD/CDD should be numeric (or coercible)
                try:
                    hdd_val = int(float(hdd)) if pd.notnull(hdd) else None
                    cdd_val = int(float(cdd)) if pd.notnull(cdd) else None
                except Exception:
                    continue
                if (hdd_val is not None) and (cdd_val is not None):
                    records.append(
                        {"State": str(state).strip(),
                         "City": str(city).strip(),
                         "HDD": hdd_val,
                         "CDD": cdd_val}
                    )

    weather_df = pd.DataFrame(records).drop_duplicates().reset_index(drop=True)
    # Sort for nicer dropdowns
    weather_df = weather_df.sort_values(["State", "City"]).reset_index(drop=True)

    # Savings Lookup (keep all columns; we’ll match the ones we need by name)
    lookup_df = pd.read_excel(xls, sheet_name=LOOKUP_SHEET, header=0)
    # Drop all-empty columns/rows
    lookup_df = lookup_df.dropna(how="all", axis=0).dropna(how="all", axis=1)

    return {"weather": weather_df, "lookup": lookup_df}


data = load_workbook()
weather_df: pd.DataFrame = data["weather"]
lookup_df: pd.DataFrame = data["lookup"]


# =========================================================
# IDENTIFY KEY COLUMNS IN LOOKUP (by fuzzy names)
# =========================================================
# NOTE: These lists are intentionally broad so the script survives punctuation/spacing differences.
COL_STATE = "State"  # Weather table only
COL_CITY = "City"    # Weather table only

COL_BASE = find_col(lookup_df, ["Base"])  # sometimes a key/ID (not required)
COL_CSWTYPE = find_col(lookup_df, ["CSW Type", "Secondary", "CSW"])
COL_BLDG_SIZE = find_col(lookup_df, ["Building Size", "Bldg Size", "Size"])
COL_BLDG_TYPE = find_col(lookup_df, ["Building Type"])
COL_HVAC = find_col(lookup_df, ["HVAC System", "HVAC System Type", "HVAC"])
COL_FUEL = find_col(lookup_df, ["Fuel Type", "Fuel"])
COL_PTHP = find_col(lookup_df, ["PTHP"])  # may or may not be present
COL_HOURS = find_col(lookup_df, ["Hours", "Operating Hours", "Annual Hours", "Op Hours"])

COL_E_KWH_SF = find_col(lookup_df, ["Electric Savings", "kWh/SF"])
COL_G_THERM_SF = find_col(lookup_df, ["Gas Savings", "therms/SF", "Therms/SF"])

# Optional EUI / Loads (if present we’ll show them)
COL_BASE_EUI = find_col(lookup_df, ["Base EUI"])
COL_CSW_EUI  = find_col(lookup_df, ["CSW EUI"])
COL_COOL_LOAD_RED = find_col(lookup_df, ["Clg Load Reduced", "Cooling Load Reduced", "Clg Load Red"])
COL_HTG_LOAD_RED  = find_col(lookup_df, ["Htg Load Reduced", "Heating Load Reduced", "Htg Load Red"])
COL_BASE_PEAK_COOL_SF = find_col(lookup_df, ["Baseline Peak Cooling", "Btuh/SF"])

# Sanity check
missing = []
for name, col in {
    "CSW Type": COL_CSWTYPE,
    "Building Size": COL_BLDG_SIZE,
    "Building Type": COL_BLDG_TYPE,
    "HVAC System": COL_HVAC,
    "Fuel Type": COL_FUEL,
    "Hours": COL_HOURS,
    "Electric Savings kWh/SF": COL_E_KWH_SF,
    "Gas Savings therms/SF": COL_G_THERM_SF,
}.items():
    if col is None:
        missing.append(name)

if missing:
    st.warning(
        "Heads up — these expected columns were not found in the **Savings Lookup** sheet: "
        + ", ".join(missing)
        + ". The app will still try to run, but please confirm the column names in the Excel file."
    )


# =========================================================
# UI — Step 1: Lead / Project info
# =========================================================
st.subheader("Step 1 — Project Information")
with st.form("lead_info"):
    colA, colB = st.columns(2)
    with colA:
        project_name = st.text_input("Project Name*", placeholder="e.g., Lakewood Office Tower")
        contact_name = st.text_input("Contact Name*", placeholder="Your name")
        company      = st.text_input("Company", placeholder="(optional)")
    with colB:
        email        = st.text_input("Email*", placeholder="name@company.com")
        phone        = st.text_input("Phone", placeholder="(optional)")
        notes        = st.text_area("Notes", placeholder="(optional)")
    lead_ok = st.form_submit_button("Save & Continue", type="primary")

if not lead_ok:
    st.stop()


# =========================================================
# UI — Step 2: Location (auto HDD/CDD)
# =========================================================
st.subheader("Step 2 — Location")
states = sorted(weather_df[COL_STATE].unique())
state = st.selectbox("State*", states)

cities_in_state = weather_df.loc[weather_df[COL_STATE] == state, COL_CITY].tolist()
city = st.selectbox("City*", cities_in_state)

sel_weather = weather_df[(weather_df[COL_STATE] == state) & (weather_df[COL_CITY] == city)]
if sel_weather.empty:
    st.error("Selected location not found in Weather Information.")
    st.stop()
HDD = int(sel_weather["HDD"].iloc[0])
CDD = int(sel_weather["CDD"].iloc[0])
st.info(f"HDD = **{HDD}**, CDD = **{CDD}** (auto from Weather Information)")


# =========================================================
# UI — Step 3: Building details
# =========================================================
st.subheader("Step 3 — Building Details")

# Pull dropdown values directly from lookup table so they match your Excel model exactly
def uniques(colname: Optional[str]) -> List[str]:
    if colname is None or colname not in lookup_df.columns:
        return []
    vals = sorted([str(x) for x in lookup_df[colname].dropna().unique().tolist()])
    # Move blanks to end if any
    return [v for v in vals if v.strip()] + [v for v in vals if not v.strip()]

bldg_sizes = uniques(COL_BLDG_SIZE) or ["Small", "Medium", "Large"]
bldg_types = uniques(COL_BLDG_TYPE) or ["Office"]
hvac_types = uniques(COL_HVAC) or ["VAV w/ Reheat", "Packaged RTU", "Heat Pump"]
fuel_types = uniques(COL_FUEL) or ["Natural Gas", "Electricity"]
csw_types  = uniques(COL_CSWTYPE) or ["Single", "Dual"]

col1, col2 = st.columns(2)
with col1:
    building_size = st.selectbox("Building Size*", bldg_sizes)
    building_type = st.selectbox("Building Type*", bldg_types)
with col2:
    hvac_type = st.selectbox("HVAC System Type*", hvac_types)
    fuel_type = st.selectbox("Heating Fuel*", fuel_types)

# PTHP flag if present
pthp_value = None
if COL_PTHP and COL_PTHP in lookup_df.columns:
    # Infer choices from table
    pthp_choices = uniques(COL_PTHP) or ["Yes", "No"]
    pthp_value = st.selectbox("PTHP", pthp_choices)

# Hours range from table (for interpolation)
if COL_HOURS and COL_HOURS in lookup_df.columns:
    hours_vals = sorted(list(set([int(x) for x in lookup_df[COL_HOURS].dropna().astype(float).astype(int)])))
    if hours_vals:
        min_h, max_h = min(hours_vals), max(hours_vals)
    else:
        min_h, max_h = 1000, 8760
else:
    min_h, max_h = 1000, 8760

colH1, colH2 = st.columns(2)
with colH1:
    operating_hours = st.slider("Annual Operating Hours*", min_value=int(min_h), max_value=int(max_h), value=min(4000, int(max_h)))
with colH2:
    floor_area = st.number_input("Total Floor Area (ft²)*", min_value=1000, value=10000, step=500)

# Envelope / window selections
st.subheader("Step 4 — Envelope & Systems")
colE1, colE2, colE3 = st.columns(3)
with colE1:
    existing_window = st.selectbox("Existing Window Type*", ["Single Pane", "Dual Pane"])
with colE2:
    csw_type = st.selectbox("Secondary Window Type*", csw_types)
with colE3:
    csw_area = st.number_input("Sq.ft. of CSW installed*", min_value=10, value=1000, step=10)

# Utility rates
st.subheader("Step 5 — Utility Rates")
colU1, colU2 = st.columns(2)
with colU1:
    elec_rate = st.number_input("Electric Rate ($/kWh)*", min_value=0.0, value=0.12, step=0.01, format="%.2f")
with colU2:
    gas_rate = st.number_input("Gas Rate ($/therm)*", min_value=0.0, value=1.10, step=0.01, format="%.2f")


# =========================================================
# FILTER + INTERPOLATE LOOKUP ROWS
# =========================================================
def filter_lookup() -> pd.DataFrame:
    df = lookup_df.copy()

    def eq(col, val):
        if col is None or col not in df.columns:
            return pd.Series([True] * len(df))  # don’t filter if column missing
        # compare as strings to be robust
        return df[col].astype(str).str.strip().str.lower() == str(val).strip().lower()

    mask = (
        eq(COL_CSWTYPE, csw_type)
        & eq(COL_BLDG_SIZE, building_size)
        & eq(COL_BLDG_TYPE, building_type)
        & eq(COL_HVAC, hvac_type)
        & eq(COL_FUEL, fuel_type)
    )
    if pthp_value is not None and (COL_PTHP in df.columns):
        mask = mask & eq(COL_PTHP, pthp_value)

    sub = df[mask].dropna(how="all")
    return sub


def interpolate_by_hours(sub: pd.DataFrame, target_hours: int, cols_to_interp: List[str]) -> pd.Series:
    """
    Linearly interpolate numeric columns in cols_to_interp to the requested target_hours.
    If exact match exists, return that row’s values.
    """
    if COL_HOURS is None or COL_HOURS not in sub.columns:
        # No hours column; take the first non-null row
        return sub.iloc[0][cols_to_interp]

    # Keep only rows with numeric hours
    tmp = sub.copy()
    tmp = tmp[pd.to_numeric(tmp[COL_HOURS], errors="coerce").notna()]
    if tmp.empty:
        return sub.iloc[0][cols_to_interp]

    tmp[COL_HOURS] = tmp[COL_HOURS].astype(float)
    tmp = tmp.sort_values(COL_HOURS)

    # Exact match?
    exact = tmp[np.isclose(tmp[COL_HOURS], target_hours)]
    if not exact.empty:
        return exact.iloc[0][cols_to_interp]

    # Neighbors
    lower = tmp[tmp[COL_HOURS] <= target_hours].tail(1)
    upper = tmp[tmp[COL_HOURS] >= target_hours].head(1)

    if lower.empty:
        return upper.iloc[0][cols_to_interp]
    if upper.empty:
        return lower.iloc[0][cols_to_interp]

    h0, h1 = float(lower[COL_HOURS].iloc[0]), float(upper[COL_HOURS].iloc[0])
    if math.isclose(h0, h1):
        return lower.iloc[0][cols_to_interp]

    frac = (target_hours - h0) / (h1 - h0)

    # Linear interpolation for numeric columns; for non-numeric we take lower value
    out = {}
    for c in cols_to_interp:
        if (c is None) or (c not in tmp.columns):
            out[c] = np.nan
            continue
        try:
            v0 = float(lower[c].iloc[0])
            v1 = float(upper[c].iloc[0])
            out[c] = v0 + frac * (v1 - v0)
        except Exception:
            out[c] = lower[c].iloc[0]
    return pd.Series(out)


# =========================================================
# RUN CALCULATION
# =========================================================
if st.button("Calculate Savings", type="primary"):
    sub = filter_lookup()

    if sub.empty:
        st.error("No matching row found in Savings Lookup for the selected inputs."
                 " Try a different combination (Building Size/Type, HVAC, Fuel, CSW Type).")
        st.stop()

    # Columns we want to pull (if present) and interpolate by hours
    interp_cols = [c for c in [COL_E_KWH_SF, COL_G_THERM_SF, COL_COOL_LOAD_RED, COL_HTG_LOAD_RED,
                               COL_BASE_EUI, COL_CSW_EUI, COL_BASE_PEAK_COOL_SF] if c is not None]
    vals = interpolate_by_hours(sub, operating_hours, interp_cols)

    # Energy savings per SF (fallback to 0 if not present)
    e_kwh_sf = float(vals.get(COL_E_KWH_SF, 0) or 0)
    g_therm_sf = float(vals.get(COL_G_THERM_SF, 0) or 0)

    # Compute totals for the installed CSW area
    kWh = e_kwh_sf * csw_area
    therms = g_therm_sf * csw_area
    dollars = kWh * elec_rate + therms * gas_rate

    st.success("Estimated Annual Savings")
    colR1, colR2, colR3 = st.columns(3)
    with colR1:
        st.metric("Electric Energy", f"{kWh:,.0f} kWh/yr")
    with colR2:
        st.metric("Gas Energy", f"{therms:,.0f} therms/yr")
    with colR3:
        st.metric("Utility Savings", f"${dollars:,.0f}/yr")

    # Optional: show EUIs and peak/cooling/heating load reductions if columns exist
    extras = []
    if COL_BASE_EUI in vals and pd.notnull(vals[COL_BASE_EUI]):
        extras.append(("Baseline EUI", f"{float(vals[COL_BASE_EUI]):,.2f} kBtu/sf-yr"))
    if COL_CSW_EUI in vals and pd.notnull(vals[COL_CSW_EUI]):
        extras.append(("CSW EUI", f"{float(vals[COL_CSW_EUI]):,.2f} kBtu/sf-yr"))
    if COL_COOL_LOAD_RED in vals and pd.notnull(vals[COL_COOL_LOAD_RED]):
        extras.append(("Cooling Load Reduced (Btuh/sf)", f"{float(vals[COL_COOL_LOAD_RED]):,.0f}"))
    if COL_HTG_LOAD_RED in vals and pd.notnull(vals[COL_HTG_LOAD_RED]):
        extras.append(("Heating Load Reduced (Btuh/sf)", f"{float(vals[COL_HTG_LOAD_RED]):,.0f}"))
    if COL_BASE_PEAK_COOL_SF in vals and pd.notnull(vals[COL_BASE_PEAK_COOL_SF]):
        extras.append(("Baseline Peak Cooling (Btuh/sf)", f"{float(vals[COL_BASE_PEAK_COOL_SF]):,.0f}"))

    if extras:
        st.markdown("#### Additional Metrics")
        for label, value in extras:
            st.write(f"- **{label}:** {value}")

    # Summary card (lead/context)
    st.markdown("#### Project Summary")
    st.json({
        "Project": project_name,
        "Contact": contact_name,
        "Company": company,
        "Email": email,
        "Phone": phone,
        "Location": f"{city}, {state}",
        "Building Size": building_size,
        "Building Type": building_type,
        "HVAC Type": hvac_type,
        "Fuel Type": fuel_type,
        "CSW Type": csw_type,
        "Operating Hours": operating_hours,
        "Floor Area (sf)": floor_area,
        "CSW Area (sf)": csw_area,
        "Elec Rate ($/kWh)": elec_rate,
        "Gas Rate ($/therm)": gas_rate
    })

# =========================================================
# DEBUG (expandable)
# =========================================================
with st.expander("Debug / Inspect tables"):
    st.write("Weather rows:", len(weather_df))
    st.dataframe(weather_df.head(20))
    st.write("Lookup rows:", len(lookup_df))
    st.dataframe(lookup_df.head(20))
    st.caption("If dropdowns are empty, confirm the column names exist in the Lookup sheet.")
