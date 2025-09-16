import math
from typing import List, Optional
import numpy as np
import pandas as pd
import streamlit as st

# =========================
# CONFIG (exact names)
# =========================
WORKBOOK_FILENAME = "CSW Savings Calculator 2_0_0_Unlocked.xlsx"
WEATHER_SHEET = "Weather Information"
LOOKUP_SHEET  = "Savings Lookup"

# Exact column names you provided (Savings Lookup)
COL_CSW_TYPE   = "CSW_Type"
COL_BLDG_SIZE  = "Building_Size"
COL_BLDG_TYPE  = "Building_Type"
COL_HVAC_TYPE  = "HVAC_System_Type"
COL_FUEL_TYPE  = "Fuel_Type"
COL_PTHP       = "PTHP"                      # Optional
COL_HOURS      = "Hours"                     # numeric
COL_E_HEAT     = "Electric_savings_Heat_kWhperSF"
COL_E_COOL     = "electric_savings_Cooling_and_Aux_kWhperSF"
COL_G_THERM    = "Gas_savings_Heat_therms"

# Optional extras (if present we’ll display)
COL_BASE_EUI   = "Base_EUI_kBtuperSFperyr"
COL_CSW_EUI    = "CSW_EUI_kBtuperSFperyr"
COL_COOL_RED   = "Clg_Load_reduced_BtuhperSF"
COL_HEAT_RED   = "htg_load_reduced_BtuhperSF"
COL_BASE_PEAK  = "baseline_peak_cooling_BtuhperSF"

# =========================
# PAGE
# =========================
st.set_page_config(page_title="CSW Savings Calculator", layout="centered")
st.title("Commercial Secondary Windows (CSW) Savings Calculator")
st.caption("Excel-driven: filters Savings Lookup by your selections and linearly interpolates by Hours.")

# =========================
# LOADERS
# =========================
@st.cache_data(show_spinner=True)
def load_weather() -> pd.DataFrame:
    """Parse Weather Information (repeating 4-col blocks [City, HDD, CDD, State])."""
    df = pd.read_excel(WORKBOOK_FILENAME, sheet_name=WEATHER_SHEET, header=None)
    recs = []
    rows, cols = df.shape
    for c0 in range(0, cols, 4):
        for r in range(rows):
            city  = df.iat[r, c0]   if c0 < cols else None
            hdd   = df.iat[r, c0+1] if c0+1 < cols else None
            cdd   = df.iat[r, c0+2] if c0+2 < cols else None
            state = df.iat[r, c0+3] if c0+3 < cols else None
            if isinstance(city, str) and isinstance(state, str):
                if city.strip().lower() == "city" or state.strip().lower() == "state":
                    continue
                try:
                    hddv = int(float(hdd))
                    cddv = int(float(cdd))
                except Exception:
                    continue
                recs.append({"State": state.strip(), "City": city.strip(), "HDD": hddv, "CDD": cddv})
    out = pd.DataFrame(recs).drop_duplicates().sort_values(["State","City"]).reset_index(drop=True)
    return out

@st.cache_data(show_spinner=True)
def load_lookup() -> pd.DataFrame:
    """Load Savings Lookup with your flattened headers; drop empty rows/cols."""
    df = pd.read_excel(WORKBOOK_FILENAME, sheet_name=LOOKUP_SHEET, header=0)
    # Normalize whitespace on column names in case Excel added trailing spaces
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all", axis=0).dropna(how="all", axis=1)
    return df

weather_df = load_weather()
lookup_df  = load_lookup()

# Validate required columns
required_cols = [COL_CSW_TYPE, COL_BLDG_SIZE, COL_BLDG_TYPE, COL_HVAC_TYPE, COL_FUEL_TYPE, COL_HOURS, COL_E_HEAT, COL_E_COOL, COL_G_THERM]
missing = [c for c in required_cols if c not in lookup_df.columns]
if missing:
    st.error(f"Missing required columns in **{LOOKUP_SHEET}**: {', '.join(missing)}")
    st.stop()

# =========================
# UI — Step 1: Lead info
# =========================
st.subheader("Step 1 — Project Information")
with st.form("lead"):
    c1, c2 = st.columns(2)
    with c1:
        project = st.text_input("Project Name*", "")
        contact = st.text_input("Contact Name*", "")
        company = st.text_input("Company", "")
    with c2:
        email   = st.text_input("Email*", "")
        phone   = st.text_input("Phone", "")
        notes   = st.text_area("Notes")
    ok = st.form_submit_button("Save & Continue", type="primary")
if not ok:
    st.stop()

# =========================
# UI — Step 2: Location
# =========================
st.subheader("Step 2 — Location")
states = sorted(weather_df["State"].unique())
state = st.selectbox("State*", states)
cities = weather_df.loc[weather_df["State"] == state, "City"].tolist()
city  = st.selectbox("City*", cities)
loc = weather_df[(weather_df["State"] == state) & (weather_df["City"] == city)]
HDD = int(loc["HDD"].iloc[0]); CDD = int(loc["CDD"].iloc[0])
st.info(f"HDD = **{HDD}**, CDD = **{CDD}** (from Weather Information)")

# =========================
# UI — Step 3: Building & Systems
# =========================
st.subheader("Step 3 — Building & Systems")

def uniq(col: str) -> List[str]:
    if col in lookup_df.columns:
        out = sorted([str(x) for x in lookup_df[col].dropna().unique()])
        return [o for o in out if o.strip()] + [o for o in out if not o.strip()]
    return []

bldg_sizes = uniq(COL_BLDG_SIZE) or ["Small","Medium","Large"]
bldg_types = uniq(COL_BLDG_TYPE) or ["Office"]
hvac_opts  = uniq(COL_HVAC_TYPE) or ["VAV w/ Reheat","Packaged RTU","Heat Pump"]
fuel_opts  = uniq(COL_FUEL_TYPE) or ["Natural Gas","Electricity"]
csw_opts   = uniq(COL_CSW_TYPE)  or ["Single","Dual"]
pthp_opts  = uniq(COL_PTHP) if COL_PTHP in lookup_df.columns else []

c1, c2, c3 = st.columns(3)
with c1:
    building_size = st.selectbox("Building Size*", bldg_sizes)
with c2:
    building_type = st.selectbox("Building Type*", bldg_types)
with c3:
    hvac_type = st.selectbox("HVAC System Type*", hvac_opts)

c4, c5, c6 = st.columns(3)
with c4:
    fuel_type = st.selectbox("Heating Fuel*", fuel_opts)
with c5:
    csw_type = st.selectbox("Secondary Window Type*", csw_opts)
with c6:
    pthp_val = st.selectbox("PTHP", pthp_opts) if pthp_opts else ""

# Hours slider from table min/max
hrs_series = pd.to_numeric(lookup_df[COL_HOURS], errors="coerce").dropna()
min_h, max_h = (int(hrs_series.min()), int(hrs_series.max())) if not hrs_series.empty else (1000, 8760)

c7, c8 = st.columns(2)
with c7:
    hours = st.slider("Annual Operating Hours*", min_value=min_h, max_value=max_h, value=min(4000, max_h))
with c8:
    floor_area = st.number_input("Total Floor Area (ft²)*", min_value=1000, value=10000, step=500)

st.subheader("Step 4 — Envelope & Rates")
c9, c10, c11 = st.columns(3)
with c9:
    existing_window = st.selectbox("Existing Window Type*", ["Single Pane","Dual Pane"])
with c10:
    csw_area = st.number_input("Sq.ft. of CSW installed*", min_value=10, value=1000, step=10)
with c11:
    elec_rate = st.number_input("Electric Rate ($/kWh)*", min_value=0.0, value=0.12, step=0.01, format="%.2f")

c12, c13 = st.columns(2)
with c12:
    gas_rate = st.number_input("Gas Rate ($/therm)*", min_value=0.0, value=1.10, step=0.01, format="%.2f")
with c13:
    pass

# =========================
# FILTER & INTERPOLATE
# =========================
def filter_lookup(df: pd.DataFrame) -> pd.DataFrame:
    sub = df.copy()
    def eq(col, val):
        if col not in sub.columns:  # if optional key missing, don't filter on it
            return pd.Series([True]*len(sub))
        return sub[col].astype(str).str.strip().str.lower() == str(val).strip().lower()

    mask = (
        eq(COL_CSW_TYPE, csw_type)
        & eq(COL_BLDG_SIZE, building_size)
        & eq(COL_BLDG_TYPE, building_type)
        & eq(COL_HVAC_TYPE, hvac_type)
        & eq(COL_FUEL_TYPE, fuel_type)
    )
    if (COL_PTHP in sub.columns) and pthp_val:
        mask = mask & eq(COL_PTHP, pthp_val)

    out = sub[mask].dropna(how="all")
    return out

def interpolate_by_hours(sub: pd.DataFrame, target_hours: int, cols: List[str]) -> pd.Series:
    # Ensure numeric, sort by hours
    tmp = sub[pd.to_numeric(sub[COL_HOURS], errors="coerce").notna()].copy()
    if tmp.empty:
        return sub.iloc[0][cols]
    tmp[COL_HOURS] = tmp[COL_HOURS].astype(float)
    tmp = tmp.sort_values(COL_HOURS)

    # exact
    exact = tmp[np.isclose(tmp[COL_HOURS], target_hours)]
    if not exact.empty:
        return exact.iloc[0][cols]

    lower = tmp[tmp[COL_HOURS] <= target_hours].tail(1)
    upper = tmp[tmp[COL_HOURS] >= target_hours].head(1)
    if lower.empty: return upper.iloc[0][cols]
    if upper.empty: return lower.iloc[0][cols]

    h0, h1 = float(lower[COL_HOURS].iloc[0]), float(upper[COL_HOURS].iloc[0])
    frac = 0.0 if math.isclose(h0, h1) else (target_hours - h0) / (h1 - h0)

    out = {}
    for c in cols:
        if c not in tmp.columns or tmp[c].isna().all():
            out[c] = np.nan
            continue
        try:
            v0 = float(lower[c].iloc[0]); v1 = float(upper[c].iloc[0])
            out[c] = v0 + frac*(v1 - v0)
        except Exception:
            out[c] = lower[c].iloc[0]
    return pd.Series(out)

# =========================
# CALCULATE
# =========================
if st.button("Calculate Savings", type="primary"):
    rows = filter_lookup(lookup_df)
    if rows.empty:
        st.error("No matching rows in Savings Lookup for this combination.")
        st.stop()

    # Columns to interpolate
    cols_to_interp = [COL_E_HEAT, COL_E_COOL, COL_G_THERM, COL_BASE_EUI, COL_CSW_EUI, COL_COOL_RED, COL_HEAT_RED, COL_BASE_PEAK]
    cols_to_interp = [c for c in cols_to_interp if c in rows.columns]

    vals = interpolate_by_hours(rows, hours, cols_to_interp)

    e_heat  = float(vals.get(COL_E_HEAT, 0) or 0.0)
    e_cool  = float(vals.get(COL_E_COOL, 0) or 0.0)
    g_therm = float(vals.get(COL_G_THERM, 0) or 0.0)

    # per-SF → totals
    kWh    = (e_heat + e_cool) * csw_area
    therms = g_therm * csw_area
    dollars = kWh * elec_rate + therms * gas_rate

    st.success("Estimated Annual Savings")
    a,b,c = st.columns(3)
    a.metric("Electric Energy", f"{kWh:,.0f} kWh/yr")
    b.metric("Gas Energy", f"{therms:,.0f} therms/yr")
    c.metric("Utility Savings", f"${dollars:,.0f}/yr")

    # Optional metrics (if present)
    extras = []
    if COL_BASE_EUI in vals and pd.notnull(vals[COL_BASE_EUI]):
        extras.append(("Baseline EUI", f"{float(vals[COL_BASE_EUI]):,.2f} kBtu/sf-yr"))
    if COL_CSW_EUI in vals and pd.notnull(vals[COL_CSW_EUI]):
        extras.append(("CSW EUI", f"{float(vals[COL_CSW_EUI]):,.2f} kBtu/sf-yr"))
    if COL_COOL_RED in vals and pd.notnull(vals[COL_COOL_RED]):
        extras.append(("Cooling Load Reduced", f"{float(vals[COL_COOL_RED]):,.0f} Btuh/sf"))
    if COL_HEAT_RED in vals and pd.notnull(vals[COL_HEAT_RED]):
        extras.append(("Heating Load Reduced", f"{float(vals[COL_HEAT_RED]):,.0f} Btuh/sf"))
    if COL_BASE_PEAK in vals and pd.notnull(vals[COL_BASE_PEAK]):
        extras.append(("Baseline Peak Cooling", f"{float(vals[COL_BASE_PEAK]):,.0f} Btuh/sf"))
    if extras:
        st.markdown("#### Additional Metrics")
        for label, value in extras:
            st.write(f"- **{label}:** {value}")

    # Summary
    st.markdown("#### Project Summary")
    st.json({
        "Project": project, "Contact": contact, "Company": company, "Email": email, "Phone": phone,
        "Location": f"{city}, {state}",
        "Building Size": building_size, "Building Type": building_type,
        "HVAC Type": hvac_type, "Fuel Type": fuel_type, "PTHP": pthp_val,
        "Operating Hours": hours, "Floor Area (sf)": floor_area,
        "CSW Area (sf)": csw_area, "Elec Rate": elec_rate, "Gas Rate": gas_rate
    })

# =========================
# Debug
# =========================
with st.expander("Debug / Inspect"):
    st.write("Weather rows:", len(weather_df)); st.dataframe(weather_df.head(20))
    st.write("Lookup rows:", len(lookup_df));  st.dataframe(lookup_df.head(20))
    st.caption("If dropdowns look odd, confirm exact text in keys (CSW_Type, Building_Size, etc.).")
