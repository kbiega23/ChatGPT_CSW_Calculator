import math
from typing import Dict, List, Optional

import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# CONFIG
# =========================================================
WORKBOOK_FILENAME = "CSW Savings Calculator 2_0_0_Unlocked.xlsx"
WEATHER_SHEET = "Weather Information"
LOOKUP_SHEET = "Savings Lookup"   # name you provided
APP_TITLE = "Commercial Secondary Windows (CSW) Savings Calculator"

st.set_page_config(page_title=APP_TITLE, layout="centered")
st.title(APP_TITLE)
st.caption("Excel-driven: reads Weather & Savings Lookup with multi-row headers + hours interpolation.")

# =========================================================
# Helpers
# =========================================================
def normalize(s: str) -> str:
    """lowercase, strip, collapse spaces and punctuation for fuzzy matching"""
    if s is None:
        return ""
    x = str(s)
    # remove common punctuation
    for ch in [",", ";", ":", "/", "\\", "(", ")", "[", "]"]:
        x = x.replace(ch, " ")
    x = " ".join(x.lower().split())
    return x

def find_col(cols: List[str], candidates: List[str]) -> Optional[str]:
    """Return first column whose normalized label contains any candidate substring."""
    ncols = [(c, normalize(c)) for c in cols]
    cand_norm = [normalize(c) for c in candidates]
    for raw, n in ncols:
        for c in cand_norm:
            if c and c in n:
                return raw
    return None

def flatten_multiindex_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Join multi-row headers into single strings: 'Parent | Child | Subchild'."""
    if isinstance(df.columns, pd.MultiIndex):
        new_cols = []
        for tup in df.columns:
            # Remove Nones/empties and join
            parts = [str(x).strip() for x in tup if (x is not None and str(x).strip() != "")]
            label = " | ".join(parts) if parts else ""
            new_cols.append(label)
        out = df.copy()
        out.columns = new_cols
        return out
    else:
        return df

# =========================================================
# Load workbook (and parse weather blocks)
# =========================================================
@st.cache_data(show_spinner=True)
def load_workbook() -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(WORKBOOK_FILENAME)

    # ---- Weather: repeating 4-col blocks [City, HDD, CDD, State]
    wsheet = pd.read_excel(xls, sheet_name=WEATHER_SHEET, header=None)
    recs = []
    rows, cols = wsheet.shape
    for c0 in range(0, cols, 4):
        for r in range(rows):
            city  = wsheet.iat[r, c0]   if c0 < cols else None
            hdd   = wsheet.iat[r, c0+1] if c0+1 < cols else None
            cdd   = wsheet.iat[r, c0+2] if c0+2 < cols else None
            state = wsheet.iat[r, c0+3] if c0+3 < cols else None
            if isinstance(city, str) and isinstance(state, str):
                if city.strip().lower() == "city" or state.strip().lower() == "state":
                    continue
                try:
                    hddv = int(float(hdd))
                    cddv = int(float(cdd))
                except Exception:
                    continue
                recs.append({"State": state.strip(), "City": city.strip(), "HDD": hddv, "CDD": cddv})
    weather = pd.DataFrame(recs).drop_duplicates().sort_values(["State","City"]).reset_index(drop=True)

    # ---- Savings Lookup: try multi-row headers then fall back to single
    # Many models put headings in 2–3 rows; try up to first 3 rows as header
    try:
        raw = pd.read_excel(xls, sheet_name=LOOKUP_SHEET, header=[0,1,2])
    except Exception:
        try:
            raw = pd.read_excel(xls, sheet_name=LOOKUP_SHEET, header=[0,1])
        except Exception:
            raw = pd.read_excel(xls, sheet_name=LOOKUP_SHEET, header=0)

    lookup = flatten_multiindex_columns(raw)
    # Drop all-empty cols/rows
    lookup = lookup.dropna(how="all", axis=0).dropna(how="all", axis=1)

    return {"weather": weather, "lookup": lookup}

data = load_workbook()
weather_df = data["weather"]
lookup_df  = data["lookup"]

# =========================================================
# Identify columns in Savings Lookup (fuzzy, tolerant of stacked labels)
# =========================================================
cols = lookup_df.columns.tolist()

COL_CSWTYPE   = find_col(cols, ["csw type", "secondary window", "csw"])
COL_BLDG_SIZE = find_col(cols, ["building size", "bldg size", "size"])
COL_BLDG_TYPE = find_col(cols, ["building type"])
COL_HVAC      = find_col(cols, ["hvac system type", "hvac system", "hvac"])
COL_FUEL      = find_col(cols, ["fuel type", "fuel"])
COL_PTHP      = find_col(cols, ["pthp"])

# Hours might sit anywhere; also sometimes appears under a parent heading
COL_HOURS     = find_col(cols, ["hours", "operating hours", "annual hours", "op hours"])

# Savings sub-columns (stacked headers become 'Electric Savings | Cooling', etc.)
COL_E_KWH_SF  = find_col(cols, [
    "electric savings | cooling", "electric savings cooling", "electric savings kwh sf cooling",
    "electric savings | heat",    "electric savings heat",   "kwh sf"
])
# We’ll later prefer “cooling” + “heat” sum if both present.
# Gas/therms:
COL_G_THERM_SF = find_col(cols, [
    "gas savings |", "gas savings therms sf", "therms sf", "gas therms"
])

# Optional metrics
COL_BASE_EUI = find_col(cols, ["base eui", "baseline eui"])
COL_CSW_EUI  = find_col(cols, ["csw eui"])
COL_COOL_RED = find_col(cols, ["clg load reduced", "cooling load reduced", "clg load red"])
COL_HTG_RED  = find_col(cols, ["htg load reduced", "heating load reduced", "htg load red"])
COL_BASE_PEAK_COOL_SF = find_col(cols, ["baseline peak cooling", "btuh sf"])

# Also try to find separate cooling/heat columns explicitly if present
COL_E_KWH_SF_COOL = find_col(cols, ["electric savings | cooling", "kwh sf cooling", "cooling + aux"])
COL_E_KWH_SF_HEAT = find_col(cols, ["electric savings | heat", "kwh sf heat"])

# Report missing but keep going
missing = []
for label, c in {
    "CSW Type": COL_CSWTYPE,
    "Building Size": COL_BLDG_SIZE,
    "Building Type": COL_BLDG_TYPE,
    "HVAC System Type": COL_HVAC,
    "Fuel Type": COL_FUEL,
    "Hours": COL_HOURS,
}.items():
    if c is None:
        missing.append(label)
if missing:
    st.warning("Could not find some key lookup columns: " + ", ".join(missing))

# =========================================================
# UI – Step 1: Lead info
# =========================================================
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

# =========================================================
# UI – Step 2: Location (auto HDD/CDD)
# =========================================================
st.subheader("Step 2 — Location")
states = sorted(weather_df["State"].unique())
state = st.selectbox("State*", states)
cities = weather_df.loc[weather_df["State"] == state, "City"].tolist()
city = st.selectbox("City*", cities)

wrow = weather_df[(weather_df["State"] == state) & (weather_df["City"] == city)]
if wrow.empty:
    st.error("Location not found in Weather Information.")
    st.stop()
HDD = int(wrow["HDD"].iloc[0]); CDD = int(wrow["CDD"].iloc[0])
st.info(f"HDD = **{HDD}**, CDD = **{CDD}**")

# =========================================================
# UI – Step 3: Building + Systems
# =========================================================
st.subheader("Step 3 — Building & Systems")
def uniq(colname: Optional[str]) -> List[str]:
    if colname and colname in lookup_df.columns:
        out = sorted([str(x) for x in lookup_df[colname].dropna().unique()])
        return [o for o in out if o.strip()] + [o for o in out if not o.strip()]
    return []

bldg_sizes = uniq(COL_BLDG_SIZE) or ["Small","Medium","Large"]
bldg_types = uniq(COL_BLDG_TYPE) or ["Office"]
hvac_opts  = uniq(COL_HVAC) or ["VAV w/ Reheat","Packaged RTU","Heat Pump"]
fuel_opts  = uniq(COL_FUEL) or ["Natural Gas","Electricity"]
csw_opts   = uniq(COL_CSWTYPE) or ["Single","Dual"]
pthp_opts  = uniq(COL_PTHP) if COL_PTHP else []

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

# Hours slider uses min/max in table if present
if COL_HOURS and COL_HOURS in lookup_df.columns:
    hrs_vals = pd.to_numeric(lookup_df[COL_HOURS], errors="coerce").dropna()
    if not hrs_vals.empty:
        min_h, max_h = int(hrs_vals.min()), int(hrs_vals.max())
    else:
        min_h, max_h = 1000, 8760
else:
    min_h, max_h = 1000, 8760

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

# =========================================================
# Filter + interpolate by Hours
# =========================================================
def filter_lookup(df: pd.DataFrame) -> pd.DataFrame:
    sub = df.copy()

    def eq(col, val):
        if not col or col not in sub.columns:
            return pd.Series([True]*len(sub))
        return sub[col].astype(str).str.strip().str.lower() == str(val).strip().lower()

    m = (
        eq(COL_CSWTYPE, csw_type)
        & eq(COL_BLDG_SIZE, building_size)
        & eq(COL_BLDG_TYPE, building_type)
        & eq(COL_HVAC, hvac_type)
        & eq(COL_FUEL, fuel_type)
    )
    if COL_PTHP and COL_PTHP in sub.columns and pthp_val:
        m = m & eq(COL_PTHP, pthp_val)

    out = sub[m].dropna(how="all")
    return out

def interpolate_by_hours(sub: pd.DataFrame, target_hours: int, cols_to_interp: List[str]) -> pd.Series:
    if not COL_HOURS or COL_HOURS not in sub.columns:
        return sub.iloc[0][cols_to_interp]

    tmp = sub[pd.to_numeric(sub[COL_HOURS], errors="coerce").notna()].copy()
    if tmp.empty:
        return sub.iloc[0][cols_to_interp]

    tmp[COL_HOURS] = tmp[COL_HOURS].astype(float)
    tmp = tmp.sort_values(COL_HOURS)

    exact = tmp[np.isclose(tmp[COL_HOURS], target_hours)]
    if not exact.empty:
        return exact.iloc[0][cols_to_interp]

    lower = tmp[tmp[COL_HOURS] <= target_hours].tail(1)
    upper = tmp[tmp[COL_HOURS] >= target_hours].head(1)
    if lower.empty: return upper.iloc[0][cols_to_interp]
    if upper.empty: return lower.iloc[0][cols_to_interp]

    h0, h1 = float(lower[COL_HOURS].iloc[0]), float(upper[COL_HOURS].iloc[0])
    frac = 0.0 if math.isclose(h0, h1) else (target_hours - h0) / (h1 - h0)

    out = {}
    for c in cols_to_interp:
        if not c or c not in tmp.columns:
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
# Calculate
# =========================================================
if st.button("Calculate Savings", type="primary"):
    sub = filter_lookup(lookup_df)
    if sub.empty:
        st.error("No matching rows in Savings Lookup for this combination. Try different selections.")
        st.stop()

    # Prefer separate Cooling/Heat electric columns if both available; else fall back to a single electric column
    electric_cols = []
    if COL_E_KWH_SF_COOL: electric_cols.append(COL_E_KWH_SF_COOL)
    if COL_E_KWH_SF_HEAT: electric_cols.append(COL_E_KWH_SF_HEAT)
    if not electric_cols and COL_E_KWH_SF:
        electric_cols = [COL_E_KWH_SF]

    cols_to_interp = [c for c in (electric_cols + [COL_G_THERM_SF, COL_BASE_EUI, COL_CSW_EUI, COL_COOL_RED, COL_HTG_RED, COL_BASE_PEAK_COOL_SF]) if c]

    vals = interpolate_by_hours(sub, hours, cols_to_interp)

    # Electric kWh per SF
    e_kwh_sf = 0.0
    for c in electric_cols:
        if c in vals and pd.notnull(vals[c]):
            e_kwh_sf += float(vals[c])

    # Gas therms per SF
    g_therm_sf = float(vals.get(COL_G_THERM_SF, 0) or 0)

    # Convert to totals
    kWh = e_kwh_sf * csw_area
    therms = g_therm_sf * csw_area
    dollars = kWh * elec_rate + therms * gas_rate

    st.success("Estimated Annual Savings")
    a, b, c = st.columns(3)
    a.metric("Electric Energy", f"{kWh:,.0f} kWh/yr")
    b.metric("Gas Energy", f"{therms:,.0f} therms/yr")
    c.metric("Utility Savings", f"${dollars:,.0f}/yr")

    # Optional extras if present
    extras = []
    if COL_BASE_EUI and pd.notnull(vals.get(COL_BASE_EUI, np.nan)):
        extras.append(("Baseline EUI", f"{float(vals[COL_BASE_EUI]):,.2f} kBtu/sf-yr"))
    if COL_CSW_EUI and pd.notnull(vals.get(COL_CSW_EUI, np.nan)):
        extras.append(("CSW EUI", f"{float(vals[COL_CSW_EUI]):,.2f} kBtu/sf-yr"))
    if COL_COOL_RED and pd.notnull(vals.get(COL_COOL_RED, np.nan)):
        extras.append(("Cooling Load Reduced (Btuh/sf)", f"{float(vals[COL_COOL_RED]):,.0f}"))
    if COL_HTG_RED and pd.notnull(vals.get(COL_HTG_RED, np.nan)):
        extras.append(("Heating Load Reduced (Btuh/sf)", f"{float(vals[COL_HTG_RED]):,.0f}"))
    if COL_BASE_PEAK_COOL_SF and pd.notnull(vals.get(COL_BASE_PEAK_COOL_SF, np.nan)):
        extras.append(("Baseline Peak Cooling (Btuh/sf)", f"{float(vals[COL_BASE_PEAK_COOL_SF]):,.0f}"))

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

# =========================================================
# Debug
# =========================================================
with st.expander("Debug / Inspect"):
    st.write("Weather rows:", len(weather_df))
    st.dataframe(weather_df.head(20))
    st.write("Lookup rows:", len(lookup_df))
    st.dataframe(lookup_df.head(20))
    st.caption("If filters are empty, check the Savings Lookup headers; stacked headers are flattened as 'Parent | Child'.")
