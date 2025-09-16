import math
from typing import Dict, List

import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# CONFIG
# =========================================================
WORKBOOK_FILENAME = "CSW Savings Calculator 2_0_0_Unlocked.xlsx"
WEATHER_SHEET = "Weather Information"
LOOKUP_SHEET  = "Savings Lookup"

APP_TITLE = "Commercial Secondary Windows (CSW) Savings Calculator"
st.set_page_config(page_title=APP_TITLE, layout="centered")
st.title(APP_TITLE)
st.caption("Excel-driven: reads Weather & Savings Lookup, filters by keys, and linearly interpolates by Hours.")

# =========================================================
# SESSION STATE (wizard + cache reload)
# =========================================================
if "reload_token" not in st.session_state:
    st.session_state.reload_token = 0
if "lead_saved" not in st.session_state:
    st.session_state.lead_saved = False

# persist form fields
for k, v in {
    "project": "", "contact": "", "company": "",
    "email": "", "phone": "", "notes": "",
    "state": "", "city": ""
}.items():
    st.session_state.setdefault(k, v)

# cache-busting without experimental APIs
col_reload, _ = st.columns([1, 3])
with col_reload:
    if st.button("Reload workbook (clear cache)"):
        st.session_state.reload_token += 1

# =========================================================
# Header normalization / matching helpers (for Savings Lookup)
# =========================================================
def _norm(s: str) -> str:
    if s is None:
        return ""
    x = str(s).lower()
    for ch in ["\xa0", " ", "_", "-", ",", ".", "/", "\\", "(", ")", "[", "]", ":", ";"]:
        x = x.replace(ch, "")
    return x

EXPECTED_CANONICAL = {
    # Keys
    "CSW_Type":        ["csw_type", "csw type", "cswtype", "csw"],
    "Building_Size":   ["building_size", "building size", "bldg_size", "bldg size", "size"],
    "Building_Type":   ["building_type", "building type"],
    "HVAC_System_Type":["hvac_system_type", "hvac system type", "hvac system", "hvac"],
    "Fuel_Type":       ["fuel_type", "fuel type", "fuel"],
    "PTHP":            ["pthp"],
    "Hours":           ["hours", "operating hours", "annual hours", "op hours"],
    # Savings per-SF
    "Electric_savings_Heat_kWhperSF": [
        "electric_savings_heat_kwhpersf", "electric savings heat kwhpersf",
        "electric_savings_kwhpersf_heat", "kwhpersfheat", "kwh per sf heat"
    ],
    "electric_savings_Cooling_and_Aux_kWhperSF": [
        "electric_savings_cooling_and_aux_kwhpersf", "electric savings cooling and aux kwhpersf",
        "electric_savings_kwhpersf_cooling", "kwhpersfcooling", "kwh per sf cooling", "cooling+aux", "coolingandaux"
    ],
    "Gas_savings_Heat_thermsperSF": [
        "gas_savings_heat_thermspersf", "gas savings heat thermspersf",
        "gas_savings_heat_therms", "gas_savings_thermspersf", "thermspersf", "therms per sf", "gas therms per sf"
    ],
}
OPTIONAL_CANONICAL = {
    "Base_EUI_kBtuperSFperyr":        ["base_eui_kbtupersfperyr", "base eui", "baseline eui"],
    "CSW_EUI_kBtuperSFperyr":         ["csw_eui_kbtupersfperyr", "csw eui"],
    "baseline_peak_cooling_BtuhperSF":["baseline_peak_cooling_btuhpersf", "baseline peak cooling", "btuh per sf", "btuh/sf"],
    "Clg_Load_reduced_BtuhperSF":     ["clg_load_reduced_btuhpersf", "cooling load reduced", "clg load reduced"],
    "htg_load_reduced_BtuhperSF":     ["htg_load_reduced_btuhpersf", "heating load reduced", "htg load reduced"],
}

def _best_header_row(df_nohdr: pd.DataFrame, scan_rows: int = 25) -> int:
    expected_norms = set(sum(EXPECTED_CANONICAL.values(), []))
    best_r, best_hits = 0, -1
    for r in range(min(scan_rows, df_nohdr.shape[0])):
        row_vals = df_nohdr.iloc[r].fillna("").astype(str).tolist()
        row_norms = {_norm(v) for v in row_vals if str(v).strip() != ""}
        hits = len(expected_norms.intersection(row_norms))
        if hits > best_hits:
            best_hits, best_r = hits, r
    return best_r

def _map_to_canonical(cols: List[str]) -> Dict[str, str]:
    inv = {}
    for canon, variants in {**EXPECTED_CANONICAL, **OPTIONAL_CANONICAL}.items():
        for v in variants:
            inv[v] = canon
    mapping = {}
    for c in cols:
        n = _norm(c)
        if n in inv:
            mapping[c] = inv[n]
    return mapping

# =========================================================
# Loaders (cache keyed by reload_token)
# =========================================================
@st.cache_data(show_spinner=True)
def load_weather(reload_token: int) -> pd.DataFrame:
    """
    Robustly read Weather Information with a possibly offset header row.
    Works with headers like:
      State | Cities | Heating Degree Days (HDD) | Cooling Degree Days (CDD)
    and variants such as "City", "HDD", "CDD".
    """
    # Read without header so we can detect the real header row
    raw = pd.read_excel(WORKBOOK_FILENAME, sheet_name=WEATHER_SHEET, header=None)
    raw = raw.dropna(how="all").dropna(how="all", axis=1).reset_index(drop=True)

    # Find the most likely header row by matching tokens
    def norm(s: str) -> str:
        if s is None: return ""
        x = str(s).lower()
        for ch in ["\xa0", " ", "_", "-", ",", ".", "/", "\\", "(", ")", "[", "]", ":", ";"]:
            x = x.replace(ch, "")
        return x

    def header_score(row_vals: list[str]) -> int:
        tokens = {norm(v) for v in row_vals if str(v).strip() != ""}
        score = 0
        if any(t.startswith("state") for t in tokens): score += 1
        if any(t.startswith("city") or t.startswith("cities") for t in tokens): score += 1
        if any(("heating" in t and "hdd" in t) or t == "hdd" for t in tokens): score += 1
        if any(("cooling" in t and "cdd" in t) or t == "cdd" for t in tokens): score += 1
        return score

    best_row, best_score = 0, -1
    scan_rows = min(25, raw.shape[0])
    for r in range(scan_rows):
        vals = raw.iloc[r].fillna("").astype(str).tolist()
        sc = header_score(vals)
        if sc > best_score:
            best_row, best_score = r, sc

    # Re-read with detected header row
    df = pd.read_excel(WORKBOOK_FILENAME, sheet_name=WEATHER_SHEET, header=best_row)
    df = df.dropna(how="all").dropna(how="all", axis=1)
    df.columns = [str(c).strip() for c in df.columns]

    # Map to canonical names
    rename_map = {}
    for c in df.columns:
        cl = c.strip().lower()
        ncl = norm(c)
        if ncl.startswith("state"):
            rename_map[c] = "State"
        elif ncl.startswith("city") or ncl.startswith("cities"):
            rename_map[c] = "City"
        elif ("heating" in cl and "hdd" in cl) or ncl == "hdd":
            rename_map[c] = "HDD"
        elif ("cooling" in cl and "cdd" in cl) or ncl == "cdd":
            rename_map[c] = "CDD"

    df = df.rename(columns=rename_map)

    needed = ["State", "City", "HDD", "CDD"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        st.error("Weather sheet appears to be missing expected columns.")
        with st.expander("Show detected Weather headers"):
            st.write(list(df.columns))
            st.write(f"Detected header row index: {best_row}")
        raise ValueError(f"Weather sheet is missing columns: {', '.join(missing)}")

    # Clean + types
    df = df[needed].copy()
    df["State"] = df["State"].astype(str).str.strip()
    df["City"]  = df["City"].astype(str).str.strip()
    df["HDD"]   = pd.to_numeric(df["HDD"], errors="coerce").astype("Int64")
    df["CDD"]   = pd.to_numeric(df["CDD"], errors="coerce").astype("Int64")

    df = df.dropna(subset=["State", "City", "HDD", "CDD"]).drop_duplicates()
    df = df.sort_values(["State", "City"]).reset_index(drop=True)
    return df

@st.cache_data(show_spinner=True)
def load_lookup(reload_token: int) -> pd.DataFrame:
    """Robust loader for 'Savings Lookup': auto-detect header row; normalize & map to canonical names."""
    df0 = pd.read_excel(WORKBOOK_FILENAME, sheet_name=LOOKUP_SHEET, header=None)
    df0 = df0.dropna(how="all").dropna(how="all", axis=1).reset_index(drop=True)

    hdr_row = _best_header_row(df0, scan_rows=25)
    df = pd.read_excel(WORKBOOK_FILENAME, sheet_name=LOOKUP_SHEET, header=hdr_row)
    df = df.dropna(how="all").dropna(how="all", axis=1)
    df.columns = [str(c).strip() for c in df.columns]

    mapping = _map_to_canonical(df.columns.tolist())
    df = df.rename(columns=mapping)

    # Also accept exact canonical headers even if casing/spaces differ
    all_canons = list(EXPECTED_CANONICAL.keys()) + list(OPTIONAL_CANONICAL.keys())
    for canon in all_canons:
        if canon not in df.columns:
            for c in df.columns:
                if _norm(c) == _norm(canon):
                    df = df.rename(columns={c: canon})
                    break

    # Backward-compat: allow old gas header name
    if "Gas_savings_Heat_thermsperSF" not in df.columns and "Gas_savings_Heat_therms" in df.columns:
        df = df.rename(columns={"Gas_savings_Heat_therms": "Gas_savings_Heat_thermsperSF"})

    df = df.dropna(how="all").reset_index(drop=True)
    return df

weather_df = load_weather(st.session_state.reload_token)
lookup_df  = load_lookup(st.session_state.reload_token)

# =========================================================
# Required columns check (Lookup)
# =========================================================
required_cols = [
    "CSW_Type", "Building_Size", "Building_Type", "HVAC_System_Type", "Fuel_Type", "Hours",
    "Electric_savings_Heat_kWhperSF", "electric_savings_Cooling_and_Aux_kWhperSF", "Gas_savings_Heat_thermsperSF"
]
missing = [c for c in required_cols if c not in lookup_df.columns]
if missing:
    st.error(f"Missing required columns in **{LOOKUP_SHEET}**: {', '.join(missing)}")
    with st.expander("Show detected columns"):
        st.write(list(lookup_df.columns))
    st.stop()

# =========================================================
# WIZARD — Step 1: Project Information (persisted)
# =========================================================
st.subheader("Step 1 — Project Information")
with st.form("lead", clear_on_submit=False):
    c1, c2 = st.columns(2)
    with c1:
        st.session_state.project = st.text_input("Project Name*", st.session_state.project, key="project_input")
        st.session_state.contact = st.text_input("Contact Name*", st.session_state.contact, key="contact_input")
        st.session_state.company = st.text_input("Company", st.session_state.company, key="company_input")
    with c2:
        st.session_state.email   = st.text_input("Email*", st.session_state.email, key="email_input")
        st.session_state.phone   = st.text_input("Phone", st.session_state.phone, key="phone_input")
        st.session_state.notes   = st.text_area("Notes", st.session_state.notes, key="notes_input")
    lead_ok = st.form_submit_button("Save & Continue", type="primary")
    if lead_ok:
        st.session_state.lead_saved = True

if not st.session_state.lead_saved:
    st.info("Fill Step 1 and click **Save & Continue** to proceed.")
    st.stop()

st.success("Step 1 saved. You can keep editing above — your entries will persist.")

# =========================================================
# WIZARD — Step 2: Location (auto HDD/CDD)
# =========================================================
st.subheader("Step 2 — Location")
if weather_df.empty:
    st.error("Weather table appears empty. Please verify the **Weather Information** sheet layout.")
    st.stop()

states = weather_df["State"].astype(str).str.strip().unique().tolist()
states.sort()
if st.session_state.state not in states:
    st.session_state.state = states[0] if states else ""
st.session_state.state = st.selectbox(
    "State*", states,
    index=states.index(st.session_state.state) if st.session_state.state in states else 0,
    key="state_select"
)

cities_df = weather_df[weather_df["State"].astype(str).str.strip() == st.session_state.state]
cities = cities_df["City"].astype(str).str.strip().unique().tolist()
cities.sort()
if st.session_state.city not in cities:
    st.session_state.city = cities[0] if cities else ""
st.session_state.city = st.selectbox(
    "City*", cities,
    index=cities.index(st.session_state.city) if st.session_state.city in cities else 0,
    key="city_select"
)

sel = cities_df[cities_df["City"].astype(str).str.strip() == st.session_state.city]
if sel.empty:
    st.error("Selected location not found in Weather Information.")
    st.stop()
HDD = int(sel["HDD"].iloc[0]); CDD = int(sel["CDD"].iloc[0])
st.info(f"HDD = **{HDD}**, CDD = **{CDD}** (from Weather Information)")

# =========================================================
# WIZARD — Step 3: Building & Systems
# =========================================================
st.subheader("Step 3 — Building & Systems")
def uniq(col: str) -> List[str]:
    out = sorted([str(x) for x in lookup_df[col].dropna().unique()])
    return [o for o in out if o.strip()] + [o for o in out if not o.strip()]

bldg_sizes = uniq("Building_Size")
bldg_types = uniq("Building_Type")
hvac_opts  = uniq("HVAC_System_Type")
fuel_opts  = uniq("Fuel_Type")
csw_opts   = uniq("CSW_Type")
pthp_opts  = uniq("PTHP") if "PTHP" in lookup_df.columns else []

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

# Hours slider from table
hrs_series = pd.to_numeric(lookup_df["Hours"], errors="coerce").dropna()
min_h, max_h = (int(hrs_series.min()), int(hrs_series.max())) if not hrs_series.empty else (1000, 8760)

c7, c8 = st.columns(2)
with c7:
    hours = st.slider("Annual Operating Hours*", min_value=min_h, max_value=max_h, value=min(4000, max_h))
with c8:
    floor_area = st.number_input("Total Floor Area (ft²)*", min_value=1000, value=10000, step=500)

# =========================================================
# WIZARD — Step 4: Envelope & Rates
# =========================================================
st.subheader("Step 4 — Envelope & Rates")
c9, c10, c11 = st.columns(3)
with c9:
    existing_window = st.selectbox("Existing Window Type*", ["Single Pane", "Dual Pane"])
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
# Filtering & Interpolation
# =========================================================
def filter_lookup(df: pd.DataFrame) -> pd.DataFrame:
    sub = df.copy()
    def eq(col, val):
        if col not in sub.columns or val == "":
            return pd.Series([True]*len(sub))
        return sub[col].astype(str).str.strip().str.lower() == str(val).strip().lower()

    mask = (
        eq("CSW_Type", csw_type)
        & eq("Building_Size", building_size)
        & eq("Building_Type", building_type)
        & eq("HVAC_System_Type", hvac_type)
        & eq("Fuel_Type", fuel_type)
    )
    if "PTHP" in sub.columns and pthp_val:
        mask = mask & eq("PTHP", pthp_val)
    return sub[mask].dropna(how="all")

def interpolate_by_hours(sub: pd.DataFrame, target_hours: int, cols: List[str]) -> pd.Series:
    tmp = sub[pd.to_numeric(sub["Hours"], errors="coerce").notna()].copy()
    if tmp.empty:
        return sub.iloc[0][cols]
    tmp["Hours"] = tmp["Hours"].astype(float)
    tmp = tmp.sort_values("Hours")

    exact = tmp[np.isclose(tmp["Hours"], target_hours)]
    if not exact.empty:
        return exact.iloc[0][cols]

    lower = tmp[tmp["Hours"] <= target_hours].tail(1)
    upper = tmp[tmp["Hours"] >= target_hours].head(1)
    if lower.empty: return upper.iloc[0][cols]
    if upper.empty: return lower.iloc[0][cols]

    h0, h1 = float(lower["Hours"].iloc[0]), float(upper["Hours"].iloc[0])
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

# =========================================================
# Calculate
# =========================================================
st.subheader("Step 5 — Calculate")
if st.button("Calculate Savings", type="primary"):
    sub = filter_lookup(lookup_df)
    if sub.empty:
        st.error("No matching rows in Savings Lookup for this combination. Try different selections.")
    else:
        COLS_INTERP = [
            "Electric_savings_Heat_kWhperSF",
            "electric_savings_Cooling_and_Aux_kWhperSF",
            "Gas_savings_Heat_thermsperSF",
            "Base_EUI_kBtuperSFperyr",
            "CSW_EUI_kBtuperSFperyr",
            "Clg_Load_reduced_BtuhperSF",
            "htg_load_reduced_BtuhperSF",
            "baseline_peak_cooling_BtuhperSF",
        ]
        cols_to_interp = [c for c in COLS_INTERP if c in sub.columns]
        vals = interpolate_by_hours(sub, hours, cols_to_interp)

        e_heat   = float(vals.get("Electric_savings_Heat_kWhperSF", 0) or 0.0)
        e_cool   = float(vals.get("electric_savings_Cooling_and_Aux_kWhperSF", 0) or 0.0)
        g_therms = float(vals.get("Gas_savings_Heat_thermsperSF", 0) or 0.0)

        # per-SF → totals
        kWh    = (e_heat + e_cool) * csw_area
        therms = g_therms * csw_area
        dollars = kWh * elec_rate + therms * gas_rate

        st.success("Estimated Annual Savings")
        a, b, c = st.columns(3)
        a.metric("Electric Energy", f"{kWh:,.0f} kWh/yr")
        b.metric("Gas Energy", f"{therms:,.0f} therms/yr")
        c.metric("Utility Savings", f"${dollars:,.0f}/yr")

        # Optional metrics if present
        extras = []
        if "Base_EUI_kBtuperSFperyr" in vals and pd.notnull(vals["Base_EUI_kBtuperSFperyr"]):
            extras.append(("Baseline EUI", f"{float(vals['Base_EUI_kBtuperSFperyr']):,.2f} kBtu/sf-yr"))
        if "CSW_EUI_kBtuperSFperyr" in vals and pd.notnull(vals["CSW_EUI_kBtuperSFperyr"]):
            extras.append(("CSW EUI", f"{float(vals['CSW_EUI_kBtuperSFperyr']):,.2f} kBtu/sf-yr"))
        if "Clg_Load_reduced_BtuhperSF" in vals and pd.notnull(vals["Clg_Load_reduced_BtuhperSF"]):
            extras.append(("Cooling Load Reduced", f"{float(vals['Clg_Load_reduced_BtuhperSF']):,.0f} Btuh/sf"))
        if "htg_load_reduced_BtuhperSF" in vals and pd.notnull(vals["htg_load_reduced_BtuhperSF"]):
            extras.append(("Heating Load Reduced", f"{float(vals['htg_load_reduced_BtuhperSF']):,.0f} Btuh/sf"))
        if "baseline_peak_cooling_BtuhperSF" in vals and pd.notnull(vals["baseline_peak_cooling_BtuhperSF"]):
            extras.append(("Baseline Peak Cooling", f"{float(vals['baseline_peak_cooling_BtuhperSF']):,.0f} Btuh/sf"))

        if extras:
            st.markdown("#### Additional Metrics")
            for label, value in extras:
                st.write(f"- **{label}:** {value}")

        # Summary
        st.markdown("#### Project Summary")
        st.json({
            "Project": st.session_state.project, "Contact": st.session_state.contact,
            "Company": st.session_state.company, "Email": st.session_state.email, "Phone": st.session_state.phone,
            "Location": f"{st.session_state.city}, {st.session_state.state}",
            "Building Size": building_size, "Building Type": building_type,
            "HVAC Type": hvac_type, "Fuel Type": fuel_type, "PTHP": pthp_val,
            "Operating Hours": hours, "Floor Area (sf)": floor_area,
            "CSW Area (sf)": csw_area, "Elec Rate": elec_rate, "Gas Rate": gas_rate
        })

# =========================================================
# Debug / Inspect
# =========================================================
with st.expander("Debug / Inspect"):
    st.write("Weather rows:", len(weather_df))
    st.dataframe(weather_df.head(30))
    st.write("Unique states:", len(weather_df["State"].unique()))
    st.write(weather_df.groupby("State")["City"].nunique().sort_values(ascending=False).head(20))
    st.write("Lookup rows:", len(lookup_df))
    st.dataframe(lookup_df.head(20))
