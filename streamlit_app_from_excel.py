
import streamlit as st
import pandas as pd
import numpy as np
import re
from pathlib import Path

# Path to the uploaded workbook in the session environment
DEFAULT_WORKBOOK_PATH = "/mnt/data/CSW Savings Calculator 2_0_0_Unlocked.xlsx"

st.set_page_config(page_title="CSW Savings Calculator (Excel-driven)", layout="wide")

st.title("CSW Savings Calculator — Excel-driven Streamlit App")
st.markdown(
    """
This app reads the master Excel workbook and uses the workbook's lookup tables to compute per-sf savings (heating kWh, cooling kWh, gas therms).
It tries to reproduce the spreadsheet logic by filtering the workbook's "Savings Lookup" table and interpolating hours when needed.
"""
)

@st.cache_data
def load_workbook_data(xlsx_path: str):
    p = Path(xlsx_path)
    if not p.exists():
        raise FileNotFoundError(f"Workbook not found at: {xlsx_path}")

    # Attempt to read likely sheets; be tolerant to small differences in sheet names
    xls = pd.ExcelFile(xlsx_path)

    def find_sheet(possible_names):
        for name in xls.sheet_names:
            lname = name.lower()
            for p in possible_names:
                if p.lower() in lname:
                    return name
        return None

    savings_sheet = find_sheet(["savings lookup", "savings", "savings_lookup", "savings lookup table"])
    weather_sheet = find_sheet(["weather", "weather information", "climate"])
    baseline_sheet = find_sheet(["baseline", "baseline eui", "baseline_eui"])

    df_savings = None
    df_weather = None
    df_baseline = None

    if savings_sheet:
        # set header row to the row that likely contains column labels; adjust if necessary
        df_savings = pd.read_excel(xls, sheet_name=savings_sheet)
        # Normalize column names to simple tokens
        df_savings.columns = [re.sub(r'[^0-9a-zA-Z_ ]', '', c).strip() for c in df_savings.columns]
    else:
        st.warning("Could not find a 'Savings Lookup' sheet in the workbook. Please check sheet names.")

    if weather_sheet:
        df_weather = pd.read_excel(xls, sheet_name=weather_sheet)
        df_weather.columns = [re.sub(r'[^0-9a-zA-Z_ ]', '', c).strip() for c in df_weather.columns]
    else:
        st.info("No weather sheet found. You can still proceed by entering HDD/CDD manually or selecting a city from a dropdown fallback.")

    if baseline_sheet:
        df_baseline = pd.read_excel(xls, sheet_name=baseline_sheet)
        df_baseline.columns = [re.sub(r'[^0-9a-zA-Z_ ]', '', c).strip() for c in df_baseline.columns]
    # Return whatever we loaded
    return df_savings, df_weather, df_baseline, savings_sheet, weather_sheet, baseline_sheet

def normalize_text(s):
    if pd.isna(s):
        return ""
    s = str(s).lower().strip()
    s = re.sub(r'[^a-z0-9 ]', '', s)
    s = re.sub(r'\s+', ' ', s)
    return s

def create_city_key(s):
    return normalize_text(s).replace(" ", "_")

def get_weather_for_city(df_weather, city_input):
    if df_weather is None or city_input is None:
        return None
    key = normalize_text(city_input)
    # try exact match on first column
    first_col = df_weather.columns[0]
    matches = df_weather[df_weather[first_col].astype(str).str.lower().str.contains(key, na=False)]
    if not matches.empty:
        # return the first match's row as dict
        row = matches.iloc[0].to_dict()
        return row
    return None

def map_building_type(user_choice):
    # Adjust mapping to match workbook codes where possible
    m = {
        "Office": "Office",
        "Hotel": "Hotel",
        "School": "SS",
        "Hospital": "Hosp",
        "Multi-family": "MF",
        "Multifamily": "MF",
        "Retail": "Retail",
        "Warehouse": "Warehouse"
    }
    return m.get(user_choice, user_choice)

def map_hvac_type(user_choice):
    # Map UI to likely sheet text. Update as necessary to match your workbook's exact values.
    m = {
        "Packaged Terminal (PTAC)": "PTAC",
        "Packaged Terminal Heat Pump (PTHP)": "PTHP",
        "Variable Air Volume (VAV)": "VAV",
        "Packaged VAV (PVAV Electric)": "PVAV_Elec",
        "Packaged VAV (PVAV Gas)": "PVAV_Gas",
        "Fan Coil Unit (FCU)": "FCU",
        "Other / Unknown": "Other"
    }
    return m.get(user_choice, user_choice)

def find_savings_row(df_savings, form):
    """
    Filters the workbook savings table to find matching rows and returns interpolated per-SF values for the requested hours.
    Expected columns in df_savings (case-insensitive): Base, CSW Type, Building Size, Building Type, HVAC System Type, Fuel Type, Hours, Heat, Cooling + Aux, Heat.1 (gas)
    """
    if df_savings is None:
        return None

    # Normalize columns for lookup (lowercase mapping)
    colmap = {c.lower(): c for c in df_savings.columns}
    # Try to map expected names
    def c(name_options):
        for opt in name_options:
            low = opt.lower()
            if low in colmap:
                return colmap[low]
        return None

    col_base = c(["Base", "base"])
    col_csw = c(["CSW Type", "csw type", "csw"])
    col_bsize = c(["Building Size", "building size", "size"])
    col_btype = c(["Building Type", "building type"])
    col_hvac = c(["HVAC System Type", "HVAC System", "HVAC"])
    col_fuel = c(["Fuel Type", "Fuel", "fuel type"])
    col_hours = c(["Hours", "Hours.1", "hours"])
    # energy columns (names may vary)
    possible_heat = [k for k in colmap.keys() if 'heat' in k and ('kwh' in k or 'heat'==k or 'heat ' in k)]
    col_heat = colmap[possible_heat[0]] if possible_heat else c(["Heat", "Heating"])
    possible_cool = [k for k in colmap.keys() if 'cool' in k or 'cooling' in k]
    col_cool = colmap[possible_cool[0]] if possible_cool else c(["Cooling + Aux", "Cooling"])
    possible_gas = [k for k in colmap.keys() if 'therm' in k or 'gas' in k or 'heat.1' in k]
    col_gas = colmap[possible_gas[0]] if possible_gas else c(["Heat.1", "Gas"])

    # Build filter values
    base_val = "Single" if "single" in form["existing_window_type"].lower() else "Double"
    csw_val = form["secondary_window_type"]
    bsize = "Large" if form["floor_area"] >= 200000 else "Mid"
    btype = map_building_type(form["building_type"])
    hvac = map_hvac_type(form["hvac_system"])
    fuel = "Electric" if "electric" in form["heating_fuel"].lower() else "Natural Gas"
    hours = float(form["operation_hours"])

    # Prepare a filtered dataframe
    df = df_savings.copy()
    # Ensure we have the matching columns
    checks = [col_base, col_csw, col_bsize, col_btype, col_hvac, col_fuel, col_hours]
    for ch in checks:
        if ch is None:
            # If a required column is missing, we cannot match; return None and the caller should handle fallback
            return None

    mask = (
        df[col_base].astype(str).str.strip().str.lower() == base_val.lower()
    ) & (
        df[col_csw].astype(str).str.strip().str.lower() == csw_val.lower()
    ) & (
        df[col_bsize].astype(str).str.strip().str.lower() == bsize.lower()
    ) & (
        df[col_btype].astype(str).str.strip().str.lower() == str(btype).lower()
    ) & (
        df[col_hvac].astype(str).str.strip().str.lower().str.replace(' ', '_') == str(hvac).lower().replace(' ', '_')
    ) & (
        df[col_fuel].astype(str).str.strip().str.lower().str.contains(fuel.lower())
    )

    candidates = df[mask].copy()
    if candidates.empty:
        # try relaxing HVAC matching (replace 'other' or simplify)
        mask2 = mask.copy()
        if col_hvac in df.columns:
            mask2 = mask.copy() & df[col_hvac].astype(str).str.strip().str.lower().str.contains('pvav|vav|ptac|pthp|fcu|pack', na=False)
            candidates = df[mask2].copy()

    if candidates.empty:
        return None

    # Convert hours column to numeric
    candidates[col_hours] = pd.to_numeric(candidates[col_hours], errors='coerce')
    candidates = candidates.dropna(subset=[col_hours])
    if candidates.empty:
        return None

    # If exact match on hours exists, take that row
    if hours in candidates[col_hours].values:
        r = candidates[candidates[col_hours] == hours].iloc[0]
        return {
            "hours_used": hours,
            "heat_kwh_sf": float(r[col_heat]) if col_heat in r else 0.0,
            "cool_kwh_sf": float(r[col_cool]) if col_cool in r else 0.0,
            "gas_therm_sf": float(r[col_gas]) if col_gas in r else 0.0,
            "provenance": {"matched_row_index": int(r.name), "interpolated": False}
        }
    # Otherwise, interpolate linearly between the nearest hours
    candidates = candidates.sort_values(col_hours)
    lower = candidates[candidates[col_hours] <= hours].tail(1)
    upper = candidates[candidates[col_hours] >= hours].head(1)
    if not lower.empty and not upper.empty and lower.index[0] != upper.index[0]:
        low = lower.iloc[0]; up = upper.iloc[0]
        h_low = float(low[col_hours]); h_up = float(up[col_hours])
        frac = (hours - h_low) / (h_up - h_low) if (h_up - h_low) != 0 else 0.0
        def interp(colname):
            v_low = float(low[colname]) if colname in low and pd.notna(low[colname]) else 0.0
            v_up  = float(up[colname])  if colname in up  and pd.notna(up[colname])  else v_low
            return v_low + (v_up - v_low) * frac
        return {
            "hours_used": hours,
            "heat_kwh_sf": interp(col_heat),
            "cool_kwh_sf": interp(col_cool),
            "gas_therm_sf": interp(col_gas),
            "provenance": {"lower_idx": int(low.name), "upper_idx": int(up.name), "interpolated": True, "frac": frac}
        }
    else:
        # snap to nearest hours row
        idx = (candidates[col_hours] - hours).abs().idxmin()
        r = candidates.loc[idx]
        return {
            "hours_used": float(r[col_hours]),
            "heat_kwh_sf": float(r[col_heat]) if col_heat in r else 0.0,
            "cool_kwh_sf": float(r[col_cool]) if col_cool in r else 0.0,
            "gas_therm_sf": float(r[col_gas]) if col_gas in r else 0.0,
            "provenance": {"matched_row_index": int(r.name), "interpolated": False, "snapped": True}
        }

def calculate_totals(savings_row, window_area_sf, elec_rate=0.12, gas_rate=1.0):
    # per-sf savings times window area -> totals. Preserve sign for gas.
    heat_kwh = savings_row["heat_kwh_sf"] * window_area_sf
    cool_kwh = savings_row["cool_kwh_sf"] * window_area_sf
    gas_therms = savings_row["gas_therm_sf"] * window_area_sf  # can be negative

    elec_cost = (heat_kwh + cool_kwh) * float(elec_rate)
    gas_cost = gas_therms * float(gas_rate)

    return {
        "heat_kwh": heat_kwh,
        "cool_kwh": cool_kwh,
        "gas_therms": gas_therms,
        "elec_cost": elec_cost,
        "gas_cost": gas_cost,
        "total_elec_kwh": heat_kwh + cool_kwh,
        "total_cost": elec_cost + gas_cost
    }

def energy_intensity_kbtu_per_sf(total_kwh, total_therms, floor_area):
    # conversions
    KWH_TO_BTU = 3412.142
    THERM_TO_BTU = 100000
    total_btu = total_kwh * KWH_TO_BTU + total_therms * THERM_TO_BTU
    kbtu = total_btu / 1000.0
    if floor_area <= 0:
        return np.nan
    return kbtu / float(floor_area)

# --- Load workbook data ---
xlsx_path = st.text_input("Path to Excel workbook", value=DEFAULT_WORKBOOK_PATH)
try:
    DF_SAVINGS, DF_WEATHER, DF_BASELINE, s_sheet, w_sheet, b_sheet = load_workbook_data(xlsx_path)
except FileNotFoundError as e:
    st.error(str(e))
    st.stop()

st.sidebar.header("Project inputs")
building_type = st.sidebar.selectbox("Building type", options=["Office", "Hotel", "School", "Hospital", "Multi-family", "Retail", "Warehouse"], index=0)
floor_area = st.sidebar.number_input("Floor area (ft²)", min_value=100.0, value=50000.0, step=100.0)
window_area = st.sidebar.number_input("Window area (ft²) - area of windows replaced", min_value=1.0, value=5000.0, step=10.0)
existing_window = st.sidebar.selectbox("Existing window (Base)", options=["Single", "Double"], index=0)
secondary_window = st.sidebar.selectbox("Secondary window (CSW Type)", options=["Single", "Double"], index=0)
hvac = st.sidebar.selectbox("HVAC system", options=["Packaged VAV (PVAV Electric)", "Packaged VAV (PVAV Gas)", "Variable Air Volume (VAV)", "Packaged Terminal (PTAC)", "Packaged Terminal Heat Pump (PTHP)", "Fan Coil Unit (FCU)", "Other / Unknown"], index=0)
heating_fuel = st.sidebar.selectbox("Heating fuel", options=["Electric", "Natural Gas"], index=0)
elec_rate = st.sidebar.number_input("Electricity cost ($/kWh)", min_value=0.0, value=0.12, step=0.01)
gas_rate = st.sidebar.number_input("Gas cost ($/therm)", min_value=0.0, value=1.0, step=0.01)

# operation hours choices: try to show available hour buckets from DF_SAVINGS if present
hours_options = [2080, 2912, 8760]  # defaults
if DF_SAVINGS is not None:
    # detect column with hours label
    nums = []
    for c in DF_SAVINGS.columns:
        try:
            col_vals = pd.to_numeric(DF_SAVINGS[c], errors='coerce').dropna().unique().tolist()
            # if small set with typical hours values, grab them
            for v in col_vals:
                if int(v) in [2080, 2912, 8760]:
                    nums.append(int(v))
        except Exception:
            continue
    if nums:
        hours_options = sorted(list(set(nums)))

operation_hours = st.sidebar.selectbox("Operation hours (choose or let app interpolate)", options=hours_options, index=0)

st.sidebar.markdown("---")
st.sidebar.write("Optional: pick a city (to use HDD/CDD from workbook weather sheet)")
city_choice = None
if DF_WEATHER is not None:
    first_col = DF_WEATHER.columns[0]
    city_list = DF_WEATHER[first_col].astype(str).tolist()
    city_choice = st.sidebar.selectbox("City (from workbook weather sheet)", options=["-- none --"] + city_list, index=0)
    if city_choice and city_choice != "-- none --":
        city_row = get_weather_for_city(DF_WEATHER, city_choice)
    else:
        city_row = None
else:
    city_row = None

if st.sidebar.button("Calculate"):
    form = {
        "building_type": building_type,
        "floor_area": floor_area,
        "window_area": window_area,
        "existing_window_type": existing_window,
        "secondary_window_type": secondary_window,
        "hvac_system": hvac,
        "heating_fuel": heating_fuel,
        "operation_hours": operation_hours
    }

    savings_row = find_savings_row(DF_SAVINGS, form)
    if savings_row is None:
        st.error("Could not find a matching savings row in the workbook for the chosen inputs. Try a different HVAC selection or check the workbook's 'Savings Lookup' table for expected codes/values.")
    else:
        totals = calculate_totals(savings_row, window_area, elec_rate, gas_rate)
        kbtu_per_sf = energy_intensity_kbtu_per_sf(totals["total_elec_kwh"], totals["gas_therms"], floor_area)

        c1, c2 = st.columns([2,3])
        with c1:
            st.subheader("Per-SF savings (from workbook)")
            st.metric("Heating kWh / ft²", f"{savings_row['heat_kwh_sf']:.6f}")
            st.metric("Cooling kWh / ft²", f"{savings_row['cool_kwh_sf']:.6f}")
            st.metric("Gas therms / ft²", f"{savings_row['gas_therm_sf']:.6f}")
            st.write("Lookup provenance:")
            st.json(savings_row["provenance"])
        with c2:
            st.subheader("Annual totals")
            st.write(f"Window area used: {window_area:.1f} ft² — Floor area: {floor_area:.1f} ft²")
            st.write(f"Electricity saved (kWh): {totals['total_elec_kwh']:.2f}")
            st.write(f"Gas change (therms): {totals['gas_therms']:.3f} (note: can be negative if gas increases)")
            st.write(f"Electricity cost change ($): {totals['elec_cost']:.2f}")
            st.write(f"Gas cost change ($): {totals['gas_cost']:.2f}")
            st.write(f"Total annual cost change ($): {totals['total_cost']:.2f}")
            st.write(f"Energy intensity change (kBtu / ft²-yr): {kbtu_per_sf:.4f}")

        st.markdown("#### Debug / verification")
        st.write("Show the first few matching candidates (unfiltered) used for lookup to aid debugging:")
        # show filtered candidates for debug
        # create same mask as in find_savings_row to display candidates
        # we'll reuse function by re-filtering simple way
        st.write("If you want, open the workbook and check the 'Savings Lookup' sheet to confirm the exact labels used for HVAC, Building Type, etc.")
        if city_row is not None:
            st.write("Weather row (from workbook):")
            st.json(city_row)
