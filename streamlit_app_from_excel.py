import streamlit as st
import pandas as pd

# ----------------------------
# Load Workbook
# ----------------------------
@st.cache_data
def load_workbook():
    # Load Excel file directly from repo root
    return pd.ExcelFile("CSW Savings Calculator 2_0_0_Unlocked.xlsx")

# ----------------------------
# App Layout
# ----------------------------
st.set_page_config(page_title="CSW Savings Calculator", layout="centered")
st.title("Commercial Secondary Window (CSW) Savings Calculator")

st.markdown(
    "This tool estimates potential energy and cost savings from installing "
    "secondary windows in commercial buildings."
)

# ----------------------------
# Step 1: Project Information
# ----------------------------
st.header("Step 1: Project Information")
project_name = st.text_input("Project Name")
contact_name = st.text_input("Contact Name")
contact_email = st.text_input("Contact Email")

# ----------------------------
# Step 2: Building Characteristics
# ----------------------------
st.header("Step 2: Building Information")

state = st.selectbox("Select State", ["-- Select --", "CA", "NY", "TX", "IL"])
city = st.selectbox("Select City", ["-- Select State First --"])

floor_area = st.number_input("Building Area (sqft)", min_value=1000, step=500)
num_floors = st.number_input("Number of Floors", min_value=1, step=1)
op_hours = st.number_input("Annual Operating Hours", min_value=1000, step=100)

# ----------------------------
# Step 3: Envelope & Systems
# ----------------------------
st.header("Step 3: Envelope & Systems")

existing_window = st.selectbox("Existing Window Type", ["Single-pane", "Dual-pane"])
hvac_type = st.selectbox("HVAC System Type", ["VAV w/ Reheat", "CAV", "Heat Pump"])
heating_fuel = st.selectbox("Heating Fuel", ["Natural Gas", "Electricity"])
csw_type = st.selectbox("Secondary Window Type", ["Single-pane", "Dual-pane"])
csw_area = st.number_input("Sqft of CSW Installed", min_value=100, step=50)

# ----------------------------
# Step 4: Utility Rates
# ----------------------------
st.header("Step 4: Utility Rates")
elec_rate = st.number_input("Electricity Rate ($/kWh)", min_value=0.01, step=0.01)
gas_rate = st.number_input("Gas Rate ($/therm)", min_value=0.01, step=0.01)

# ----------------------------
# Step 5: Results (Demo Hardcoded)
# ----------------------------
st.header("Step 5: Estimated Results")

if st.button("Run Calculation"):
    # Demo math (replace with regression logic later)
    baseline_energy = floor_area * 10  # fake baseline (kWh)
    savings = baseline_energy * 0.15   # assume 15% savings
    annual_savings = savings * elec_rate

    st.success(f"Estimated Energy Savings: {savings:,.0f} kWh/year")
    st.success(f"Estimated Utility Savings: ${annual_savings:,.0f}/year")
    st.info("Peak Load Reduction: ~12% (demo assumption)")

# ----------------------------
# Debug Section (Optional)
# ----------------------------
with st.expander("Debug Info"):
    try:
        wb = load_workbook()
        st.write("Workbook Sheets Found:", wb.sheet_names)
    except Exception as e:
        st.error(f"Could not load workbook: {e}")
