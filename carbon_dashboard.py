import streamlit as st
import openpyxl
import warnings

# Suppress warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.filterwarnings('ignore', category=DeprecationWarning)

# Set page configuration
st.set_page_config(page_title="Carbon Credit Dashboard", layout="wide")
st.title("Carbon Credit Project Financing Management Dashboard")

# Excel file path
excel_path = r"C:\\Users\\adhar\\Downloads\\ZE_ICAR (4).xlsx"

# Function to safely convert to number
def convert_to_number(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0

# Function to load Excel data
@st.cache_data
def load_excel_data(file_path):
    data = {}
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet_names = wb.sheetnames

        unified_sheet_name = next((name for name in sheet_names if "unified" in name.lower()), None)
        if not unified_sheet_name:
            st.error("Could not find 'UNIFIED USER INPUT' sheet in Excel file")
            return None

        unified_sheet = wb[unified_sheet_name]
        data["unified"] = {
            "SINK": unified_sheet['B1'].value or "Biofertilizers",
            "Emitter Unit": unified_sheet['B2'].value or "Hectare",
            "Carbon Credits Per Year": convert_to_number(unified_sheet['B3'].value),
            "Sink Size": convert_to_number(unified_sheet['B4'].value),
            "Total Project Cost": convert_to_number(unified_sheet['E1'].value),
            "Faire Trade Price": convert_to_number(unified_sheet['E2'].value),
            "Total CC Generated": convert_to_number(unified_sheet['E3'].value),
            "Expected Price/CC": convert_to_number(unified_sheet['E4'].value)
        }

        data["sink_data"] = {}
        for sheet_name in sheet_names:
            if sheet_name.lower() != unified_sheet_name.lower():
                sheet = wb[sheet_name]
                data["sink_data"][sheet_name.strip()] = {
                    "CC per year/ha": convert_to_number(sheet['B4'].value),
                    "Total Cost": convert_to_number(sheet['F6'].value)
                }

        sink_sheets = [name.strip() for name in sheet_names if "unified" not in name.lower()]
        data["sink_options"] = sink_sheets

        emitter_map = {sheet: ("Unit" if ("solar" in sheet.lower() or "irrigation" in sheet.lower()) else "Hectare") for sheet in sink_sheets}
        data["emitter_map"] = emitter_map

        return data
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None

# Function to calculate derived values
def calculate_values(sink, sink_size, excel_data):
    if not excel_data or not sink:
        return {}

    sink_data = excel_data.get("sink_data", {}).get(sink, {})

    carbon_credits_per_year = sink_data.get("CC per year/ha", 0.0)
    total_cc_generated = carbon_credits_per_year * sink_size
    total_project_cost = sink_data.get("Total Cost", 0.0) * sink_size
    fair_trade_price_per_cc = total_project_cost / total_cc_generated if total_cc_generated > 0 else 0.0
    expected_price_cc = fair_trade_price_per_cc * 1.5

    return {
        "Emitter Unit": excel_data["emitter_map"].get(sink, "Hectare"),
        "Carbon Credits Per Year": carbon_credits_per_year,
        "Total CC Generated": total_cc_generated,
        "Total Project Cost": total_project_cost,
        "Faire Trade Price": fair_trade_price_per_cc,
        "Expected Price/CC": expected_price_cc
    }

# Load Excel data
excel_data = load_excel_data(excel_path)

default_values = excel_data.get("unified", {}) if excel_data else {}
sink_options = excel_data.get("sink_options", ["AWD in Paddy", "Crop Residue Management", "Biofertilizers", "SolarIrrigation Efficiency"])

col1, col2 = st.columns(2)

with col1:
    with st.container(border=True):
        st.markdown("<div style='background-color:#e6f0ff; padding:10px; border-radius:10px'>", unsafe_allow_html=True)
        st.subheader("Project Configuration", divider="blue")
        sink = st.selectbox(
            "SINK",
            options=sink_options,
            index=sink_options.index(default_values.get("SINK", "Biofertilizers")) if default_values.get("SINK") in sink_options else 0,
            key="sink"
        )
        sink_size = st.number_input(
            "Sink Size (in units)/year",
            min_value=0.0,
            value=float(default_values.get("Sink Size", 200000.0)),
            format="%.0f",
            step=1.0,
            key="sink_size"
        )
        calculated_values = calculate_values(sink, sink_size, excel_data)
        st.text_input("Emitter Unit", value=calculated_values.get("Emitter Unit", "Hectare"), disabled=True)
        st.markdown("</div>", unsafe_allow_html=True)

with col2:
    with st.container(border=True):
        st.markdown("<div style='background-color:#e6ffe6; padding:10px; border-radius:10px'>", unsafe_allow_html=True)
        st.subheader("Financial Metrics", divider="green")
        st.number_input("Carbon Credits Per Year", value=calculated_values.get("Carbon Credits Per Year", 0.0), format="%.2f", disabled=True)
        st.number_input("Total CC Generated", value=calculated_values.get("Total CC Generated", 0.0), format="%.2f", disabled=True)
        st.number_input("Total Project Cost ($)", value=calculated_values.get("Total Project Cost", 0.0), format="%.2f", disabled=True)
        st.number_input("Fair Trade Price per CC ($)", value=calculated_values.get("Faire Trade Price", 0.0), format="%.2f", disabled=True)
        st.number_input("Expected price / CC ($)", value=calculated_values.get("Expected Price/CC", 0.0), format="%.2f", disabled=True)
        st.markdown("</div>", unsafe_allow_html=True)

# Display current config
# current_config = {
#     "SINK": sink,
#     "Sink Size": sink_size,
#     **calculated_values
# }
# st.subheader("Current Configuration")
# st.json(current_config)