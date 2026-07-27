import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, date
from io import BytesIO

# Configure page
st.set_page_config(page_title="Harvest Report - Data Collection", layout="wide")

# Title
st.title("🦐 Harvest Report - Data Collection")
st.subheader("KMN Aqua Services")
st.markdown("---")

# ------------------------------------------------------------------
# Google Sheets connection
# ------------------------------------------------------------------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

WORKSHEET_NAME = "Harvest Data"  # tab name inside the Google Sheet

COLUMNS = [
    "Timestamp", "Customer", "Farm Name", "Zone", "Area", "Species Culture",
    "Number of Ponds Harvest", "ABW", "Harvest Date", "Harvest Type",
    "Harvest In KG", "Harvest Reason", "Remark", "Assigned Marketing Manager",
    "Technician",
]


@st.cache_resource
def get_gsheet_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    return gspread.authorize(creds)


@st.cache_resource
def get_worksheet():
    client = get_gsheet_client()
    sheet = client.open_by_url(st.secrets["sheet"]["url"])
    try:
        ws = sheet.worksheet(WORKSHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = sheet.add_worksheet(title=WORKSHEET_NAME, rows=1000, cols=len(COLUMNS))
        ws.append_row(COLUMNS)
        return ws

    # Make sure the header row is exactly COLUMNS (fixes sheets that were
    # created manually, or have stray blank header cells).
    first_row = ws.row_values(1)
    if first_row != COLUMNS:
        ws.update("A1", [COLUMNS])
    return ws


def load_data() -> pd.DataFrame:
    ws = get_worksheet()
    try:
        records = ws.get_all_records(expected_headers=COLUMNS)
    except Exception:
        # Fallback: header row is malformed/empty (e.g. brand new sheet).
        # Rewrite it cleanly and retry.
        values = ws.get_all_values()
        if not values or values[0] != COLUMNS:
            ws.update("A1", [COLUMNS])
        records = ws.get_all_records(expected_headers=COLUMNS)
    if not records:
        return pd.DataFrame(columns=COLUMNS)
    return pd.DataFrame(records)


def save_row(new_row: dict):
    ws = get_worksheet()
    if not ws.get_all_values():
        ws.append_row(COLUMNS)
    ws.append_row([str(new_row.get(col, "")) for col in COLUMNS])


def clear_all_data():
    ws = get_worksheet()
    ws.clear()
    ws.append_row(COLUMNS)


# ------------------------------------------------------------------
# Customer list (still loaded from Excel bundled in the repo)
# ------------------------------------------------------------------
@st.cache_data
def load_customer_data():
    df = pd.read_excel("Customer List.xlsx")
    return df


customer_df = load_customer_data()
customers = customer_df["Customer ID"].astype(str) + " - " + customer_df["Customer Name"]
unique_customers = customers.unique().tolist()
farms = customer_df["Farm Name"].dropna().unique().tolist()
areas = customer_df["Area"].dropna().unique().tolist()

# Define options
ZONE_OPTIONS = ["Chilaw", "Puttalam", "Middle Zone", "Batticaloa"]
SPECIES_CULTURE = ["Vannamei", "Monodon", "Other"]
HARVEST_TYPE_OPTIONS = ["Full Harvest", "Partial", "Part"]
HARVEST_REASON_OPTIONS = [
    "Reached the Target Size",
    "Reached the Target Price",
    "WSSP Desease",
    "EHP Desease",
    "Low Survival Rate",
    "Stoped Growing",
    "Water Quality Issue",
    "Environmental Reasons",
    "Over Stockings",
    "Lower Stockings of feeds",
    "Equipment's issues",
    "Other",
    "O2 Drop"
]
MARKETING_MANAGER_OPTIONS = ["Mr. Ansadeen", "Mr. Jagath", "Mr. Dilip"]
TECHNICIAN_OPTIONS = ["Mr. Vishmika", "Mr. Ashen", "Mr. Janaka", "Mr. Shashika", "Mr. Janushan"]

# Initialize session state
if 'submission_count' not in st.session_state:
    st.session_state.submission_count = 0
if 'selected_customer' not in st.session_state:
    st.session_state.selected_customer = unique_customers[0] if unique_customers else ""
if 'selected_farm' not in st.session_state:
    st.session_state.selected_farm = farms[0] if farms else ""
if 'selected_zone' not in st.session_state:
    st.session_state.selected_zone = ZONE_OPTIONS[0]
if 'selected_area' not in st.session_state:
    st.session_state.selected_area = areas[0] if areas else ""
if 'selected_marketing_manager' not in st.session_state:
    st.session_state.selected_marketing_manager = MARKETING_MANAGER_OPTIONS[0]
if 'selected_technician' not in st.session_state:
    st.session_state.selected_technician = TECHNICIAN_OPTIONS[0] if TECHNICIAN_OPTIONS else ""

# Get indices for pre-selected values
customer_index = unique_customers.index(st.session_state.selected_customer) if st.session_state.selected_customer in unique_customers else 0
farm_index = farms.index(st.session_state.selected_farm) if st.session_state.selected_farm in farms else 0
zone_index = ZONE_OPTIONS.index(st.session_state.selected_zone) if st.session_state.selected_zone in ZONE_OPTIONS else 0
area_index = areas.index(st.session_state.selected_area) if st.session_state.selected_area in areas else 0
mm_index = MARKETING_MANAGER_OPTIONS.index(st.session_state.selected_marketing_manager) if st.session_state.selected_marketing_manager in MARKETING_MANAGER_OPTIONS else 0
technician_index = TECHNICIAN_OPTIONS.index(st.session_state.selected_technician) if st.session_state.selected_technician in TECHNICIAN_OPTIONS else 0

# ------------------------------------------------------------------
# Form
# ------------------------------------------------------------------
with st.form(f"harvest_report_form_{st.session_state.submission_count}"):
    st.subheader("📋 Enter Harvest Report Data")

    col1, col2 = st.columns(2)
    with col1:
        customer = st.selectbox("Customer *", unique_customers, index=customer_index)
    with col2:
        farm = st.selectbox("Farm Name", farms, index=farm_index if farms else None)

    col3, col4 = st.columns(2)
    with col3:
        zone = st.selectbox("Zone *", ZONE_OPTIONS, index=zone_index)
    with col4:
        area = st.selectbox("Area *", areas if areas else [""], index=area_index)

    species = st.selectbox("Species Culture *", SPECIES_CULTURE)

    col5, col6 = st.columns(2)
    with col5:
        ponds_harvest = st.number_input("Number of ponds harvest *", min_value=0, step=1)
    with col6:
        abw = st.number_input("ABW *", min_value=0.0, step=0.1)

    col7, col8 = st.columns(2)
    with col7:
        harvest_date = st.date_input("Harvest Date *", value=date.today())
    with col8:
        harvest_type = st.selectbox("Harvest Type *", HARVEST_TYPE_OPTIONS)

    harvest_kg = st.number_input("Harvest In KG *", min_value=0.0, step=0.1)

    st.markdown("#### 🦐 Harvest Reason *")
    harvest_reason = st.multiselect("Select applicable reasons", HARVEST_REASON_OPTIONS)

    st.markdown("#### 📝 Remark")
    remark = st.text_area("Additional remarks or notes", placeholder="Enter any additional information", height=80)

    st.markdown("#### 👔 Assigned Marketing Manager *")
    marketing_manager = st.selectbox("Select marketing manager", MARKETING_MANAGER_OPTIONS, index=mm_index, key="mm")

    st.markdown("#### 👤 Technician *")
    technician = st.selectbox("Select technician", TECHNICIAN_OPTIONS, index=technician_index, key="tech")

    submitted = st.form_submit_button("✅ Submit Data", width='stretch')

    if submitted:
        if not customer or not zone or not area or not species or not harvest_type or not harvest_reason or not marketing_manager or not technician:
            st.error("❌ Please fill in all required fields (marked with *)")
        else:
            new_row = {
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Customer": customer,
                "Farm Name": farm,
                "Zone": zone,
                "Area": area,
                "Species Culture": species,
                "Number of Ponds Harvest": ponds_harvest,
                "ABW": abw,
                "Harvest Date": harvest_date.strftime("%Y-%m-%d"),
                "Harvest Type": harvest_type,
                "Harvest In KG": harvest_kg,
                "Harvest Reason": ", ".join(harvest_reason) if harvest_reason else "",
                "Remark": remark,
                "Assigned Marketing Manager": marketing_manager,
                "Technician": technician,
            }

            try:
                save_row(new_row)
            except Exception as e:
                st.error(f"❌ Could not save to Google Sheet: {e}")
            else:
                st.session_state.selected_customer = customer
                st.session_state.selected_farm = farm
                st.session_state.selected_zone = zone
                st.session_state.selected_area = area
                st.session_state.selected_marketing_manager = marketing_manager
                st.session_state.selected_technician = technician
                st.session_state.submission_count += 1

                st.success("✅ Data saved successfully!")
                import time
                time.sleep(1)
                st.rerun()

# ------------------------------------------------------------------
# Display saved data
# ------------------------------------------------------------------
st.markdown("---")
st.subheader("📊 Saved Data")

try:
    df = load_data()
except Exception as e:
    st.error(f"❌ Could not load data from Google Sheet: {e}")
    df = pd.DataFrame(columns=COLUMNS)

if len(df) > 0:
    st.write(f"Total records: **{len(df)}**")

    st.dataframe(df, width='stretch', height=400)

    col1, col2 = st.columns(2)

    with col1:
        csv = df.to_csv(index=False)
        st.download_button(
            label="📥 Download as CSV",
            data=csv,
            file_name=f"harvest_report_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            width='stretch',
        )

    with col2:
        excel_buffer = BytesIO()
        df.to_excel(excel_buffer, index=False, sheet_name="Harvest Report Data")
        excel_buffer.seek(0)
        st.download_button(
            label="📥 Download as Excel",
            data=excel_buffer,
            file_name=f"harvest_report_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch',
        )
else:
    st.info("ℹ️ No data saved yet. Fill out the form above to get started!")

st.markdown("<p style='text-align: center; color: gray;'>KMN Aqua Services - Harvest Report Monitoring System</p>", unsafe_allow_html=True)
