import base64
from datetime import datetime, date
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import requests
import streamlit as st

# ======================================================
# Page & Layout Configuration
# ======================================================
st.set_page_config(page_title="Purchased Invoice Entry", page_icon="🧾", layout="wide")

st.title("🧾 Purchased Invoice Entry System")

# Google Apps Script Web App Endpoint & Sheet ID
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyOhnkbEa7JHQk49LAH1xg4PXMzeMVwPDfesMbOxeOvsDjcOjL7dEXYBGRymn-Sx6MKkQ/exec"
SPREADSHEET_ID = "1PCJ3BWAj6Wz1N-55XpuWltvfvYd7KD1q194D3N7MzIg"
WORKSHEET_NAME = "Purchase_Invoice_Log"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

COMPANY_LIST = [
    "Select Company",
    "Shivam Enterprise",
    "P.S.INDUSTRIES",
    "Others"
]

UNIT_OPTIONS = ["per kg", "per pcs", "per gross", "per ream", "per box", "per pkt"]

# ======================================================
# Database Connector (gspread)
# ======================================================
@st.cache_resource(ttl=3600)
def connect_purchase_log_sheet():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID).worksheet(WORKSHEET_NAME)

purchase_sheet = connect_purchase_log_sheet()

# ======================================================
# Apps Script Google Drive Upload Helper
# ======================================================
def upload_scan_to_drive_via_appscript(file_bytes, file_name, mime_type, company_name):
    encoded_file = base64.b64encode(file_bytes).decode("utf-8")
    payload = {
        "action": "upload_purchase_invoice",
        "company": company_name.strip(),
        "fileName": file_name,
        "mimeType": mime_type,
        "fileData": encoded_file,
    }
    try:
        response = requests.post(WEB_APP_URL, json=payload, timeout=30)
        if response.status_code == 200:
            res_json = response.json()
            if res_json.get("status") == "success":
                return res_json.get("link", "")
            else:
                st.error(f"Apps Script Error: {res_json.get('message', 'Unknown error')}")
        else:
            st.error(f"HTTP Error {response.status_code}: Could not reach Apps Script Endpoint.")
    except Exception as e:
        st.error(f"Failed to connect to Apps Script: {e}")
    return None

# ======================================================
# Session State Setup
# ======================================================
if "form_key" not in st.session_state:
    st.session_state.form_key = 0
if "upload_modal_data" not in st.session_state:
    st.session_state.upload_modal_data = None

key_suffix = st.session_state.form_key

# ======================================================
# Primary Navigation Tabs
# ======================================================
tab1, tab2 = st.tabs(["➕ New Purchase Entry", "📋 Purchase History Log"])

# ------------------------------------------------------
# TAB 1: ENTRY FORM
# ------------------------------------------------------
with tab1:
    st.subheader("Enter Purchased Invoice Details")

    col1, col2 = st.columns(2)
    with col1:
        inv_date = st.date_input("Invoice Date *", value=date.today(), key=f"inv_date_{key_suffix}")
        
        selected_company = st.selectbox("Company Name *", COMPANY_LIST, key=f"comp_{key_suffix}")
        if selected_company == "Others":
            company_name = st.text_input("Enter Manual Company Name *", key=f"manual_comp_{key_suffix}")
        else:
            company_name = selected_company

        inv_number = st.text_input("Invoice Number *", placeholder="e.g. INV/2026/001", key=f"inv_num_{key_suffix}")
        product_name = st.text_input("Product Name *", placeholder="e.g. Maplitho Paper", key=f"prod_name_{key_suffix}")
        unit = st.selectbox("Unit *", UNIT_OPTIONS, key=f"unit_{key_suffix}")

    with col2:
        quantity = st.number_input("Quantity *", min_value=0.0, value=0.0, step=1.0, format="%.2f", key=f"qty_{key_suffix}")
        rate = st.number_input("Rate *", min_value=0.0, value=0.0, step=0.5, format="%.2f", key=f"rate_{key_suffix}")
        
        # 1. Automatic Taxable Value calculation (Quantity * Rate)
        taxable_value = round(quantity * rate, 2)
        st.number_input("Taxable Value (₹) [Qty × Rate]", value=float(taxable_value), disabled=True, key=f"tax_val_{key_suffix}")

        # 2. GST Percentage Selection & Live Calculation
        gst_pct = st.selectbox("GST Percentage (%) *", GST_PERCENTAGE_OPTIONS, index=1, key=f"gst_pct_{key_suffix}")
        
        half_gst_pct = gst_pct / 2.0
        sgst_val = round(taxable_value * (half_gst_pct / 100.0), 2)
        cgst_val = round(taxable_value * (half_gst_pct / 100.0), 2)
        total_gst = sgst_val + cgst_val

        # Display live GST metrics
        c_m1, c_m2, c_m3 = st.columns(3)
        c_m1.metric(f"SGST ({half_gst_pct}%)", f"₹ {sgst_val:,.2f}")
        c_m2.metric(f"CGST ({half_gst_pct}%)", f"₹ {cgst_val:,.2f}")
        c_m3.metric(f"Total GST ({gst_pct}%)", f"₹ {total_gst:,.2f}")

        # 3. Round Off Adjustment
        round_off = st.number_input("Round Off (+/-)", value=0.0, step=0.01, format="%.2f", key=f"round_{key_suffix}")
        
        # Final Invoice Value Calculation
        calc_inv_val = round(taxable_value + total_gst + round_off, 2)
        inv_value = st.number_input("Final Invoice Value (₹) *", min_value=0.0, value=float(calc_inv_val), step=1.0, format="%.2f", key=f"inv_val_{key_suffix}")

    remark = st.text_area("Remark / Notes", placeholder="Supplier details, batch numbers, or terms...", key=f"rem_{key_suffix}")

    # Initial Validation & Submission
    if st.button("Proceed to Upload Scan ➔", type="primary", use_container_width=True):
        if selected_company == "Select Company" or not company_name.strip():
            st.warning("Please select or enter a valid Company Name.")
        elif not inv_number.strip():
            st.warning("Please enter an Invoice Number.")
        elif not product_name.strip():
            st.warning("Please enter a Product Name.")
        elif quantity <= 0:
            st.warning("Quantity must be greater than zero.")
        elif rate <= 0:
            st.warning("Rate must be greater than zero.")
        else:
            # Store validated form data in session state and trigger upload modal
            st.session_state.upload_modal_data = {
                "inv_date": inv_date.strftime("%d/%m/%Y"),
                "company_name": company_name.strip(),
                "inv_number": inv_number.strip(),
                "product_name": product_name.strip(),
                "quantity": quantity,
                "rate": rate,
                "unit": unit,
                "taxable_value": taxable_value,
                "gst_pct": gst_pct,
                "total_gst": total_gst,
                "round_off": round_off,
                "inv_value": inv_value,
                "remark": remark.strip()
            }
            st.rerun()

# ------------------------------------------------------
# MODAL DIALOG FOR SCAN COPY UPLOAD
# ------------------------------------------------------
if st.session_state.upload_modal_data is not None:
    data_entry = st.session_state.upload_modal_data

    @st.dialog("📸 Upload Purchased Invoice Scan Copy")
    def open_scan_upload_modal():
        st.info(f"**Invoice No:** `{data_entry['inv_number']}` | **Company:** `{data_entry['company_name']}` | **Total Value:** `₹ {data_entry['inv_value']:,.2f}`")
        
        input_mode = st.radio(
            "Select Scan Source",
            ["📷 Camera Capture", "📁 Upload File from Device"],
            horizontal=True
        )

        file_data = None
        if input_mode == "📷 Camera Capture":
            file_data = st.camera_input("Capture Invoice Image")
        else:
            file_data = st.file_uploader(
                "Choose Invoice Scan (Image/PDF)", 
                type=["png", "jpg", "jpeg", "pdf"]
            )

        col_m1, col_m2 = st.columns(2)
        with col_m1:
            if st.button("Finalize & Save Entry", type="primary", use_container_width=True):
                if file_data is None:
                    st.warning("Please upload or capture a scan copy before submitting.")
                else:
                    with st.spinner("Uploading Scan to Drive & Saving History Log..."):
                        # Prepare File Payload
                        bytes_content = file_data.getvalue()
                        clean_inv = data_entry["inv_number"].replace("/", "_")
                        ext = file_data.name.split(".")[-1] if hasattr(file_data, "name") and "." in file_data.name else "png"
                        file_name = f"Purchase_{clean_inv}.{ext}"
                        mime_type = getattr(file_data, "type", "image/png") or "image/png"

                        # 1. Upload scan file to Google Drive via Apps Script
                        drive_link = upload_scan_to_drive_via_appscript(
                            bytes_content,
                            file_name,
                            mime_type,
                            data_entry["company_name"]
                        )

                        if drive_link:
                            # 2. Capture automatic timestamp
                            entry_timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

                            # 3. Append complete entry to Google Sheet
                            # Layout: [Invoice Date, Entry Timestamp, Company Name, Invoice No, Product Name, Quantity, Rate, Unit, Taxable Value, GST %, Total GST, Round Off, Invoice Value, Scan Link, Remark]
                            purchase_sheet.append_row([
                                data_entry["inv_date"],
                                entry_timestamp,
                                data_entry["company_name"],
                                data_entry["inv_number"],
                                data_entry["product_name"],
                                data_entry["quantity"],
                                data_entry["rate"],
                                data_entry["unit"],
                                data_entry["taxable_value"],
                                f"{data_entry['gst_pct']}%",
                                data_entry["total_gst"],
                                data_entry["round_off"],
                                data_entry["inv_value"],
                                drive_link,
                                data_entry["remark"]
                            ])

                            st.success("✅ Purchase entry & scan logged successfully!")
                            st.session_state.upload_modal_data = None
                            st.session_state.form_key += 1
                            st.rerun()

        with col_m2:
            if st.button("Cancel Entry", use_container_width=True):
                st.session_state.upload_modal_data = None
                st.rerun()

    open_scan_upload_modal()

# ------------------------------------------------------
# TAB 2: HISTORY LOG VIEW
# ------------------------------------------------------
with tab2:
    st.header("Purchase History Log")
    if st.button("🔄 Refresh History Log"):
        st.cache_data.clear()
        st.rerun()

    raw_data = purchase_sheet.get_all_values()
    if len(raw_data) > 1:
        headers = raw_data[0]
        df_log = pd.DataFrame(raw_data[1:], columns=headers)
        
        st.dataframe(
            df_log,
            column_config={
                "Scan Copy Link": st.column_config.LinkColumn(
                    "Scan Copy Link",
                    display_text="📄 View Document"
                )
            },
            use_container_width=True
        )
    else:
        st.info("No purchase entries recorded yet.")
