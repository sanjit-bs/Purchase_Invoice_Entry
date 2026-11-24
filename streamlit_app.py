import base64
from datetime import datetime, date
import json
import gspread
from google.oauth2.service_account import Credentials
from google import genai
from google.genai import types
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
    "KALIKA PRESS PVT. LTD.",
    "P.S.INDUSTRIES",
    "New Subir Paper Mart",
    "Reflection Papers Pvt. Ltd.",
    "RIPCO TRADERS PVT. LTD.",
    "Super  Adhesives",
    "Calcutta Carton Company",
    "KRISHNA TRADERS",
    "Sanjay Sales Corporation",
    "Others"
]

UNIT_OPTIONS = ["per kg", "per pcs", "per gross", "per ream", "per box", "per pkt", "per roll", "Others"]
GST_PERCENTAGE_OPTIONS = [0.0, 5.0, 12.0, 18.0, 28.0]

# ======================================================
# Database Connector & History Cache
# ======================================================
@st.cache_resource(ttl=3600)
def connect_purchase_log_sheet():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID).worksheet(WORKSHEET_NAME)

purchase_sheet = connect_purchase_log_sheet()

@st.cache_data(ttl=15)
def load_history_df():
    raw_data = purchase_sheet.get_all_values()
    if len(raw_data) > 1:
        return pd.DataFrame(raw_data[1:], columns=raw_data[0])
    return pd.DataFrame()

# ======================================================
# AI Invoice OCR Extractor (Gemini Vision)
# ======================================================
def extract_invoice_data_from_scan(file_bytes, mime_type):
    """
    Reads an invoice image/PDF using Gemini Vision and returns structured fields.
    Returns (dict, None) on success, or (None, error_message) on failure.
    """
    api_key = st.secrets.get("GEMINI_API_KEY", "")
    if not api_key:
        return None, "GEMINI_API_KEY not found in secrets."

    prompt = """
    Extract the following details from this Indian purchase tax invoice.
    Return ONLY a valid JSON object with these exact keys:
    - "invoice_date": string in "DD/MM/YYYY" format (or empty string if unreadable)
    - "company_name": seller/supplier company name as string (or empty string)
    - "invoice_number": invoice number as string (or empty string)
    - "product_name": primary item/product description as string (or empty string)
    - "hsn": HSN / SAC code as string (e.g. "4802", or empty string if not visible)
    - "unit": one of ["per kg", "per pcs", "per gross", "per ream", "per box", "per pkt", "per roll", "Others"] based on unit (default "per kg")
    - "custom_unit": if unit is "Others", specify the unit string (otherwise empty string)
    - "quantity": numeric float (0.0 if unreadable)
    - "rate": numeric float unit price before tax (0.0 if unreadable)
    - "gst_percentage": total GST rate as float (e.g., 5.0, 12.0, 18.0, 28.0, or 0.0)
    - "round_off": numeric float round off adjustment (0.0 if none)
    """

    try:
        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[
                types.Part.from_bytes(data=file_bytes, mime_type=mime_type),
                prompt
            ],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                temperature=0.1
            )
        )
        extracted = json.loads(response.text)
        if not extracted.get("invoice_number") and not extracted.get("company_name") and float(extracted.get("quantity", 0)) <= 0:
            return None, "Scan is unclear or could not detect invoice fields."
        return extracted, None
    except Exception as e:
        return None, f"Auto-read failed ({e})."


def match_company_dropdown(extracted_name):
    """Matches extracted supplier name against COMPANY_LIST or falls back to 'Others'."""
    if not extracted_name:
        return "Select Company", ""
    clean_ext = extracted_name.strip().lower()
    for comp in COMPANY_LIST[1:-1]:
        if comp.lower() in clean_ext or clean_ext in comp.lower():
            return comp, comp
    return "Others", extracted_name.strip()

# ======================================================
# HSN Verification Helper
# ======================================================
def verify_company_hsn(df_history, company_name, product_name, entered_hsn):
    """
    Checks if the company has previous entries in Purchase_Invoice_Log and compares HSN codes.
    Returns (is_mismatch: bool, warning_message: str)
    """
    if df_history.empty or "Company Name" not in df_history.columns or "HSN" not in df_history.columns:
        return False, ""

    comp_clean = company_name.strip().lower()
    prod_clean = product_name.strip().lower()
    hsn_clean = entered_hsn.strip().lower()

    comp_rows = df_history[df_history["Company Name"].astype(str).str.strip().str.lower() == comp_clean]
    if comp_rows.empty:
        return False, ""

    valid_hsn_rows = comp_rows[comp_rows["HSN"].astype(str).str.strip() != ""]
    if valid_hsn_rows.empty:
        return False, ""

    # 1. Same Product Name, Different HSN
    if "Product Name" in valid_hsn_rows.columns:
        same_prod_rows = valid_hsn_rows[valid_hsn_rows["Product Name"].astype(str).str.strip().str.lower() == prod_clean]
        if not same_prod_rows.empty:
            prev_prod_hsns = sorted(same_prod_rows["HSN"].astype(str).str.strip().unique().tolist())
            if hsn_clean not in [h.lower() for h in prev_prod_hsns]:
                return True, (
                    f"⚠️ **HSN Mismatch for Existing Product!**\n\n"
                    f"Company **{company_name}** already has **{product_name}** recorded under HSN: **`{', '.join(prev_prod_hsns)}`**, "
                    f"but you entered **`{entered_hsn}`**."
                )

    # 2. Different HSN from all previously recorded HSNs for this company
    prev_company_hsns = sorted(valid_hsn_rows["HSN"].astype(str).str.strip().unique().tolist())
    if hsn_clean not in [h.lower() for h in prev_company_hsns]:
        prev_prods = []
        if "Product Name" in valid_hsn_rows.columns:
            prev_prods = sorted(valid_hsn_rows["Product Name"].astype(str).str.strip().unique().tolist())
        prod_summary = f" (Products: {', '.join(prev_prods)})" if prev_prods else ""
        return True, (
            f"⚠️ **New HSN Detected for {company_name}!**\n\n"
            f"Previous HSN(s) recorded for this company: **`{', '.join(prev_company_hsns)}`**{prod_summary}.\n\n"
            f"You entered HSN: **`{entered_hsn}`** for **{product_name}**."
        )

    return False, ""

# ======================================================
# Apps Script Google Drive Upload & Save Helper
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

def execute_final_save(entry_data, scan_obj):
    with st.spinner("Uploading Scan to Google Drive & Saving Entry..."):
        clean_inv = entry_data["inv_number"].replace("/", "_")
        ext = scan_obj["name"].split(".")[-1] if "." in scan_obj["name"] else "png"
        file_name = f"Purchase_{clean_inv}.{ext}"

        drive_link = upload_scan_to_drive_via_appscript(
            scan_obj["bytes"],
            file_name,
            scan_obj["mime"],
            entry_data["company_name"]
        )

        if drive_link:
            entry_timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            purchase_sheet.append_row([
                entry_data["inv_date"],
                entry_timestamp,
                entry_data["company_name"],
                entry_data["inv_number"],
                entry_data["product_name"],
                entry_data["hsn"],
                entry_data["quantity"],
                entry_data["rate"],
                entry_data["unit"],
                entry_data["taxable_value"],
                f"{entry_data['gst_pct']}%",
                entry_data["total_gst"],
                entry_data["round_off"],
                entry_data["inv_value"],
                drive_link,
                entry_data["remark"]
            ])

            st.toast("✅ Purchase entry & scan logged successfully!")
            st.cache_data.clear()
            st.session_state.scanned_file_payload = None
            st.session_state.last_scanned_file_id = None
            st.session_state.hsn_hold_state = None
            st.session_state.form_key += 1
            st.rerun()

# ======================================================
# Session State Setup
# ======================================================
if "form_key" not in st.session_state:
    st.session_state.form_key = 0
if "last_scanned_file_id" not in st.session_state:
    st.session_state.last_scanned_file_id = None
if "scanned_file_payload" not in st.session_state:
    st.session_state.scanned_file_payload = None
if "hsn_hold_state" not in st.session_state:
    st.session_state.hsn_hold_state = None

key_suffix = st.session_state.form_key

# ======================================================
# Primary Navigation Tabs
# ======================================================
tab1, tab2 = st.tabs(["➕ New Purchase Entry", "📋 Purchase History Log"])

# ------------------------------------------------------
# TAB 1: AI SCAN AUTO-FILL + HSN VERIFICATION FORM
# ------------------------------------------------------
with tab1:
    st.subheader("1️⃣ Upload or Capture Invoice Scan (AI Auto-Fill)")

    with st.container(border=True):
        scan_source = st.radio(
            "Select Scan Source",
            ["📁 Upload File from Device", "📷 Camera Capture"],
            horizontal=True,
            key=f"scan_src_{key_suffix}"
        )

        if scan_source == "📷 Camera Capture":
            uploaded_scan = st.camera_input("Capture Invoice Image", key=f"cam_{key_suffix}")
        else:
            uploaded_scan = st.file_uploader(
                "Upload Invoice Scan (PNG, JPG, PDF) — Computer will auto-fill fields below",
                type=["png", "jpg", "jpeg", "pdf"],
                key=f"upl_{key_suffix}"
            )

        if uploaded_scan is not None:
            file_bytes = uploaded_scan.getvalue()
            file_id = f"{getattr(uploaded_scan, 'name', 'cam')}_{len(file_bytes)}"
            mime_type = getattr(uploaded_scan, "type", "image/png") or "image/png"

            st.session_state.scanned_file_payload = {
                "bytes": file_bytes,
                "name": getattr(uploaded_scan, "name", "camera_capture.png"),
                "mime": mime_type
            }

            if st.session_state.last_scanned_file_id != file_id:
                with st.spinner("🤖 Reading invoice scan and auto-filling form..."):
                    extracted, err_msg = extract_invoice_data_from_scan(file_bytes, mime_type)
                    st.session_state.last_scanned_file_id = file_id

                    if extracted:
                        # 1. Date
                        try:
                            parsed_dt = datetime.strptime(extracted.get("invoice_date", ""), "%d/%m/%Y").date()
                            st.session_state[f"inv_date_{key_suffix}"] = parsed_dt
                        except ValueError:
                            pass

                        # 2. Company
                        dd_comp, manual_comp = match_company_dropdown(extracted.get("company_name", ""))
                        st.session_state[f"comp_{key_suffix}"] = dd_comp
                        if dd_comp == "Others":
                            st.session_state[f"manual_comp_{key_suffix}"] = manual_comp

                        # 3. Text, HSN & Unit
                        st.session_state[f"inv_num_{key_suffix}"] = str(extracted.get("invoice_number", ""))
                        st.session_state[f"prod_name_{key_suffix}"] = str(extracted.get("product_name", ""))
                        st.session_state[f"hsn_{key_suffix}"] = str(extracted.get("hsn", ""))

                        ext_unit = str(extracted.get("unit", "per kg"))
                        if ext_unit in UNIT_OPTIONS:
                            st.session_state[f"unit_{key_suffix}"] = ext_unit
                            if ext_unit == "Others":
                                st.session_state[f"manual_unit_{key_suffix}"] = str(extracted.get("custom_unit", ""))

                        # 4. Numbers & GST
                        st.session_state[f"qty_{key_suffix}"] = float(extracted.get("quantity", 0.0) or 0.0)
                        st.session_state[f"rate_{key_suffix}"] = float(extracted.get("rate", 0.0) or 0.0)
                        st.session_state[f"round_{key_suffix}"] = float(extracted.get("round_off", 0.0) or 0.0)

                        ext_gst = float(extracted.get("gst_percentage", 5.0) or 5.0)
                        if ext_gst in GST_PERCENTAGE_OPTIONS:
                            st.session_state[f"gst_pct_{key_suffix}"] = ext_gst

                        st.success("✅ Scan read successfully! Please verify the auto-filled values below.")
                    else:
                        st.warning(f"⚠️ Could not read scan automatically ({err_msg}) **Please enter the details manually below.**")

    st.subheader("2️⃣ Verify or Enter Purchased Invoice Details")

    col1, col2 = st.columns(2)
    with col1:
        inv_date = st.date_input("Invoice Date *", value=date.today(), format="DD/MM/YYYY", key=f"inv_date_{key_suffix}")
        
        selected_company = st.selectbox("Company Name *", COMPANY_LIST, key=f"comp_{key_suffix}")
        if selected_company == "Others":
            company_name = st.text_input("Enter Manual Company Name *", key=f"manual_comp_{key_suffix}")
        else:
            company_name = selected_company

        inv_number = st.text_input("Invoice Number *", placeholder="e.g. INV/2026/001", key=f"inv_num_{key_suffix}")
        product_name = st.text_input("Product Name *", placeholder="e.g. Maplitho Paper", key=f"prod_name_{key_suffix}")
        hsn_code = st.text_input("HSN Code *", placeholder="e.g. 4802", key=f"hsn_{key_suffix}")

        selected_unit = st.selectbox("Unit *", UNIT_OPTIONS, key=f"unit_{key_suffix}")
        if selected_unit == "Others":
            final_unit = st.text_input("Enter Custom Unit *", placeholder="e.g. per bundle / per set", key=f"manual_unit_{key_suffix}")
        else:
            final_unit = selected_unit

    with col2:
        quantity = st.number_input("Quantity *", min_value=0.0, step=1.0, format="%.2f", key=f"qty_{key_suffix}")
        rate = st.number_input("Rate *", min_value=0.0, step=0.5, format="%.2f", key=f"rate_{key_suffix}")
        
        # 1. Automatic Taxable Value calculation (Quantity * Rate)
        taxable_value = round(quantity * rate, 2)
        st.number_input(
            "Taxable Value (₹) [Qty × Rate]",
            value=float(taxable_value),
            format="%.2f",
            disabled=True,
            key=f"tax_val_{key_suffix}_{taxable_value}"
        )

        # 2. GST Percentage Selection & Live Calculation
        gst_pct = st.selectbox("GST Percentage (%) *", GST_PERCENTAGE_OPTIONS, index=1, key=f"gst_pct_{key_suffix}")
        
        half_gst_pct = gst_pct / 2.0
        sgst_val = round(taxable_value * (half_gst_pct / 100.0), 2)
        cgst_val = round(taxable_value * (half_gst_pct / 100.0), 2)
        total_gst = round(sgst_val + cgst_val, 2)

        c_m1, c_m2, c_m3 = st.columns(3)
        c_m1.metric(f"SGST ({half_gst_pct}%)", f"₹ {sgst_val:,.2f}")
        c_m2.metric(f"CGST ({half_gst_pct}%)", f"₹ {cgst_val:,.2f}")
        c_m3.metric(f"Total GST ({gst_pct}%)", f"₹ {total_gst:,.2f}")

        # 3. Round Off Adjustment
        round_off = st.number_input("Round Off (+/-)", step=0.01, format="%.2f", key=f"round_{key_suffix}")
        
        # 4. Final Invoice Value Calculation
        calc_inv_val = round(taxable_value + total_gst + round_off, 2)
        inv_val_key = f"inv_val_{key_suffix}"
        last_calc_key = f"last_calc_{key_suffix}"

        if st.session_state.get(last_calc_key) != calc_inv_val:
            st.session_state[inv_val_key] = float(calc_inv_val)
            st.session_state[last_calc_key] = calc_inv_val

        inv_value = st.number_input(
            "Final Invoice Value (₹) *",
            min_value=0.0,
            step=1.0,
            format="%.2f",
            key=inv_val_key
        )

    remark = st.text_area("Remark / Notes", placeholder="Supplier details, batch numbers, or terms...", key=f"rem_{key_suffix}")

    # Submission + HSN Verification Check
    if st.button("✅ Save Purchase Entry & Upload Scan", type="primary", use_container_width=True):
        if st.session_state.scanned_file_payload is None:
            st.warning("⚠️ Please upload or capture the invoice scan copy at the top first.")
        elif selected_company == "Select Company" or not company_name.strip():
            st.warning("⚠️ Please select or enter a valid Company Name.")
        elif not inv_number.strip():
            st.warning("⚠️ Please enter an Invoice Number.")
        elif not product_name.strip():
            st.warning("⚠️ Please enter a Product Name.")
        elif not hsn_code.strip():
            st.warning("⚠️ Please enter an HSN Code.")
        elif selected_unit == "Others" and not final_unit.strip():
            st.warning("⚠️ Please enter the custom Unit name.")
        elif quantity <= 0:
            st.warning("⚠️ Quantity must be greater than zero.")
        elif rate <= 0:
            st.warning("⚠️ Rate must be greater than zero.")
        else:
            entry_payload = {
                "inv_date": inv_date.strftime("%d/%m/%Y"),
                "company_name": company_name.strip(),
                "inv_number": inv_number.strip(),
                "product_name": product_name.strip(),
                "hsn": hsn_code.strip(),
                "quantity": quantity,
                "rate": rate,
                "unit": final_unit.strip(),
                "taxable_value": taxable_value,
                "gst_pct": gst_pct,
                "total_gst": total_gst,
                "round_off": round_off,
                "inv_value": inv_value,
                "remark": remark.strip()
            }

            # Verify HSN against historical entries of the same company
            df_hist = load_history_df()
            is_mismatch, warn_msg = verify_company_hsn(
                df_hist,
                entry_payload["company_name"],
                entry_payload["product_name"],
                entry_payload["hsn"]
            )

            if is_mismatch:
                st.session_state.hsn_hold_state = {
                    "warning": warn_msg,
                    "payload": entry_payload
                }
                st.rerun()
            else:
                execute_final_save(entry_payload, st.session_state.scanned_file_payload)

    # --------------------------------------------------
    # HSN MISMATCH CONFIRMATION HOLD BANNER
    # --------------------------------------------------
    if st.session_state.hsn_hold_state is not None:
        hold_info = st.session_state.hsn_hold_state
        with st.container(border=True):
            st.warning(hold_info["warning"])
            st.markdown("**Submission is on hold.** Please confirm if this is a different/new product for this company:")
            c_conf1, c_conf2 = st.columns(2)
            with c_conf1:
                if st.button("✅ Confirm: This is another product (Save Entry)", type="primary", use_container_width=True):
                    if st.session_state.scanned_file_payload is not None:
                        execute_final_save(hold_info["payload"], st.session_state.scanned_file_payload)
                    else:
                        st.error("Scan file missing from session. Please re-attach the scan above.")
            with c_conf2:
                if st.button("❌ Cancel & Edit HSN / Product", use_container_width=True):
                    st.session_state.hsn_hold_state = None
                    st.rerun()

# ------------------------------------------------------
# TAB 2: HISTORY LOG VIEW
# ------------------------------------------------------
with tab2:
    st.header("Purchase History Log")
    if st.button("🔄 Refresh History Log"):
        st.cache_data.clear()
        st.rerun()

    df_log = load_history_df()
    if not df_log.empty:
        st.dataframe(
            df_log.iloc[::-1],
            column_config={
                "Scan Copy Link": st.column_config.LinkColumn(
                    "Scan Copy Link",
                    display_text="📄 View Document"
                )
            },
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("No purchase entries recorded yet.")
