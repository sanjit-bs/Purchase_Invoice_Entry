import base64
from datetime import datetime, date
import io
import re
import fitz  # PyMuPDF
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from PIL import Image, ImageOps
import pytesseract
import requests
import streamlit as st

# ======================================================
# Page & Layout Configuration
# ======================================================
st.set_page_config(page_title="Purchased Invoice Entry", page_icon="🧾", layout="wide")

st.title("🧾 Purchased Invoice Entry System")

# Google Apps Script Web App Endpoint, Sheet ID & Root Shared Drive Folder ID
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyOhnkbEa7JHQk49LAH1xg4PXMzeMVwPDfesMbOxeOvsDjcOjL7dEXYBGRymn-Sx6MKkQ/exec"
SPREADSHEET_ID = "1PCJ3BWAj6Wz1N-55XpuWltvfvYd7KD1q194D3N7MzIg"
WORKSHEET_NAME = "Purchase_Invoice_Log"

# Paste the Root Folder ID from your father's Google Drive (shared with your Service Account email)
SHARED_DRIVE_FOLDER_ID = "https://drive.google.com/drive/folders/15JHQ9jS0EmlmqKnY943eVdUDwD26b1X5"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

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

PRODUCT_CATEGORY_OPTIONS = [
    "Select Category",
    "Paper",
    "Board / Duplex",
    "Adhesive / Gum",
    "Ink / Chemical",
    "Packaging / Carton",
    "Printing Plate / Consumables",
    "Others"
]

UNIT_OPTIONS = ["per kg", "per pcs", "per gross", "per ream", "per box", "per pkt", "per roll", "Others"]
GST_PERCENTAGE_OPTIONS = [0.0, 5.0, 12.0, 18.0, 28.0]

# ======================================================
# Google Sheets & Shared Google Drive Connectors
# ======================================================
@st.cache_resource(ttl=3600)
def get_google_clients():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    gs_client = gspread.authorize(creds)
    sheet = gs_client.open_by_key(SPREADSHEET_ID).worksheet(WORKSHEET_NAME)
    drive_service = build("drive", "v3", credentials=creds)
    return sheet, drive_service

purchase_sheet, drive_service = get_google_clients()

@st.cache_data(ttl=15)
def load_history_df():
    raw_data = purchase_sheet.get_all_values()
    if len(raw_data) > 1:
        return pd.DataFrame(raw_data[1:], columns=raw_data[0])
    return pd.DataFrame()

@st.cache_data(ttl=30)
def list_drive_folder_contents(folder_id):
    """Lists both subfolders and PDF/Image files inside a given Google Drive folder."""
    if not folder_id or folder_id == "PASTE_YOUR_SHARED_FOLDER_ID_HERE":
        return [], []
    try:
        # 1. Fetch Subfolders
        folder_query = (
            f"'{folder_id}' in parents and trashed = false and "
            "mimeType = 'application/vnd.google-apps.folder'"
        )
        folder_res = drive_service.files().list(
            q=folder_query,
            fields="files(id, name)",
            orderBy="name asc",
            pageSize=100,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        subfolders = folder_res.get("files", [])

        # 2. Fetch PDF & Image Files
        file_query = (
            f"'{folder_id}' in parents and trashed = false and "
            "(mimeType='application/pdf' or mimeType contains 'image/')"
        )
        file_res = drive_service.files().list(
            q=file_query,
            fields="files(id, name, mimeType, webViewLink, createdTime)",
            orderBy="createdTime desc",
            pageSize=100,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        files = file_res.get("files", [])

        return subfolders, files
    except Exception as e:
        st.error(f"Could not read Google Drive folder: {e}")
        return [], []

def download_drive_file_bytes(file_id):
    request = drive_service.files().get_media(fileId=file_id, supportsAllDrives=True)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buffer.getvalue()

# ======================================================
# Free Local OCR Extractor (Tesseract + PyMuPDF)
# ======================================================
def preprocess_and_ocr_image(pil_img):
    gray = ImageOps.grayscale(pil_img)
    gray = ImageOps.autocontrast(gray)
    return pytesseract.image_to_string(gray, config="--oem 3 --psm 6")

def extract_text_from_scan(file_bytes, mime_type):
    if "pdf" in mime_type.lower():
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        full_text = ""
        for page in doc:
            page_text = page.get_text("text").strip()
            if len(page_text) > 30:
                full_text += page_text + "\n"
            else:
                pix = page.get_pixmap(dpi=250)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                full_text += preprocess_and_ocr_image(img) + "\n"
        return full_text.strip()
    else:
        img = Image.open(io.BytesIO(file_bytes))
        return preprocess_and_ocr_image(img).strip()

def parse_indian_invoice_text(raw_text):
    if not raw_text or len(raw_text.strip()) < 10:
        return None, "Could not detect readable text on the scan."

    extracted = {
        "invoice_date": "",
        "company_name": "",
        "invoice_number": "",
        "product_name": "",
        "hsn": "",
        "unit": "per kg",
        "quantity": 0.0,
        "rate": 0.0,
        "gst_percentage": 5.0,
        "round_off": 0.0,
    }

    text_upper = raw_text.upper()

    for comp in COMPANY_LIST[1:-1]:
        clean_tokens = [t for t in re.split(r"[^A-Z0-9]+", comp.upper()) if len(t) > 2 and t not in ["PVT", "LTD", "COMPANY", "CORPORATION"]]
        if clean_tokens and all(tok in text_upper for tok in clean_tokens[:2]):
            extracted["company_name"] = comp
            break

    date_match = re.search(
        r"(?:DATE|DATED|INV\.?\s*DATE)[\s:\-]*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})",
        raw_text,
        re.IGNORECASE,
    )
    if not date_match:
        date_match = re.search(r"\b(\d{1,2}[\/\-]\d{1,2}[\/\-]20\d{2})\b", raw_text)
    if date_match:
        raw_dt = date_match.group(1).replace("-", "/").replace(".", "/")
        parts = raw_dt.split("/")
        if len(parts) == 3:
            if len(parts[2]) == 2:
                parts[2] = "20" + parts[2]
            extracted["invoice_date"] = f"{parts[0].zfill(2)}/{parts[1].zfill(2)}/{parts[2]}"

    inv_match = re.search(
        r"(?:INVOICE\s*NO|INV\.?\s*NO|BILL\s*NO|CHALLAN\s*NO)[\s.:\-#]*([A-Z0-9\/\-]{2,25})",
        raw_text,
        re.IGNORECASE,
    )
    if inv_match:
        candidate = inv_match.group(1).strip()
        if candidate.upper() not in ["DATE", "DATED", "TO", "GSTIN"]:
            extracted["invoice_number"] = candidate

    hsn_match = re.search(r"(?:HSN(?:\/SAC)?|SAC)[\s.:\-]*(\d{4,8})\b", raw_text, re.IGNORECASE)
    if hsn_match:
        extracted["hsn"] = hsn_match.group(1)

    split_gst = re.search(r"(?:CGST|SGST)[\s@:\-]*(\d{1,2}(?:\.\d+)?)\s*%", raw_text, re.IGNORECASE)
    full_gst = re.search(r"(?:IGST|GST)[\s@:\-]*(\d{1,2}(?:\.\d+)?)\s*%", raw_text, re.IGNORECASE)
    if split_gst:
        doubled = round(float(split_gst.group(1)) * 2, 1)
        if doubled in GST_PERCENTAGE_OPTIONS:
            extracted["gst_percentage"] = doubled
    elif full_gst:
        val = float(full_gst.group(1))
        if val in GST_PERCENTAGE_OPTIONS:
            extracted["gst_percentage"] = val

    unit_map = {
        "per kg": r"\b(?:KGS?|KILOGRAMS?)\b",
        "per pcs": r"\b(?:PCS|PIECES?|NOS)\b",
        "per gross": r"\b(?:GROSS|GRUS)\b",
        "per ream": r"\b(?:REAMS?)\b",
        "per box": r"\b(?:BOX(?:ES)?)\b",
        "per pkt": r"\b(?:PKTS?|PACKETS?)\b",
        "per roll": r"\b(?:ROLLS?)\b",
    }
    for u_label, pattern in unit_map.items():
        if re.search(pattern, text_upper):
            extracted["unit"] = u_label
            break

    qty_rate_match = re.search(
        r"(\d+(?:,\d{3})*(?:\.\d+)?)\s*(?:KGS?|PCS|NOS|REAMS?|GROSS|BOX|PKT|ROLL)?\s*[@xX×]?\s*(\d+(?:,\d{3})*(?:\.\d{1,2}))",
        raw_text,
        re.IGNORECASE,
    )
    if qty_rate_match:
        try:
            q_val = float(qty_rate_match.group(1).replace(",", ""))
            r_val = float(qty_rate_match.group(2).replace(",", ""))
            if 0 < q_val < 1000000 and 0 < r_val < 1000000:
                extracted["quantity"] = q_val
                extracted["rate"] = r_val
        except ValueError:
            pass

    ro_match = re.search(r"ROUND\s*OFF[\s:\-]*([+-]?\s*\d+(?:\.\d{1,2})?)", raw_text, re.IGNORECASE)
    if ro_match:
        try:
            extracted["round_off"] = float(ro_match.group(1).replace(" ", ""))
        except ValueError:
            pass

    found_any = any([
        extracted["company_name"],
        extracted["invoice_number"],
        extracted["hsn"],
        extracted["quantity"] > 0,
    ])
    if not found_any:
        return None, "Could not match invoice fields from scan."

    return extracted, None

def run_local_invoice_ocr(file_bytes, mime_type):
    try:
        raw_text = extract_text_from_scan(file_bytes, mime_type)
        parsed, err = parse_indian_invoice_text(raw_text)
        return parsed, raw_text, err
    except pytesseract.TesseractNotFoundError:
        return None, "", "Tesseract OCR engine not installed (add 'tesseract-ocr' to packages.txt and reboot)."
    except Exception as e:
        return None, "", f"OCR error: {e}"

def match_company_dropdown(extracted_name):
    if not extracted_name:
        return "Select Company", ""
    if extracted_name in COMPANY_LIST:
        return extracted_name, extracted_name
    return "Others", extracted_name.strip()

# ======================================================
# HSN Verification Helper
# ======================================================
def verify_company_hsn(df_history, company_name, product_name, entered_hsn):
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
                entry_data["product_category"],
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
            st.session_state.raw_ocr_text = ""
            st.session_state.hsn_hold_state = None
            st.session_state.form_key += 1
            st.rerun()

# ======================================================
# Session State Setup (Including Folder Navigation Stack)
# ======================================================
if "form_key" not in st.session_state:
    st.session_state.form_key = 0
if "last_scanned_file_id" not in st.session_state:
    st.session_state.last_scanned_file_id = None
if "scanned_file_payload" not in st.session_state:
    st.session_state.scanned_file_payload = None
if "hsn_hold_state" not in st.session_state:
    st.session_state.hsn_hold_state = None
if "raw_ocr_text" not in st.session_state:
    st.session_state.raw_ocr_text = ""
if "drive_folder_stack" not in st.session_state:
    st.session_state.drive_folder_stack = [{"id": SHARED_DRIVE_FOLDER_ID, "name": "Main Shared Folder"}]

key_suffix = st.session_state.form_key

# ======================================================
# Primary Navigation Tabs
# ======================================================
tab1, tab2 = st.tabs(["➕ New Purchase Entry", "📋 Purchase History Log"])

# ------------------------------------------------------
# TAB 1: SCAN / CAPTURE / SUBFOLDER GDRIVE + ENTRY FORM
# ------------------------------------------------------
with tab1:
    st.subheader("1️⃣ Attach Invoice Scan (Device Upload, Camera, or Shared Drive)")

    with st.container(border=True):
        scan_source = st.radio(
            "Select Scan Source",
            ["📁 Upload File from Device", "📷 Camera Capture", "☁️ Pick from Google Drive Folder"],
            horizontal=True,
            key=f"scan_src_{key_suffix}"
        )

        file_bytes = None
        file_id = None
        file_name_src = "scan.png"
        mime_type = "image/png"

        if scan_source == "📷 Camera Capture":
            uploaded_scan = st.camera_input("Capture Invoice Image", key=f"cam_{key_suffix}")
            if uploaded_scan is not None:
                file_bytes = uploaded_scan.getvalue()
                file_name_src = "camera_capture.png"
                mime_type = getattr(uploaded_scan, "type", "image/png") or "image/png"
                file_id = f"cam_{len(file_bytes)}"

        elif scan_source == "📁 Upload File from Device":
            uploaded_scan = st.file_uploader(
                "Upload Invoice Scan (PNG, JPG, PDF) — Computer will read text & fill fields below",
                type=["png", "jpg", "jpeg", "pdf"],
                key=f"upl_{key_suffix}"
            )
            if uploaded_scan is not None:
                file_bytes = uploaded_scan.getvalue()
                file_name_src = uploaded_scan.name
                mime_type = getattr(uploaded_scan, "type", "image/png") or "image/png"
                file_id = f"upl_{file_name_src}_{len(file_bytes)}"

        else:
            if SHARED_DRIVE_FOLDER_ID == "PASTE_YOUR_SHARED_FOLDER_ID_HERE":
                st.info("Set `SHARED_DRIVE_FOLDER_ID` in `app.py` and share that folder from your father's Google Drive with your Service Account email.")
            else:
                current_folder = st.session_state.drive_folder_stack[-1]
                breadcrumb = " / ".join([f["name"] for f in st.session_state.drive_folder_stack])

                nav_c1, nav_c2, nav_c3 = st.columns([3.5, 1, 1])
                with nav_c1:
                    st.markdown(f"**📂 Current Folder:** `{breadcrumb}`")
                with nav_c2:
                    if len(st.session_state.drive_folder_stack) > 1:
                        if st.button("⬅️ Back One Folder", use_container_width=True):
                            st.session_state.drive_folder_stack.pop()
                            st.rerun()
                with nav_c3:
                    if st.button("🔄 Refresh", use_container_width=True):
                        list_drive_folder_contents.clear()
                        st.rerun()

                subfolders, drive_files = list_drive_folder_contents(current_folder["id"])

                # Subfolder Navigation Selector
                if subfolders:
                    folder_names = ["Select Subfolder to Open..."] + [sf["name"] for sf in subfolders]
                    chosen_sub = st.selectbox(
                        "📁 Open Subfolder:",
                        folder_names,
                        key=f"sub_sel_{current_folder['id']}_{key_suffix}"
                    )
                    if chosen_sub != "Select Subfolder to Open...":
                        target_sf = next(sf for sf in subfolders if sf["name"] == chosen_sub)
                        st.session_state.drive_folder_stack.append({"id": target_sf["id"], "name": target_sf["name"]})
                        st.rerun()

                # File Selector inside the active folder/subfolder
                if not drive_files:
                    st.caption("No PDF or Image files directly inside this folder. Open a subfolder above if files are inside.")
                else:
                    file_options = ["Select an Invoice File..."] + [f["name"] for f in drive_files]
                    chosen_name = st.selectbox(
                        "📄 Select Invoice File:",
                        file_options,
                        key=f"gd_file_sel_{current_folder['id']}_{key_suffix}"
                    )

                    if chosen_name != "Select an Invoice File...":
                        chosen_meta = next(f for f in drive_files if f["name"] == chosen_name)
                        file_id = f"gd_{chosen_meta['id']}"
                        file_name_src = chosen_meta["name"]
                        mime_type = chosen_meta["mimeType"]

                        if st.session_state.last_scanned_file_id != file_id:
                            with st.spinner(f"Downloading '{file_name_src}' from Google Drive..."):
                                file_bytes = download_drive_file_bytes(chosen_meta["id"])
                        elif st.session_state.scanned_file_payload is not None:
                            file_bytes = st.session_state.scanned_file_payload["bytes"]

        # Run Local OCR once per newly selected/uploaded file
        if file_bytes is not None and file_id is not None:
            st.session_state.scanned_file_payload = {
                "bytes": file_bytes,
                "name": file_name_src,
                "mime": mime_type
            }

            if st.session_state.last_scanned_file_id != file_id:
                with st.spinner("🔍 Scanning document locally with Tesseract OCR..."):
                    extracted, raw_text, err_msg = run_local_invoice_ocr(file_bytes, mime_type)
                    st.session_state.last_scanned_file_id = file_id
                    st.session_state.raw_ocr_text = raw_text

                    if extracted:
                        try:
                            parsed_dt = datetime.strptime(extracted.get("invoice_date", ""), "%d/%m/%Y").date()
                            st.session_state[f"inv_date_{key_suffix}"] = parsed_dt
                        except ValueError:
                            pass

                        dd_comp, manual_comp = match_company_dropdown(extracted.get("company_name", ""))
                        st.session_state[f"comp_{key_suffix}"] = dd_comp
                        if dd_comp == "Others":
                            st.session_state[f"manual_comp_{key_suffix}"] = manual_comp

                        if extracted.get("invoice_number"):
                            st.session_state[f"inv_num_{key_suffix}"] = str(extracted["invoice_number"])
                        if extracted.get("hsn"):
                            st.session_state[f"hsn_{key_suffix}"] = str(extracted["hsn"])

                        ext_unit = str(extracted.get("unit", "per kg"))
                        if ext_unit in UNIT_OPTIONS:
                            st.session_state[f"unit_{key_suffix}"] = ext_unit

                        if extracted.get("quantity", 0.0) > 0:
                            st.session_state[f"qty_{key_suffix}"] = float(extracted["quantity"])
                        if extracted.get("rate", 0.0) > 0:
                            st.session_state[f"rate_{key_suffix}"] = float(extracted["rate"])
                        if extracted.get("round_off", 0.0) != 0:
                            st.session_state[f"round_{key_suffix}"] = float(extracted["round_off"])

                        ext_gst = float(extracted.get("gst_percentage", 5.0))
                        if ext_gst in GST_PERCENTAGE_OPTIONS:
                            st.session_state[f"gst_pct_{key_suffix}"] = ext_gst

                        st.success("✅ Scan read locally! Please review and complete any remaining fields below.")
                    else:
                        st.warning(f"⚠️ Could not auto-fill from scan ({err_msg}) **Please enter the details manually below.**")

        if st.session_state.raw_ocr_text:
            with st.expander("📄 View Raw Scanned Text (For Quick Copy/Reference)"):
                st.text(st.session_state.raw_ocr_text)

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

        selected_category = st.selectbox("Product Category *", PRODUCT_CATEGORY_OPTIONS, key=f"prod_cat_{key_suffix}")
        if selected_category == "Others":
            final_category = st.text_input("Enter Custom Product Category *", placeholder="e.g. Lamination Film", key=f"manual_cat_{key_suffix}")
        else:
            final_category = selected_category

        hsn_code = st.text_input("HSN Code *", placeholder="e.g. 4802", key=f"hsn_{key_suffix}")

        selected_unit = st.selectbox("Unit *", UNIT_OPTIONS, key=f"unit_{key_suffix}")
        if selected_unit == "Others":
            final_unit = st.text_input("Enter Custom Unit *", placeholder="e.g. per bundle / per set", key=f"manual_unit_{key_suffix}")
        else:
            final_unit = selected_unit

    with col2:
        quantity = st.number_input("Quantity *", min_value=0.0, step=1.0, format="%.2f", key=f"qty_{key_suffix}")
        rate = st.number_input("Rate *", min_value=0.0, step=0.5, format="%.2f", key=f"rate_{key_suffix}")
        
        taxable_value = round(quantity * rate, 2)
        st.number_input(
            "Taxable Value (₹) [Qty × Rate]",
            value=float(taxable_value),
            format="%.2f",
            disabled=True,
            key=f"tax_val_{key_suffix}_{taxable_value}"
        )

        gst_pct = st.selectbox("GST Percentage (%) *", GST_PERCENTAGE_OPTIONS, index=1, key=f"gst_pct_{key_suffix}")
        
        half_gst_pct = gst_pct / 2.0
        sgst_val = round(taxable_value * (half_gst_pct / 100.0), 2)
        cgst_val = round(taxable_value * (half_gst_pct / 100.0), 2)
        total_gst = round(sgst_val + cgst_val, 2)

        c_m1, c_m2, c_m3 = st.columns(3)
        c_m1.metric(f"SGST ({half_gst_pct}%)", f"₹ {sgst_val:,.2f}")
        c_m2.metric(f"CGST ({half_gst_pct}%)", f"₹ {cgst_val:,.2f}")
        c_m3.metric(f"Total GST ({gst_pct}%)", f"₹ {total_gst:,.2f}")

        round_off = st.number_input("Round Off (+/-)", step=0.01, format="%.2f", key=f"round_{key_suffix}")
        
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

    if st.button("✅ Save Purchase Entry & Upload Scan", type="primary", use_container_width=True):
        if st.session_state.scanned_file_payload is None:
            st.warning("⚠️ Please attach an invoice scan (Upload, Camera, or Google Drive) at the top first.")
        elif selected_company == "Select Company" or not company_name.strip():
            st.warning("⚠️ Please select or enter a valid Company Name.")
        elif not inv_number.strip():
            st.warning("⚠️ Please enter an Invoice Number.")
        elif not product_name.strip():
            st.warning("⚠️ Please enter a Product Name.")
        elif selected_category == "Select Category" or not final_category.strip():
            st.warning("⚠️ Please select or enter a Product Category.")
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
                "product_category": final_category.strip(),
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
