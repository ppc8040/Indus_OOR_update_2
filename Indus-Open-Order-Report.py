import os
import requests
import pandas as pd
import logging
import time
import xml.etree.ElementTree as ET
from openpyxl.utils import get_column_letter as get_excel_column_letter
from datetime import datetime
import numpy as np
import traceback
import msal


class AuthConfig:
    CLIENT_ID = os.environ.get('AZURE_CLIENT_ID', '')
    CLIENT_SECRET = os.environ.get('AZURE_CLIENT_SECRET', '')
    authority = os.environ.get('AZURE_AUTHORITY', '')

def get_access_token():
    auth_config = AuthConfig()
    app = msal.ConfidentialClientApplication(
        auth_config.CLIENT_ID, 
        authority=auth_config.authority, 
        client_credential=auth_config.CLIENT_SECRET
    )
    
    result = app.acquire_token_silent(scopes=['https://graph.microsoft.com/.default'], account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=['https://graph.microsoft.com/.default'])
    
    if 'access_token' in result:
        return result['access_token']
    else:
        raise Exception("Failed to acquire access token.")

session = requests.Session()

def update_excel_data_via_graph_api(
    df,
    site_id,
    file_server_relative_url,
    worksheet_name="Open Order Report"
):
    """
    Updates Excel data using Microsoft Graph Excel API while users are actively editing.
    This bypasses file locks by working directly with Excel's co-authoring system.
    """
    access_token = get_access_token()
    
    # Create a persistent session for better performance and co-authoring support
    session_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/createSession"
    session_headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    session_payload = {
        "persistChanges": True
    }
    
    session_response = session.post(session_url, headers=session_headers, json=session_payload)
    if session_response.status_code not in [200, 201]:
        logger.error(f"Failed to create session: {session_response.status_code} {session_response.text}")
        return False
    
    session_id = session_response.json().get("id")
    logger.info(f"Created Excel session: {session_id}")
    
    # Headers for all subsequent requests
    api_headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
        "workbook-session-id": session_id
    }

    # Step 1: Check if worksheet exists
    ws_list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets"
    ws_list_resp = session.get(ws_list_url, headers=api_headers)
    ws_names = []
    if ws_list_resp.status_code == 200:
        ws_names = [ws['name'] for ws in ws_list_resp.json().get('value', [])]
    else:
        logger.warning("Could not retrieve worksheet list; proceeding and will handle errors if any.")

    # Step 2: Create worksheet if not found
    if worksheet_name not in ws_names:
        create_ws_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets/add"
        create_ws_payload = {"name": worksheet_name}
        create_ws_resp = session.post(create_ws_url, headers=api_headers, json=create_ws_payload)
        if create_ws_resp.status_code not in [200, 201]:
            logger.error(f"Failed to create worksheet '{worksheet_name}': {create_ws_resp.status_code} {create_ws_resp.text}")
            return False
        logger.info(f"Worksheet '{worksheet_name}' created.")

    
    try:
        # Clear existing data in the worksheet
        clear_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/usedRange/clear"
        clear_payload = {"applyTo": "Contents"}
        clear_response = session.post(clear_url, headers=api_headers, json=clear_payload)
        
        if clear_response.status_code not in [200, 204]:
            logger.warning(f"Clear operation response: {clear_response.status_code}")
        
        # Prepare data for Excel API
        # Convert DataFrame to list of lists (rows)
        headers = df.columns.tolist()
        data_conv = df.astype(str).where(~df.isna(), "")
        data_rows = [headers] + data_conv.values.tolist()
        
        # Calculate range for the data
        num_rows = len(data_rows)
        num_cols = len(headers)
        range_address = f"A1:{get_excel_column_letter(num_cols)}{num_rows}"
        
        # Update the range with new data
        update_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/range(address='{range_address}')"
        update_payload = {
            "values": data_rows
        }
        
        update_response = session.patch(update_url, headers=api_headers, json=update_payload)
        
        if update_response.status_code in [200, 201]:
            logger.info("Excel data updated successfully via Graph API")
            set_excel_update_timestamp(
                site_id,
                file_server_relative_url,
                worksheet_name,
                access_token,
                session_id
            )
            
            # Set right alignment for all cells in the data range
            align_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/range(address='{range_address}')/format"
            align_payload = {"horizontalAlignment": "Right"}
            resp_align = session.patch(align_url, headers=api_headers, json=align_payload)
            if resp_align.status_code in [200, 201]:
                logger.info("Right alignment applied to full data range in SharePoint Excel")
            else:
                logger.warning(f"Alignment update failed: {resp_align.status_code}, {resp_align.text}")

            # Apply all 6 borders one by one (must be done per-API limitations)
            border_ids = [
                "EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"
            ]
            for border_id in border_ids:
                border_url = (f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/"
                            f"workbook/worksheets('{worksheet_name}')/range(address='{range_address}')/format/borders('{border_id}')")
                border_payload = {
                    "style": "Continuous",
                    "color": "#000000",
                    "weight": "Thin"
                }
                resp_border = session.patch(border_url, headers=api_headers, json=border_payload)
                if resp_border.status_code in [200, 201]:
                    logger.info(f"Applied border {border_id}")
                else:
                    logger.warning(f"Border {border_id} update failed: {resp_border.status_code}, {resp_border.text}")

            
            # Assume headers in first row and columns from A to last_col
            header_range = f"A1:R1"

            # PATCH fill color
            header_fill_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/range(address='{header_range}')/format/fill"
            header_fill_payload = {"color": "#2F4F4F"}
            fill_response = session.patch(header_fill_url, headers=api_headers, json=header_fill_payload)
            if fill_response.status_code in [200,201]:
                logger.info("Header fill (background) applied.")
            else:
                logger.warning(f"Header fill failed: {fill_response.status_code}, {fill_response.text}")

            # PATCH font color and bold
            header_font_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/range(address='{header_range}')/format/font"
            header_font_payload = {"bold": True, "color": "#FFFFFF"}
            font_response = session.patch(header_font_url, headers=api_headers, json=header_font_payload)
            if font_response.status_code in [200,201]:
                logger.info("Header font color/bold applied.")
            else:
                logger.warning(f"Header font update failed: {font_response.status_code}, {font_response.text}")

            
            # Apply auto-fit columns
            autofit_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/range(address='{range_address}')/format/autofitColumns"
            autofit_resp = session.post(autofit_url, headers=api_headers)
            if autofit_resp.status_code in [200, 201, 204]:
                logger.info("Auto-fit columns applied successfully")
            else:
                logger.warning(f"Auto-fit columns failed: {autofit_resp.status_code}, {autofit_resp.text}")
            
            # Set auto-filter
            autofilter_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/autoFilter/apply"
            autofilter_payload = {
                "range": f"A1:{get_excel_column_letter(num_cols)}{num_rows}"
            }
            
            filter_response = session.post(autofilter_url, headers=api_headers, json=autofilter_payload)
            if filter_response.status_code in [200, 201]:
                logger.info("Auto-filter applied")
            
            return True
        else:
            logger.error(f"Failed to update Excel data: {update_response.status_code} {update_response.text}")
            return False
            
    except Exception as e:
        logger.error(f"Error updating Excel via Graph API: {str(e)}")
        return False
    finally:
        # Close the session
        close_session_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/closeSession"
        close_headers = {
            "Authorization": f"Bearer {access_token}",
            "workbook-session-id": session_id
        }
        close_response = session.post(close_session_url, headers=close_headers)
        if close_response.status_code in [200, 204]:
            logger.info("Excel session closed successfully")

def set_excel_update_timestamp(site_id, file_server_relative_url, worksheet_name, access_token, session_id):
    api_headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
        "workbook-session-id": session_id
    }
    # R1 = 18th column, 1st row (Excel columns are 1-indexed, but API expects A1 notation)
    update_datetime = datetime.now().strftime('Generated On: %d-%m-%Y %H:%M:%S')
    payload = {
        "values": [[update_datetime]]
    }
    update_url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/"
        f"workbook/worksheets('{worksheet_name}')/range(address='R1')"
    )
    r = session.patch(update_url, headers=api_headers, json=payload)
    if r.status_code in [200, 201]:
        logger.info("Timestamp for last update written to cell R1")
        return True
    else:
        logger.error(f"Failed to update timestamp in cell R1: {r.status_code} {r.text}")
        return False


def get_excel_column_letter(col_num):
    """Convert column number to Excel column letter (e.g., 1 -> A, 27 -> AA)"""
    letter = ""
    while col_num > 0:
        col_num -= 1
        letter = chr(col_num % 26 + ord('A')) + letter
        col_num //= 26
    return letter

def compute_indus_ship_date(row):
    # If already filled (notna and is number/date), use as-is
    if pd.notna(row['Indus Ship Date']) and str(row['Indus Ship Date']).strip() != "":
        return row['Indus Ship Date']
    # Use logic based on Customer Code
    base_date = row['Date of receiving PO#']
    code = str(row['Customer Code']).strip().upper()
    if code in ("IVT", "BTT"):
        if pd.notna(base_date):
            return base_date + pd.Timedelta(days=30)
    elif code == "OHP":
        if pd.notna(base_date):
            return base_date + pd.Timedelta(days=45)
    elif code in ("HTC", "ILA", "WFI", "STL", "SVA", "SIG"):
        if pd.notna(base_date):
            return base_date + pd.Timedelta(days=60)
    else:
        if pd.notna(base_date):
            return base_date + pd.Timedelta(days=90)
    return np.nan  # fallback if logic does not apply


# LOGGING SETUP
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("open_order_report.log", encoding="utf-8"), logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# SAP Configuration from environment variables
SAP_CONFIG = {
    "username": os.environ.get('SAP_USERNAME', ''),
    "password": os.environ.get('SAP_PASSWORD', ''),
    "sales_order_api_url": os.environ.get('SAP_SALES_ORDER_API_URL', ''),
    "stock_api_url": os.environ.get('SAP_STOCK_API_URL', ''),
    "fixed_excel_file": "Open_Order_Report_MASTER.xlsx"
}
SHAREPOINT_CONFIG = {
    "site_id": os.environ.get('SHAREPOINT_SITE_ID', ''),
    "file_path": os.environ.get('SHAREPOINT_FILE_PATH', ''),
    "worksheet_name": "Open Order Report"
}
INVENTORY_OUTPUT = [
    {"label": "Stock in 3001", "plant": "3000", "storloc": "3001"},
    {"label": "Stock in 3003", "plant": "3000", "storloc": "3003"},
    {"label": "Stock in 1000", "plant": "1000", "storloc": None},
    {"label": "Stock in 2000", "plant": "2000", "storloc": None},
]
OUTPUT_COLUMNS = [
    "Customer Code", "Customer PO#", "Date of receiving PO#", "Customer Required Delivery Date",
    "Sales Document", "Sales Document Item", "Customer Material", "Indus Product",
    "Ordered Quantity", "Delivered Quantity", "Open Quantity",
    "Stock in 3001", "Stock in 3003", "Stock in 1000", "Stock in 2000",
    "To be Manufactured", "Indus Ship Date"
]
FIELD_XML_MAPPING = {
    "Customer Code": "SearchTerm1",
    "Customer PO#": "PurchaseOrderByCustomer",
    "Date of receiving PO#": "CustomerPurchaseOrderDate",
    "Customer Required Delivery Date": "RequestedDeliveryDate",
    "Sales Document": "SalesOrder",
    "Sales Document Item": "SalesOrderItem",
    "Customer Material": "MaterialByCustomer",
    "Indus Product": "Product",
    "Ordered Quantity": "ScheduleLineOrderQuantity",
    "Delivered Quantity": "DeliveredQtyInOrderQtyUnit",
    "Open Quantity": "OpenConfdDelivQtyInOrdQtyUnit",
    "Indus Ship Date": "YY1_IndusShipDate_SDI"
}

def retry_api_call(func, max_retries=3, delay=2, *args, **kwargs):
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            logger.warning(f"API call attempt {attempt+1}/{max_retries} failed: {str(e)}")
            if attempt == max_retries - 1:
                raise
            time.sleep(delay * (2 ** attempt))

def parse_xml_entries_to_records(xml_text):
    ns = {'d': "http://schemas.microsoft.com/ado/2007/08/dataservices", 'm': "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata"}
    root = ET.fromstring(xml_text)
    entries = []
    for entry in root.findall('.//{http://www.w3.org/2005/Atom}entry'):
        props = entry.find('.//{http://schemas.microsoft.com/ado/2007/08/dataservices/metadata}properties')
        rec = {}
        for col in OUTPUT_COLUMNS:
            xml_field = FIELD_XML_MAPPING.get(col, col.replace(" ", ""))
            el = props.find(f'.//{{{ns["d"]}}}{xml_field}') if props is not None else None
            rec[col] = el.text if el is not None else ""
        entries.append(rec)
    return entries

def get_sales_order_data():
    session = requests.Session()
    session.auth = (SAP_CONFIG['username'], SAP_CONFIG['password'])
    session.headers.update({'Accept': 'application/xml'})

    def api_call():
        response = session.get(SAP_CONFIG['sales_order_api_url'], timeout=60)
        response.raise_for_status()
        return response.text

    xml_text = retry_api_call(api_call)
    records = parse_xml_entries_to_records(xml_text)
    df = pd.DataFrame(records)

    logger.info(f"Loaded {len(df)} SO lines from XML SAP analytical API.")

    # Ensure all expected columns exist using set difference and assignment only for missing
    missing_cols = set(OUTPUT_COLUMNS) - set(df.columns)
    for col in missing_cols:
        df[col] = ""

    return df


def parse_stock_xml_entries(xml_text, location_label):
    ns = {'d': "http://schemas.microsoft.com/ado/2007/08/dataservices", 'm': "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata"}
    root = ET.fromstring(xml_text)
    entries = []
    for entry in root.findall('.//{http://www.w3.org/2005/Atom}entry'):
        props = entry.find('.//{http://schemas.microsoft.com/ado/2007/08/dataservices/metadata}properties')
        rec = {}
        rec["Material"] = props.find('.//{http://schemas.microsoft.com/ado/2007/08/dataservices}Material').text if props is not None and props.find('.//{http://schemas.microsoft.com/ado/2007/08/dataservices}Material') is not None else ""
        rec["Plant"] = props.find('.//{http://schemas.microsoft.com/ado/2007/08/dataservices}Plant').text if props is not None and props.find('.//{http://schemas.microsoft.com/ado/2007/08/dataservices}Plant') is not None else ""
        rec["StorageLocation"] = props.find('.//{http://schemas.microsoft.com/ado/2007/08/dataservices}StorageLocation').text if props is not None and props.find('.//{http://schemas.microsoft.com/ado/2007/08/dataservices}StorageLocation') is not None else ""
        rec["Batch"] = props.find('.//{http://schemas.microsoft.com/ado/2007/08/dataservices}Batch').text if props is not None and props.find('.//{http://schemas.microsoft.com/ado/2007/08/dataservices}Batch') is not None else ""
        rec["MatlWrhsStkQtyInMatlBaseUnit"] = props.find('.//{http://schemas.microsoft.com/ado/2007/08/dataservices}MatlWrhsStkQtyInMatlBaseUnit').text if props is not None and props.find('.//{http://schemas.microsoft.com/ado/2007/08/dataservices}MatlWrhsStkQtyInMatlBaseUnit') is not None else "0"
        rec["_location_label"] = location_label
        entries.append(rec)
    return entries

def get_stock_data(materials, batch_mode=True, odata_max=50):
    stock_records = []

    # Normalize input materials once using set and strip
    materials = [str(m).strip() for m in set(materials) if str(m).strip()]
    if not materials:
        logger.warning("No materials found from sales orders, skipping stock queries.")
        return pd.DataFrame([])

    session = requests.Session()
    session.auth = (SAP_CONFIG['username'], SAP_CONFIG['password'])
    session.headers.update({'Accept': 'application/xml'})

    for loc in INVENTORY_OUTPUT:
        plant, storloc, col_label = loc["plant"], loc.get("storloc", None), loc["label"]
        steps = (len(materials) + odata_max - 1) // odata_max

        for i in range(steps):
            mslice = materials[i * odata_max:(i + 1) * odata_max]
            mat_filter = " or ".join([f"Material eq '{m}'" for m in mslice])
            filter_str = f"({mat_filter}) and Plant eq '{plant}'"
            if storloc:
                filter_str += f" and StorageLocation eq '{storloc}'"
            url = SAP_CONFIG['stock_api_url'] + f"?$filter={filter_str}"

            def api_call():
                response = session.get(url, timeout=90)
                response.raise_for_status()
                return response.text

            try:
                xml_result = retry_api_call(api_call)
                result_records = parse_stock_xml_entries(xml_result, col_label)
                stock_records.extend(result_records)
            except Exception as e:
                logger.warning(f"Stock API failed for location {col_label}: {str(e)}")

    stock_df = pd.DataFrame(stock_records)
    logger.info(f"Loaded {len(stock_df)} stock records from XML for locations.")
    return stock_df

def format_output_df(sales_df, stock_df):
    logger.info("Preprocessing and merging for unique SO/Item/Product lines with true MB52 stocks.")
    # Prepare aggregation keys for merging schedule lines
    group_keys = [
        "Customer Code", "Customer PO#", "Sales Document", "Sales Document Item", 
        "Customer Material", "Indus Product"
    ]
    # Numeric and date conversion for merge
    num_cols = ["Ordered Quantity", "Delivered Quantity", "Open Quantity"]
    sales_df[num_cols] = sales_df[num_cols].apply(pd.to_numeric, errors="coerce").fillna(0)

    for date_col in ["Date of receiving PO#", "Customer Required Delivery Date", "Indus Ship Date"]:
        sales_df[date_col] = pd.to_datetime(sales_df[date_col], errors="coerce")
    # Merge schedule lines: sum quantities for the same SO/Item/Product
    agg_dict = {
        "Ordered Quantity": "sum",
        "Delivered Quantity": "sum",
        "Open Quantity": "sum",
        "Date of receiving PO#": "min",
        "Customer Required Delivery Date": "min",
        "Indus Ship Date": "min"
    }
    for col in group_keys:
        agg_dict[col] = "first"
    agg_dict["Indus Ship Date"] = "first"
    merged_df = sales_df.groupby(group_keys, dropna=False, as_index=False).agg(agg_dict)
    # Collect MB52 stock for each FG key (per Indus Product) only once per plant/storage for final output
    logger.info("Building accurate MB52-derived plant/storloc stock for FG in output.")
    stocks_by_fg = {}
    for loc in INVENTORY_OUTPUT:
        label, plant, storloc = loc["label"], loc["plant"], loc["storloc"]
        if not stock_df.empty:
            mask = (stock_df["Plant"] == plant)
            if storloc:
                mask &= (stock_df["StorageLocation"] == storloc)
            stock_filtered = stock_df[mask]
            stocks_by_fg[label] = stock_filtered.groupby("Material")["MatlWrhsStkQtyInMatlBaseUnit"].apply(
                lambda g: pd.to_numeric(g, errors="coerce").fillna(0).sum()
            ).to_dict()
        else:
            stocks_by_fg[label] = {}
        merged_df[label] = merged_df["Indus Product"].map(stocks_by_fg[label]).fillna(0)
        merged_df[label] = pd.to_numeric(merged_df[label], errors="coerce").fillna(0)

    # Dates for Excel output
    # Format date columns to 'DD-MMM-YY' (e.g., 19-Sep-25)
    for date_col in ["Date of receiving PO#", "Customer Required Delivery Date", "Indus Ship Date"]:
        merged_df[date_col] = (
            pd.to_datetime(merged_df[date_col], errors="coerce")
            .dt.strftime("%d-%b-%y")  # '19-Sep-25'
            .str.replace('-Sep-', '-Sept-', regex=False)  # force "Sept" for Excel compatibility
            .replace("NaT", "")
        )
    # Output reorder
    for col in OUTPUT_COLUMNS:
        if col not in merged_df.columns:
            # Use empty string for text and 0 for numeric columns known in advance
            if col in ["To be Manufactured", "Ordered Quantity", "Delivered Quantity", "Open Quantity",
                    "Stock in 3001", "Stock in 3003", "Stock in 1000", "Stock in 2000"]:
                merged_df[col] = 0
            else:
                merged_df[col] = ""
    merged_df = merged_df[OUTPUT_COLUMNS]

        # Filter out rows without customer code/search term
    before_customer_filter = merged_df.shape[0]
    merged_df = merged_df[
        (merged_df["Customer Code"].notna()) & 
        (merged_df["Customer Code"].astype(str).str.strip() != "")
    ]
    after_customer_filter = merged_df.shape[0]
    logger.info(f"Filtered rows without Customer Code: {before_customer_filter - after_customer_filter} rows removed, {after_customer_filter} rows remaining.")

    logger.info(f"Final output lines for Excel: {len(merged_df)}")
    # Remove every line item where Indus Product starts with "6"
    before_rows = merged_df.shape[0]
    merged_df = merged_df[~merged_df["Indus Product"].astype(str).str.startswith("6")]
    after_rows = merged_df.shape[0]
    logger.info(f"Filtered Indus Product starting with '6': {before_rows - after_rows} rows removed, {after_rows} rows remaining.")
    # Ensure date columns are datetime
    merged_df['Date of receiving PO#'] = pd.to_datetime(merged_df['Date of receiving PO#'], errors='coerce')
    merged_df['Indus Ship Date'] = merged_df.apply(compute_indus_ship_date, axis=1)
    merged_df['Indus Ship Date'] = pd.to_datetime(merged_df['Indus Ship Date'], errors='coerce').dt.strftime('%d-%b-%y').replace("NaT", "")
    # Standardize date format for all key date columns (DD-MMM-YY, e.g., 19-Sep-25, with Sep replaced by Sept for Excel compatibility)
    for date_col in ["Date of receiving PO#", "Customer Required Delivery Date", "Indus Ship Date"]:
        merged_df[date_col] = (
            pd.to_datetime(merged_df[date_col], errors="coerce")
            .dt.strftime("%d-%b-%y")                # e.g., 19-Sep-25
            .str.replace('-Sep-', '-Sept-', regex=False)  # Replace Sep with Sept
            .replace("NaT", "")
        )


    return merged_df


def fifo_allocate_stock_to_orders(report, stock_df):
    """
    Robust FIFO allocation: ensure all keys are globally normalized, so
    allocations match even if MB52 and Indus Product codes differ by case, zero, or spaces.
    Plant/location ordering is: 3001, 3003, 1000, 2000.
    """
    logger.info("Starting corrected FIFO allocation for open orders.")

    def norm_key(val):
        return str(val).strip().upper().lstrip("0")

    stock_df = stock_df.copy()
    stock_df["Material_NORM"] = stock_df["Material"].apply(norm_key)
    stock_df["Plant"] = stock_df["Plant"].astype(str).str.strip()
    stock_df["StorageLocation"] = stock_df["StorageLocation"].fillna("").astype(str).str.strip()
    if "StockType" in stock_df.columns:
        stock_df = stock_df[stock_df["StockType"].isin(["Unrestricted", "Quality Inspection"])]

    qty_col = "MatlWrhsStkQtyInMatlBaseUnit"

    # Precompute sums by Material_NORM and Plant and StorageLocation
    # For 3000 plant with StorageLocation 3001 and 3003 separately
    stock_3000 = stock_df[stock_df["Plant"] == "3000"].copy()
    stock_3000[qty_col] = pd.to_numeric(stock_3000[qty_col], errors='coerce')
    stock_3000_grouped = stock_3000.groupby(["Material_NORM", "StorageLocation"])[qty_col].sum()

    # For plants 1000 and 2000 all storage locations summed regardless
    stock_other = stock_df[stock_df["Plant"].isin(["1000", "2000"])].copy()
    stock_other[qty_col] = pd.to_numeric(stock_other[qty_col], errors="coerce")
    stock_other_grouped = stock_other.groupby(["Material_NORM", "Plant"])[qty_col].sum()

    indus_col = "Indus Product"
    all_mats = report[indus_col]

    all_mats_norm = report["Indus Product"].apply(norm_key).unique()

    mat_loc_stock = {}
    for m in all_mats_norm:
        s3001 = stock_3000_grouped.get((m, "3001"), 0.0)
        s3003 = stock_3000_grouped.get((m, "3003"), 0.0)
        s1000 = stock_other_grouped.get((m, "1000"), 0.0)
        s2000 = stock_other_grouped.get((m, "2000"), 0.0)
        mat_loc_stock[m] = [s3001, s3003, s1000, s2000]

    st_cols = ["Stock in 3001", "Stock in 3003", "Stock in 1000", "Stock in 2000"]
    oq_col = "Open Quantity"
    indus_col = "Indus Product"
    updated_cols = {col: [] for col in st_cols}
    tbm_vals = []

    norm_key_vec = np.vectorize(norm_key)

    report_indus_norm = norm_key_vec(report[indus_col].values)
    open_quantities = report[oq_col].fillna(0).values.astype(float)

    # Iterate with itertuples for speed
    for norm_mat, openq in zip(report_indus_norm, open_quantities):
        allocated = [0.0, 0.0, 0.0, 0.0]
        stock_remaining = mat_loc_stock.get(norm_mat, [0.0, 0.0, 0.0, 0.0]).copy()
        required = openq

        for i in range(4):
            alloc_now = min(stock_remaining[i], required)
            allocated[i] = alloc_now
            stock_remaining[i] -= alloc_now
            required -= alloc_now
            if required <= 0:
                break

        mat_loc_stock[norm_mat] = stock_remaining
        for c, v in zip(st_cols, allocated):
            updated_cols[c].append(v)
        tbm_vals.append(required if required > 0 else 0)

    for c in st_cols:
        report[c] = updated_cols[c]
    report["To be Manufactured"] = tbm_vals

    logger.info("Robust FIFO allocation complete with normalized keys.")
    return report


# def update_excel_report(df, excel_file):
#     logger.info("Excel file update started.")
#     if os.path.exists(excel_file):
#         wb = load_workbook(excel_file)
#         # Overwrite sheet(s) if present
#         if "Open Order Report" in wb.sheetnames:
#             del wb["Open Order Report"]
#     else:
#         wb = Workbook()

#     # === Main report sheet ===
#     ws = wb.create_sheet("Open Order Report", 0)
#     ws["R1"] = f"Generated On: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
#     # Write data cell by cell so no row is skipped
#     for row_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):  # Start from Excel row 1
#         for col_idx, value in enumerate(row, 1):  # Start from Excel col 1
#             cell = ws.cell(row=row_idx, column=col_idx, value=value)
#             # Apply header formatting only to the first row
#             if row_idx == 1:
#                 cell.font = Font(bold=True, color="FFFFFF")
#                 cell.fill = PatternFill(start_color="2F4F4F", end_color="2F4F4F", fill_type="solid")
#                 cell.alignment = Alignment(horizontal="center")
#                 cell.border = Border(left=Side(style="thin"), right=Side(style="thin"),
#                                     top=Side(style="thin"), bottom=Side(style="thin"))
#     # Adjust column widths
#     for col in ws.columns:
#         max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
#         ws.column_dimensions[get_excel_column_letter(col[0].column)].width = max(10, min(max_length + 2, 50))
    
#         # Freeze the first row
#     ws.freeze_panes = ws['A2']

#     # Add autofilter to header row
#     ws.auto_filter.ref = "A1:Q1"

#     # Get total rows/cols
#     max_row = ws.max_row
#     max_col = ws.max_column

#     # Define border and alignment styles
#     thin = Side(border_style="thin", color="000000")
#     all_borders = Border(left=thin, right=thin, top=thin, bottom=thin)
#     right_align = Alignment(horizontal="right")

#     # Apply border and right alignment to all data cells (including header)
#     for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
#         for cell in row:
#             cell.border = all_borders
#             cell.alignment = right_align


#     wb.save(excel_file)
#     logger.info(f"Excel Open Order Report updated: {excel_file}")

def main():
    try:
        logger.info("=== Open Order Report Generation START (SAP XML Mode) ===")
        sales_df = get_sales_order_data()
        materials_list = sales_df["Indus Product"].dropna().unique().tolist()
        materials_list = [m for m in materials_list if m]
        stock_df = get_stock_data(materials_list, batch_mode=True)
        final_report_df = format_output_df(sales_df, stock_df)

        from pandas.errors import OutOfBoundsDatetime
        for col in ["Date of receiving PO#", "Customer Required Delivery Date"]:
            dt_short = pd.to_datetime(final_report_df[col], format="%d-%b-%y", errors="coerce")
            dt_long  = pd.to_datetime(final_report_df[col], format="%d-%b-%Y", errors="coerce")
            final_report_df[col + "_sortkey"] = dt_short.combine_first(dt_long)

        final_report_df["Date of receiving PO#_sortkey"] = pd.to_datetime(final_report_df["Date of receiving PO#"], dayfirst=True, errors="coerce")
        final_report_df["Customer Required Delivery Date_sortkey"] = pd.to_datetime(final_report_df["Customer Required Delivery Date"], dayfirst=True, errors="coerce")

        final_report_df = final_report_df.sort_values(
            by=[
                "Indus Product",
                "Date of receiving PO#_sortkey",
                "Customer Required Delivery Date_sortkey",
                "Sales Document",
                "Sales Document Item"
            ],
            ascending=[True, True, True, True, True],
            kind="mergesort"
        ).reset_index(drop=True)

        final_report_df = final_report_df.drop(columns=["Date of receiving PO#_sortkey", "Customer Required Delivery Date_sortkey"])

        # FIFO allocation
        final_report_df_fifo = fifo_allocate_stock_to_orders(final_report_df, stock_df)

        # Save locally
        # update_excel_report(final_report_df_fifo, SAP_CONFIG["fixed_excel_file"])

        upload_success = update_excel_data_via_graph_api(
            final_report_df_fifo,
            "201c767f-23cc-4eff-b76c-47a275706531,094b21f9-6ed4-47c3-a751-3b217a18b98e",
            "Indus Open Order Report for trial.xlsx",
            "Open Order Report"
        )

        if not upload_success:
            logger.error("Upload to SharePoint failed.")

        n_unique_so = final_report_df["Sales Document"].nunique()
        n_unique_products = final_report_df["Indus Product"].nunique()
        n_nonzero_mfg = final_report_df.loc[final_report_df["To be Manufactured"] > 0].shape[0]
        logger.info(f"Summary: Unique Sales Orders: {n_unique_so}, Unique Products: {n_unique_products}, Rows To Be Mfg>0: {n_nonzero_mfg}")
        print(f"Summary -- Unique SOs: {n_unique_so}, Unique Products: {n_unique_products}, ToBeMfg>0: {n_nonzero_mfg}")
        logger.info("=== Open Order Report Generation COMPLETE ===")
        print(f"✅ {len(final_report_df)} unique order lines exported to {SAP_CONFIG['fixed_excel_file']}")
    except Exception as e:
        logger.error(f"FAILED: {str(e)}")
        logger.error("Traceback:\n" + traceback.format_exc())
        print("❌ See open_order_report.log for details.")


if __name__ == "__main__":

    main()
