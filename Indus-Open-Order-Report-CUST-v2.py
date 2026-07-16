import requests
import pandas as pd
import logging
import time
import os
import xml.etree.ElementTree as ET
from openpyxl.utils import get_column_letter as get_excel_column_letter
from datetime import datetime
import numpy as np
import traceback
import msal


class AuthConfig:
    CLIENT_ID = 'a2239be7-12cd-442d-983f-ea7e316ef767'
    CLIENT_SECRET = 'sho8Q~sanlYEX3SKXjYJb0NSXL.gPRpdi63-tcmp'

authority = 'https://login.microsoftonline.com/ed97e9bb-e119-4bd4-ab00-307d64bdf908'

def get_access_token():
    auth_config = AuthConfig()
    app = msal.ConfidentialClientApplication(
        auth_config.CLIENT_ID, authority=authority, client_credential=auth_config.CLIENT_SECRET
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

        session.patch(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/range(address='G2:G{num_rows}')/format",
            headers=api_headers,
            json={"numberFormat": [["@"]]}
        )

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
            header_range = f"A1:T1"

            # PATCH fill color
            header_fill_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/range(address='{header_range}')/format/fill"
            header_fill_payload = {"color": "#2F4F4F"}
            fill_response = session.patch(header_fill_url, headers=api_headers, json=header_fill_payload)
            if fill_response.status_code in [200,201]:
                logger.info("Header fill (background) applied.")
            else:
                logger.warning(f"Header fill failed: {fill_response.status_code}, {fill_response.text}")
            
            # PATCH special fill color for S1 header only: Indus Ship Date (Production)
            special_header_fill_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/range(address='S1')/format/fill"
            special_header_fill_payload = {"color": "#036B8E"}  # light blue
            special_fill_response = session.patch(special_header_fill_url, headers=api_headers, json=special_header_fill_payload)
            if special_fill_response.status_code in [200, 201]:
                logger.info("Special light blue fill applied to S1 header.")
            else:
                logger.warning(f"Special S1 header fill failed: {special_fill_response.status_code}, {special_fill_response.text}")

            # PATCH font color and bold
            header_font_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/range(address='{header_range}')/format/font"
            header_font_payload = {"bold": True, "color": "#FFFFFF"}
            font_response = session.patch(header_font_url, headers=api_headers, json=header_font_payload)
            if font_response.status_code in [200,201]:
                logger.info("Header font color/bold applied.")
            else:
                logger.warning(f"Header font update failed: {font_response.status_code}, {font_response.text}")
            
            # Set center alignment for header row
            header_align_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/range(address='{header_range}')/format"
            header_align_payload = {"horizontalAlignment": "Center"}
            align_resp = session.patch(header_align_url, headers=api_headers, json=header_align_payload)
            if align_resp.status_code in [200, 201]:
                logger.info("Header center alignment applied.")
            else:
                logger.warning(f"Header alignment failed: {align_resp.status_code}, {align_resp.text}")

            
            # Set fixed column widths and enable text wrapping
            column_widths = {
                'A': 70, 'B': 100, 'C': 90, 'D': 120, 'E': 90, 'F': 80, 'G': 130, 'H': 90,
                'I': 50, 'J': 50, 'K': 50, 'L': 40, 'M': 40, 'N': 40, 'O': 40, 'P': 80,
                'Q': 90, 'R': 80, 'S': 110, 'T': 150
            }

            for col_letter, width in column_widths.items():
                col_range = f"{col_letter}:{col_letter}"
                col_format_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/range(address='{col_range}')/format"
                col_format_payload = {
                    "columnWidth": width,
                    "wrapText": True
                }
                col_resp = session.patch(col_format_url, headers=api_headers, json=col_format_payload)
                if col_resp.status_code in [200, 201]:
                    logger.info(f"Column {col_letter} width set to {width} with text wrapping")
                else:
                    logger.warning(f"Column {col_letter} formatting failed: {col_resp.status_code}")

            
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
    # T1 = 20th column, 1st row (timestamp shifted right after inserting new column S)
    update_datetime = datetime.now().strftime('Generated On: %d-%m-%Y %H:%M:%S')
    payload = {
        "values": [[update_datetime]]
    }
    update_url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/"
        f"workbook/worksheets('{worksheet_name}')/range(address='T1')"
    )
    r = session.patch(update_url, headers=api_headers, json=payload)
    if r.status_code in [200, 201]:
        logger.info("Timestamp for last update written to cell T1")
        return True
    else:
        logger.error(f"Failed to update timestamp in cell T1: {r.status_code} {r.text}")
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

def compute_indus_ship_date_production(row):
    po_date = pd.to_datetime(row.get('Date of receiving PO#'), errors='coerce')
    crd_date = pd.to_datetime(row.get('Customer Required Delivery Date'), errors='coerce')

    po_plus_6_weeks = po_date + pd.Timedelta(weeks=6) if pd.notna(po_date) else pd.NaT
    crd_minus_12_weeks = crd_date - pd.Timedelta(weeks=12) if pd.notna(crd_date) else pd.NaT

    if pd.notna(po_plus_6_weeks) and pd.notna(crd_minus_12_weeks):
        return max(po_plus_6_weeks, crd_minus_12_weeks)
    elif pd.notna(po_plus_6_weeks):
        return po_plus_6_weeks
    elif pd.notna(crd_minus_12_weeks):
        return crd_minus_12_weeks
    else:
        return np.nan

# LOGGING SETUP
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("open_order_report.log", encoding="utf-8"), logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# CONFIGURATION
SAP_CONFIG = {
    "username": "INDUS_OPEN_ORDER",
    "password": "w6fo3zfq/bDYr}8>BU@{i3oFd[dq#Zv\/eL&}SKF",
    "sales_order_api_url": "https://my409486-api.s4hana.cloud.sap/sap/opu/odata/sap/YY1_INDUSOPENORDERAPI_CDS/YY1_IndusOpenOrderAPI",
    "stock_api_url": "https://my409486-api.s4hana.cloud.sap/sap/opu/odata/sap/API_MATERIAL_STOCK_SRV/A_MatlStkInAcctMod",
    "fixed_excel_file": "Open_Order_Report_MASTER.xlsx"
}
INVENTORY_OUTPUT = [
    {"label": "3001", "plant": "3000", "storloc": "3001"},
    {"label": "3003", "plant": "3000", "storloc": "3003"},
    {"label": "1000", "plant": "1000", "storloc": None},
    {"label": "2000", "plant": "2000", "storloc": None},
]
OUTPUT_COLUMNS = [
    "Customer Code", "Customer PO#", "Date of receiving PO#", "Customer Required Delivery Date",
    "Sales Document", "Sales Document Item", "Customer Material", "Indus Product",
    "Ordered Quantity", "Delivered Quantity", "Open Quantity",
    "3001", "3003", "1000", "2000",
    "To be Manufactured", "Indus Ship Date", "Actual Customer", "Indus Ship Date (Production)"
]
FIELD_XML_MAPPING = {
    "Customer Code": "SearchTerm1",
    "Customer PO#": "PurchaseOrderByCustomer",
    "Date of receiving PO#": "CustomerPurchaseOrderDate",
    "Customer Required Delivery Date": "YY1_CUSTOMERITEMDATE_SDI",
    "Sales Document": "SalesOrder",
    "Sales Document Item": "SalesOrderItem",
    "Customer Material": "MaterialByCustomer",
    "Indus Product": "Product",
    "Ordered Quantity": "ScheduleLineOrderQuantity",
    "Delivered Quantity": "DeliveredQtyInOrderQtyUnit",
    "Open Quantity": "OpenConfdDelivQtyInOrdQtyUnit",
    "Indus Ship Date": "YY1_IndusShipDate_SDI",
    "Incoterms": "IncotermsClassification"
}

def compute_actual_customer(df):
    """
    Compute Actual Customer column:
    - Use Customer Code if it's not "PLS"
    - If Customer Code is "PLS", find another Customer Code for the same Indus Product that is not "PLS"
    - If even that fails, use static mapping provided
    """

    # STATIC mapping from your supplied list (override/fallback)
    static_mapping = {
        "CHW0272": "RDT", "CEW0413": "RDT", "CHW0240": "RDT", "CCN0672": "RDT", "CEW2082": "BAC", "CEN1343": "BAC", "CCN1265": "BAC",
        "CES1788": "BAC", "CCG1689": "BAC", "CCN1272": "BAC", "CES1319": "BAC", "CES1318": "BAC", "CES1723": "BAC", "CEU1733": "BAC",
        "CES2755": "BAC", "CEV0415": "BAC", "CEV0417": "BAC", "CEN1600": "BAC", "CCJ1264": "BAC", "CES1721": "BAC", "CEW1599": "BAC",
        "CCN1266": "BAC", "CES1787": "BAC", "CEV0424": "BAC", "CES1789": "BAC", "CES1383": "BAC", "CEU2827": "BAC", "CES1320": "BAC",
        "CCR1684": "BAC", "CHX0277": "RDT", "CJX0350": "RDT", "CEX0426": "RDT", "CCW0768": "RDT", "CEW0635": "RDT", "CJX0163": "RDT",
        "CCR0860": "RDT", "AJW0427": "WFI", "AEW1813": "WFI", "AEW1814": "WFI", "AJW0428": "WFI", "AJW0429": "WFI", "AJW0430": "WFI",
        "AEW1858": "WFI", "AEW1857": "WFI", "AEW1582": "WFI", "CHX0605": "RDE", "CCR1749": "SCT", "CCQ1748": "SCT", "CCQ2053": "SCT",
        "CCR1997": "ICE", "CCR1998": "ICE", "CCR1999": "ICE", "CCR2000": "ICE", "ACN1334": "ICE", "CEU2552": "BAC", "AJX0462": "ENR"
    }

    # First, build mapping from the real data as before
    non_pls_rows = df[df['Customer Code'] != 'PLS']
    if not non_pls_rows.empty:
        non_pls_mapping = non_pls_rows.groupby('Indus Product')['Customer Code'].first().to_dict()
    else:
        non_pls_mapping = {}

    def get_actual_customer(row):
        customer_code = row['Customer Code']
        indus_product = row['Indus Product']
        # Priority 1: If not PLS, use directly
        if customer_code != 'PLS':
            return customer_code
        # Priority 2: If PLS and another Customer Code exists for this product, use it
        if indus_product in non_pls_mapping:
            return non_pls_mapping[indus_product]
        # Priority 3: Use static mapping if available
        if indus_product in static_mapping:
            return static_mapping[indus_product]
        # Final fallback—still say 'PLS'
        return 'PLS'

    df['Actual Customer'] = df.apply(get_actual_customer, axis=1)
    return df

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
    ns = {'d': "http://schemas.microsoft.com/ado/2007/08/dataservices", 
          'm': "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata"}
    root = ET.fromstring(xml_text)
    entries = []
    
    # Debug: Log available fields from first entry
    first_entry = root.find('.//{http://www.w3.org/2005/Atom}entry')
    if first_entry:
        props = first_entry.find('.//{http://schemas.microsoft.com/ado/2007/08/dataservices/metadata}properties')
        if props is not None:
            available_fields = [child.tag.replace('{http://schemas.microsoft.com/ado/2007/08/dataservices}', '') 
                              for child in props]
            logger.info(f"Available XML fields in API response: {', '.join(available_fields[:20])}")  # Log first 20 fields
            
            # Check if Incoterms field exists
            incoterms_variants = ['IncotermsClassification', 'Incoterms', 'IncoTerms', 
                                 'IncotermsClassif', 'YY1_IncotermsClassification_SDH']
            found_incoterms = [field for field in available_fields if any(inc.lower() in field.lower() for inc in ['incoterm', 'inco'])]
            if found_incoterms:
                logger.info(f"Found Incoterms-related fields: {found_incoterms}")
            else:
                logger.warning("NO Incoterms field found in API response! Incoterms will be empty for all records.")
    
    for entry in root.findall('.//{http://www.w3.org/2005/Atom}entry'):
        props = entry.find('.//{http://schemas.microsoft.com/ado/2007/08/dataservices/metadata}properties')
        rec = {}
        
        # Parse standard output columns
        for col in OUTPUT_COLUMNS:
            if col == "Indus Ship Date (Production)":
                rec[col] = ""
                continue

            xml_field = FIELD_XML_MAPPING.get(col, col.replace(" ", ""))
            el = props.find(f'.//{{{ns["d"]}}}{xml_field}') if props is not None else None
            rec[col] = el.text if el is not None else ""
        
        # Parse Incoterms field - try multiple possible field names
        incoterms_value = ""
        possible_incoterms_fields = [
            'IncotermsClassification',  # Standard SAP field
            'Incoterms',                # Simplified name
            'IncoTerms',                # Alternative casing
            'YY1_IncotermsClassification_SDH',  # Custom field at header
            'YY1_Incoterms_SDH',        # Custom field variant
            'to_SalesOrder/IncotermsClassification'  # Navigation to header
        ]
        
        for field_name in possible_incoterms_fields:
            incoterms_el = props.find(f'.//{{{ns["d"]}}}{field_name}') if props is not None else None
            if incoterms_el is not None and incoterms_el.text:
                incoterms_value = incoterms_el.text.strip().upper()
                break
        
        rec["Incoterms"] = incoterms_value
        
        # Log only non-empty Incoterms for debugging
        if incoterms_value:
            logger.info(f"Found Incoterms: {incoterms_value} for SO: {rec.get('Sales Document', 'N/A')}, "
                       f"Item: {rec.get('Sales Document Item', 'N/A')}")
        
        entries.append(rec)
    
    # Summary logging
    exw_count = sum(1 for e in entries if e.get('Incoterms', '') == 'EXW')
    non_empty_incoterms = sum(1 for e in entries if e.get('Incoterms', ''))
    logger.info(f"Parsed {len(entries)} entries from XML")
    logger.info(f"Incoterms data: {non_empty_incoterms} records with Incoterms, {exw_count} with EXW")
    
    if non_empty_incoterms == 0:
        logger.error("*** CRITICAL: No Incoterms data found in ANY record! ***")
        logger.error("*** The API might not expose Incoterms at item level. ***")
        logger.error("*** You may need to use Sales Order Header API or add Incoterms to CDS view. ***")
    
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

    # === CRITICAL FIX: Ensure Incoterms column exists BEFORE groupby ===
    if "Incoterms" not in sales_df.columns:
        logger.error("✗ CRITICAL: Incoterms column NOT in sales_df! Adding empty column.")
        sales_df["Incoterms"] = ""
    else:
        exw_in_salesdf = (sales_df["Incoterms"] == "EXW").sum()
        total_with_incoterms = sales_df["Incoterms"].notna().sum()
        logger.info(f"✓ Incoterms column found in sales_df: {exw_in_salesdf} EXW out of {total_with_incoterms} records")

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
        "Indus Ship Date": "first",
        "Incoterms": "first"  # ← MUST BE HERE to survive groupby
    }

    for col in group_keys:
        agg_dict[col] = "first"

    logger.info(f"Aggregation keys: {list(agg_dict.keys())}")

    merged_df = sales_df.groupby(group_keys, dropna=False, as_index=False).agg(agg_dict)

    # CRITICAL: Verify Incoterms survived the groupby
    incoterms_after_merge = merged_df.get("Incoterms", None)
    if incoterms_after_merge is not None:
        non_empty_incoterms = incoterms_after_merge.notna().sum()
        exw_count_after = (incoterms_after_merge == "EXW").sum()
        logger.info(f"✓ Incoterms preserved after groupby: {non_empty_incoterms} records with data, {exw_count_after} with EXW")
    else:
        logger.error("✗ CRITICAL ERROR: Incoterms LOST after groupby operation!")
        merged_df["Incoterms"] = ""  # Fallback


    logger.info(f"Incoterms preserved in merged data: {merged_df['Incoterms'].notna().sum()} rows with Incoterms")

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
                    "3001", "3003", "1000", "2000"]:
                merged_df[col] = 0
            else:
                merged_df[col] = ""

    # Compute Actual Customer column before final reordering
    merged_df = compute_actual_customer(merged_df)

    # Ensure Incoterms is included in the output for FIFO allocation
    columns_for_fifo = OUTPUT_COLUMNS.copy()
    if "Incoterms" not in columns_for_fifo:
        columns_for_fifo.append("Incoterms")

    merged_df = merged_df[columns_for_fifo]
    logger.info(f"Columns passed to FIFO: {list(merged_df.columns)}")


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
    merged_df['Customer Required Delivery Date'] = pd.to_datetime(merged_df['Customer Required Delivery Date'], errors='coerce')

    merged_df['Indus Ship Date'] = merged_df.apply(compute_indus_ship_date, axis=1)
    merged_df['Indus Ship Date (Production)'] = merged_df.apply(compute_indus_ship_date_production, axis=1)

    merged_df['Indus Ship Date'] = pd.to_datetime(merged_df['Indus Ship Date'], errors='coerce')
    merged_df['Indus Ship Date (Production)'] = pd.to_datetime(merged_df['Indus Ship Date (Production)'], errors='coerce')

    # Standardize date format for all key date columns (DD-MMM-YY, e.g., 19-Sep-25, with Sep replaced by Sept for Excel compatibility)
    for date_col in ["Date of receiving PO#", "Customer Required Delivery Date", "Indus Ship Date", "Indus Ship Date (Production)"]:
        merged_df[date_col] = (
            pd.to_datetime(merged_df[date_col], errors="coerce")
            .dt.strftime("%d-%b-%y")
            .str.replace('-Sep-', '-Sept-', regex=False)
            .replace("NaT", "")
        )
        # Remove Incoterms from final output (used internally for allocation only)
    if "Incoterms" in merged_df.columns:
        logger.info("Removing Incoterms column from final output (used only for internal allocation logic)")
        # Store it temporarily for the allocation function
        merged_df["_Incoterms_internal"] = merged_df["Incoterms"]
        # Don't drop yet - will be handled after allocation

    return merged_df


def fifo_allocate_stock_to_orders(report, stock_df):
    """
    Robust FIFO allocation: ensure all keys are globally normalized, so
    allocations match even if MB52 and Indus Product codes differ by case, zero, or spaces.
    Plant/location ordering is: 3001, 3003, 1000, 2000.

    ALLOCATION RULE: Only orders whose Customer Code is in the ALLOWED_3001_3003_CUSTOMERS
    whitelist will receive stock allocation from 3001 and 3003.
    All other customer codes will NOT receive any allocation from 3001 or 3003.
    """
    logger.info("="*80)
    logger.info("FIFO ALLOCATION PHASE - Starting corrected FIFO allocation with Customer Code whitelist logic")
    logger.info("="*80)

    # ===== WHITELIST: Only these Customer Codes get 3001 and 3003 allocation =====
    ALLOWED_3001_3003_CUSTOMERS = {
        "BAC", "FHP", "CPT", "DRT", "DUK", "ENR", "GRK", "HOA", "HEL", "HTE",
        "HVA", "HYD", "ICE", "ILA", "ISR", "ITW", "IWA", "JMF", "KYN", "LNW",
        "NRC", "NYL", "SPX", "PER", "RAC", "RDT", "RPC", "RSG", "SCT", "STG",
        "STR", "SVA", "TAI", "TFS", "TMC", "WFI", "ETE", "WBT", "RDE", "LER",
        "ACE", "SPS", "MLT", "OHO", "SPX", "BON", "AML", "NYL", "PLK", "FHP",
        "LER", "ITT", "MTX", "MTA", "NVL"
    }

    def norm_key(val):
        return str(val).strip().upper().lstrip("0")

    # ===== Customer Code Validation =====
    logger.info(f"Customer Code column present in report: {'Customer Code' in report.columns}")
    logger.info(f"Report shape: {report.shape}, Columns: {list(report.columns)}")

    if "Customer Code" not in report.columns:
        logger.error("✗ CRITICAL ERROR: Customer Code column NOT FOUND in report!")
        report["Customer Code"] = ""

    # Extract Customer Code values for allocation logic
    customer_code_series = report["Customer Code"].fillna("").astype(str).str.strip().str.upper()
    customer_code_values = customer_code_series.values

    # Log Customer Code whitelist statistics
    total_records = len(customer_code_values)
    whitelisted_count = sum(1 for cc in customer_code_values if cc in ALLOWED_3001_3003_CUSTOMERS)
    non_whitelisted_count = total_records - whitelisted_count

    logger.info(f"Total orders: {total_records}")
    logger.info(f"Whitelisted customers (eligible for 3001/3003): {whitelisted_count}")
    logger.info(f"Non-whitelisted customers (excluded from 3001/3003): {non_whitelisted_count}")

    # ===== Continue with rest of FIFO logic =====
    stock_df = stock_df.copy()
    stock_df["Material_NORM"] = stock_df["Material"].apply(norm_key)
    stock_df["Plant"] = stock_df["Plant"].astype(str).str.strip()
    stock_df["StorageLocation"] = stock_df["StorageLocation"].fillna("").astype(str).str.strip()

    if "StockType" in stock_df.columns:
        stock_df = stock_df[stock_df["StockType"].isin(["Unrestricted", "Quality Inspection"])]

    qty_col = "MatlWrhsStkQtyInMatlBaseUnit"

    # Precompute sums by Material_NORM and Plant and StorageLocation
    stock_3000 = stock_df[stock_df["Plant"] == "3000"].copy()
    stock_3000[qty_col] = pd.to_numeric(stock_3000[qty_col], errors='coerce')
    stock_3000_grouped = stock_3000.groupby(["Material_NORM", "StorageLocation"])[qty_col].sum()

    stock_other = stock_df[stock_df["Plant"].isin(["1000", "2000"])].copy()
    stock_other[qty_col] = pd.to_numeric(stock_other[qty_col], errors="coerce")
    stock_other_grouped = stock_other.groupby(["Material_NORM", "Plant"])[qty_col].sum()

    indus_col = "Indus Product"
    all_mats_norm = report["Indus Product"].apply(norm_key).unique()

    mat_loc_stock = {}
    for m in all_mats_norm:
        s3001 = stock_3000_grouped.get((m, "3001"), 0.0)
        s3003 = stock_3000_grouped.get((m, "3003"), 0.0)
        s1000 = stock_other_grouped.get((m, "1000"), 0.0)
        s2000 = stock_other_grouped.get((m, "2000"), 0.0)
        mat_loc_stock[m] = [s3001, s3003, s1000, s2000]

    st_cols = ["3001", "3003", "1000", "2000"]
    oq_col = "Open Quantity"

    updated_cols = {col: [] for col in st_cols}
    tbm_vals = []

    norm_key_vec = np.vectorize(norm_key)
    report_indus_norm = norm_key_vec(report[indus_col].values)
    open_quantities = report[oq_col].fillna(0).values.astype(float)

    non_whitelisted_allocation_count = 0  # Track how many non-whitelisted orders were processed

    # ===== MAIN ALLOCATION LOOP WITH CUSTOMER CODE WHITELIST LOGIC =====
    for idx, (norm_mat, openq, customer_code) in enumerate(zip(report_indus_norm, open_quantities, customer_code_values)):
        allocated = [0.0, 0.0, 0.0, 0.0]
        stock_remaining = mat_loc_stock.get(norm_mat, [0.0, 0.0, 0.0, 0.0]).copy()
        required = openq

        # CRITICAL: Check if this customer is allowed to receive 3001/3003 stock
        is_allowed_3001_3003 = (customer_code in ALLOWED_3001_3003_CUSTOMERS)

        if not is_allowed_3001_3003 and openq > 0:
            logger.debug(f"[Row {idx}] Non-whitelisted Customer Code '{customer_code}' - "
                         f"SO: {report.iloc[idx].get('Sales Document', 'N/A')}, "
                         f"Item: {report.iloc[idx].get('Sales Document Item', 'N/A')}, "
                         f"Material: {norm_mat}, Open Qty: {openq:.2f} → Skipping 3001 and 3003")
            non_whitelisted_allocation_count += 1

        # Allocation loop with Customer Code whitelist logic
        for i in range(4):
            # CRITICAL: Skip 3001 (index 0) and 3003 (index 1) if Customer Code is NOT in whitelist
            if not is_allowed_3001_3003 and i in [0, 1]:
                allocated[i] = 0.0
                if required > 0 and idx < 10:  # Log first 10 for debugging
                    logger.debug(f"[Row {idx}] Skipping location {st_cols[i]} for non-whitelisted customer '{customer_code}'")
                continue

            alloc_now = min(stock_remaining[i], required)
            allocated[i] = alloc_now
            stock_remaining[i] -= alloc_now
            required -= alloc_now

            if alloc_now > 0 and idx < 10:  # Log first 10 for debugging
                logger.debug(f"[Row {idx}] Allocated {alloc_now:.2f} from {st_cols[i]}")

            if required <= 0:
                break

        mat_loc_stock[norm_mat] = stock_remaining

        for c, v in zip(st_cols, allocated):
            updated_cols[c].append(v)

        to_be_mfg = openq - allocated[0] - allocated[1] - allocated[2]
        tbm_vals.append(to_be_mfg if to_be_mfg > 0 else 0)

        # Detailed logging for non-whitelisted orders
        if not is_allowed_3001_3003 and openq > 0:
            logger.info(f"[NON-WHITELISTED-ALLOC] Customer: '{customer_code}', "
                        f"SO: {report.iloc[idx].get('Sales Document', 'N/A')}, "
                        f"Item: {report.iloc[idx].get('Sales Document Item', 'N/A')}, "
                        f"Material: {norm_mat}, Open Qty: {openq:.2f} → "
                        f"[3001: {allocated[0]:.2f}, 3003: {allocated[1]:.2f}, 1000: {allocated[2]:.2f}, 2000: {allocated[3]:.2f}], "
                        f"TBM: {to_be_mfg:.2f}")

    # Assign results back to report
    for c in st_cols:
        report[c] = updated_cols[c]

    report["To be Manufactured"] = tbm_vals

    # ===== VALIDATION & LOGGING =====
    logger.info("="*80)
    logger.info("FIFO ALLOCATION VALIDATION RESULTS:")
    logger.info("="*80)
    logger.info(f"Total orders processed: {total_records}")
    logger.info(f"Non-whitelisted orders (excluded from 3001/3003): {non_whitelisted_allocation_count}")
    logger.info(f"Whitelisted orders (eligible for 3001/3003): {total_records - non_whitelisted_allocation_count}")

    # Verify no non-whitelisted customer got 3001/3003 allocation
    non_whitelisted_incorrectly_allocated = 0
    for idx, cc in enumerate(customer_code_values):
        if cc not in ALLOWED_3001_3003_CUSTOMERS:
            if updated_cols["3001"][idx] > 0 or updated_cols["3003"][idx] > 0:
                non_whitelisted_incorrectly_allocated += 1
                logger.warning(f"✗ Non-whitelisted customer '{cc}' at row {idx} received 3001/3003 allocation!")

    if non_whitelisted_incorrectly_allocated > 0:
        logger.error(f"✗ VALIDATION FAILED: {non_whitelisted_incorrectly_allocated} non-whitelisted orders incorrectly received 3001/3003 allocation!")
    else:
        logger.info(f"✓ VALIDATION PASSED: All non-whitelisted customers correctly excluded from 3001/3003 allocation")

    logger.info("="*80)
    logger.info("Robust FIFO allocation complete with normalized keys and Customer Code whitelist logic.")
    logger.info("="*80)

    return report

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

        # FIFO allocation with Incoterms logic
        # NEW CODE (FIXED):
        # CRITICAL: Verify Incoterms exists before FIFO allocation
        logger.info(f"Before FIFO allocation - Incoterms column present: {'Incoterms' in final_report_df.columns}")
        if "Incoterms" in final_report_df.columns:
            exw_before_fifo = (final_report_df["Incoterms"] == "EXW").sum()
            non_empty_before_fifo = final_report_df["Incoterms"].notna().sum()
            logger.info(f"Before FIFO allocation - {exw_before_fifo} EXW orders out of {non_empty_before_fifo} with Incoterms")
        else:
            logger.error("✗ ERROR: Incoterms missing before FIFO allocation!")

        # FIFO allocation with Incoterms logic
        final_report_df_fifo = fifo_allocate_stock_to_orders(final_report_df, stock_df)

        # Verify FIFO processing
        if "Incoterms" in final_report_df_fifo.columns:
            logger.info("✓ Incoterms preserved after FIFO allocation")
        else:
            logger.warning("Incoterms column removed or missing after FIFO")

        # Remove Incoterms column from final output (used only internally)
        if "Incoterms" in final_report_df_fifo.columns:
            logger.info("Removing Incoterms column from final Excel output")
            final_report_df_fifo = final_report_df_fifo.drop(columns=["Incoterms"])

        if "_Incoterms_internal" in final_report_df_fifo.columns:
            final_report_df_fifo = final_report_df_fifo.drop(columns=["_Incoterms_internal"])

        # Save locally
        # update_excel_report(final_report_df_fifo, SAP_CONFIG["fixed_excel_file"])

        upload_success = update_excel_data_via_graph_api(
            final_report_df_fifo,
            "201c767f-23cc-4eff-b76c-47a275706531,094b21f9-6ed4-47c3-a751-3b217a18b98e",
            "Indus Open Order Report New.xlsx",
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
