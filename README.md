# Indus Open Order Report v2 — Customer & EXW Variants

Active deployment of the Indus Open Order Report automation. Reads open sales order data from SAP S/4HANA (via SharePoint Excel), processes it per-customer and per-Incoterm (EXW), and sends formatted HTML email reports with Excel attachments via Microsoft Graph API. Runs on a schedule via GitHub Actions.

## What It Does

- Authenticates to Microsoft 365 using MSAL client credentials (credentials in GitHub Secrets / `.env`)
- Reads open order data from the SharePoint-hosted SAP export workbook via Graph API
- Filters, sorts, and groups orders by customer or Incoterm (EXW)
- Builds colour-coded HTML email tables showing order number, material, quantities, delivery dates, and status
- Attaches a formatted `.xlsx` file per report run
- Sends to the sales and operations distribution list

## Scripts

| File | Description |
|------|-------------|
| `Indus-Open-Order-Report.py` | Base script — full open order report for all customers |
| `Indus-Open-Order-Report-CUST.py` | Customer-split variant — one email per customer |
| `Indus-Open-Order-Report-EXW.py` | EXW Incoterm variant — filtered for EXW orders only |

Run the variant matching your reporting requirement. All three share the same authentication and SharePoint read logic.

## Difference from `Indus_OOR_update`

`Indus_OOR_update` is the original single-script version (v1). This repository (`Indus_OOR_update_2`) is the **current active deployment** with:
- Customer-split reporting (`-CUST` variant)
- EXW Incoterm filtering (`-EXW` variant)
- `.env.template` for local credential setup
- Improved `.gitignore` and secrets management

## Tech Stack

- **Python 3.11**
- `msal` — Microsoft OAuth2 token acquisition
- `requests` — Graph API calls
- `pandas` — Data processing and filtering
- `openpyxl` — Formatted Excel attachment generation
- **GitHub Actions** — scheduled execution

## Setup

```bash
pip install -r requirements.txt
```

Copy `.env.template` to `.env` and fill in your credentials:

```
CLIENT_ID=...
CLIENT_SECRET=...
TENANT_ID=...
SITE_ID=...
```

For GitHub Actions, add the same values as repository secrets.

---
*Indus International FZC — Internal Automation Tool*
