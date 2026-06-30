# GLS-SPA Internal Workflow App

An automated internal tool built for GLS-SPA Lighting to streamline the generation of Internal Work Orders (IWO) and Proforma Invoices (PI) from raw Bill of Quantities (BOQ) spreadsheets. 

This application eliminates the tedious manual data entry required to map messy client BOQs into the standard GLS catalog specifications, and it provides a sleek, modern interface for finalizing quotes and generating PDF invoices.

## Features

### 1. BOQ to IWO Converter (Smart Catalog Matching Engine)
- **Excel Parsing:** Upload raw `.xlsx` BOQ sheets from clients. The app automatically detects headers, descriptions, and quantities.
- **Smart Text Extraction Engine (`matching_engine.py`):** Uses Regex and string-matching to dissect messy BOQ strings (e.g. `"GS-DE-FDL-49-7W-4000K-36Deg"`) into precise component-level attributes:
  - Product Series (e.g. GS-DE-FDL)
  - Wattage / Dimensions (e.g. 7W, Cutout 56mm)
  - Color Temperature (CCT)
  - Beam Angle (Deg)
  - Housing Color
- **Automated Specs Generation:** Automatically maps the extracted parameters to the GLS catalog (hardcoded specs, PDF page mapping, and exact descriptions) to populate `product_description`, `led_details`, `driver_details`, and `accessories`.
- **Fuzzy Catalog Search:** If the engine can't find an exact match, users can use the "Remap Catalog Specs" modal to perform a fuzzy text search against `catalog_db.json` and manually bind the specification block.
- **IWO Excel Export:** Export the fully populated and finalized workspace directly to a clean `.xlsx` Internal Work Order sheet.

### 2. Modern Proforma Invoice Generator
- **Visual Document Editor:** A sleek, side-by-side interactive invoice layout inspired by modern invoicing platforms (Refrens-style).
- **Client Database:** Save and load frequent Buyer details (Name, Address, GSTIN, Contact) using the built-in JSON database to speed up billing.
- **Dynamic Calculation:** Automatically calculates Line Amounts, GST (18%), and Grand Totals.
- **Number to Words Engine:** Automatically converts the numerical Grand Total into Indian Rupee words (e.g., "Ten Thousand Rupees Only").
- **PDF Generation (`reportlab`):** With a single click of "Export PDF", the backend generates a highly professional, print-ready PDF Proforma Invoice using ReportLab, retaining all formatting, company details, terms and conditions, and bank details.

## Tech Stack

- **Backend:** Python, Flask (`app.py`)
- **Frontend:** Vanilla HTML, CSS (`style.css`), JavaScript (`script.js`)
- **Excel Parsing/Writing:** `openpyxl`
- **PDF Generation:** `reportlab`
- **Database:** Flat JSON files (`buyers_db.json`, `catalog_db.json`, `GLS_SPA_Product_Database.csv`)

## File Structure Overview

- `app.py`: The main Flask server. Handles routing, file uploads, PDF generation logic, and API endpoints.
- `matching_engine.py`: The brain of the operation. Contains the `CATALOG_SPECS` dictionary and the parsing logic to translate BOQ strings into catalog matches.
- `templates/index.html`: The Single Page Application (SPA) frontend. Contains the sidebar, the BOQ view, and the modern Invoice Generator view.
- `static/style.css`: The styling engine. Handles the sleek UI, dark/light theme switching, and the modern invoice layout.
- `static/script.js`: The frontend logic. Handles tab switching, asynchronous fetch requests, dynamic table rendering, state management (`workspaceItems`), and calculation logic.
- `generate_catalog_db.py` / `extract_text.py`: Utility scripts used to scrape text from the massive GLS PDF catalog and index it into `catalog_db.json` for the fuzzy search functionality.

## How to Run Locally

1. **Prerequisites:** 
   - Python 3.8+ installed.
   - `pip install flask openpyxl reportlab PyPDF2 pdfplumber`
2. **Start the Server:**
   - Run the flask application:
     ```bash
     python app.py
     ```
3. **Access the App:**
   - Open your browser and navigate to `http://localhost:8000` (or whichever port Flask runs on locally).

## Standard Workflow

1. **Upload:** Go to the "BOQ -> IWO" tab. Upload a client `.xlsx` BOQ file.
2. **Review & Edit:** The matching engine will process the rows. Review the matches. If a row is red or has incorrect specs, click the magnifying glass to remap it manually using the catalog search.
3. **Export IWO:** Click "Export to Excel (IWO)" to get the standard internal sheet.
4. **Generate Invoice:** Switch to the "Proforma Invoice" tab. Your uploaded items will carry over. 
5. **Bill To:** Select a saved buyer or type in new details and click "Save Buyer".
6. **Price:** Enter the rates for the products (Default is ₹0).
7. **Export:** Click "Export PDF" to generate the final Proforma Invoice!
