from flask import Flask, request, jsonify, send_from_directory, render_template
import os
import re
import json
import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import matching_engine

app = Flask(__name__, static_folder="static", template_folder="templates")

UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
GENERATED_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "generated")
CATALOG_JSON_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "catalog_text.json")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)

def detect_columns(row_vals):
    desc_col_idx = None
    qty_col_idx = None
    
    exclude_desc = [
        "hsn", "mat code", "material code", "color", "angle", "qty", "quantity", 
        "unit", "uom", "price", "rate", "amount", "cost", "make", "brand", 
        "manufacturer", "driver", "led", "image", "picture", "photo", 
        "location", "area", "sr.", "s.no", "serial", "remark", "status", 
        "inspection", "packing", "stickering", "branding", "watt", "temp", 
        "cct", "dimension", "size", "height", "cutout", "dia", "voltage"
    ]
    
    desc_priority_1 = ["description", "item name", "item description", "particulars", "specification"]
    desc_priority_2 = ["item", "luminaire", "product description", "particular"]
    desc_priority_3 = ["product", "name", "model", "code", "type"]
    
    # First, find Qty column
    for c_idx, val in enumerate(row_vals):
        if val:
            val_lower = str(val).lower().strip()
            if val_lower == "qty" or val_lower == "quantity":
                qty_col_idx = c_idx + 1
                break
                
    if qty_col_idx is None:
        for c_idx, val in enumerate(row_vals):
            if val:
                val_lower = str(val).lower().strip()
                if "qty" in val_lower or "quantity" in val_lower:
                    qty_col_idx = c_idx + 1
                    break

    # Now, find Description column
    for priority in [desc_priority_1, desc_priority_2, desc_priority_3]:
        for c_idx, val in enumerate(row_vals):
            if val:
                val_lower = str(val).lower().strip()
                is_excluded = any(ex in val_lower for ex in exclude_desc)
                if not is_excluded:
                    if any(kw in val_lower for kw in priority):
                        desc_col_idx = c_idx + 1
                        return desc_col_idx, qty_col_idx
                        
    # Fallback to first non-excluded column
    for c_idx, val in enumerate(row_vals):
        if val:
            val_lower = str(val).lower().strip()
            is_excluded = any(ex in val_lower for ex in exclude_desc)
            if not is_excluded and c_idx + 1 != qty_col_idx:
                desc_col_idx = c_idx + 1
                break
                
    return desc_col_idx, qty_col_idx


@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/upload-boq", methods=["POST"])
def upload_boq():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
        
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Empty filename"}), 400
        
    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)
    
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheet = wb.active
        
        # Detect header row
        header_row_idx = None
        desc_col_idx = None
        qty_col_idx = None
        
        # Scan first 15 rows to find headers
        for r_idx in range(1, 16):
            row_vals = [sheet.cell(row=r_idx, column=c_idx).value for c_idx in range(1, sheet.max_column + 1)]
            row_str = " ".join([str(v).lower() for v in row_vals if v is not None])
            if "qty" in row_str or "quantity" in row_str:
                d_col, q_col = detect_columns(row_vals)
                if d_col and q_col:
                    header_row_idx = r_idx
                    desc_col_idx = d_col
                    qty_col_idx = q_col
                    break
                
        if not header_row_idx:
            header_row_idx = 4
            desc_col_idx = 3 # Col C
            qty_col_idx = 4  # Col D
            
        items = []
        sr_no = 1
        
        # Read data rows
        for r_idx in range(header_row_idx + 1, sheet.max_row + 1):
            desc_val = sheet.cell(row=r_idx, column=desc_col_idx).value
            qty_val = sheet.cell(row=r_idx, column=qty_col_idx).value
            
            if desc_val is not None:
                desc_str = str(desc_val).strip()
                if desc_str == "" or desc_str.lower().startswith("sales team") or desc_str.lower().startswith("total"):
                    continue
                
                # Parse Qty and Unit
                qty = 0
                unit = "Nos"
                if qty_val is not None:
                    qty_str = str(qty_val).strip()
                    unit_match = re.search(r'(\d+)\s*[- ]*\s*([a-zA-Z]+)', qty_str)
                    if unit_match:
                        qty = int(unit_match.group(1))
                        unit = unit_match.group(2).capitalize()
                        if unit == "Nos" or unit == "No":
                            unit = "Nos"
                    else:
                        try:
                            qty = int(float(qty_str))
                        except ValueError:
                            qty = qty_str
                            
                if "mtr" in desc_str.lower():
                    unit = "Mtr"
                    
                # Run matching engine
                parsed_info = matching_engine.parse_product_code(desc_str)
                specs = matching_engine.lookup_catalog_database(parsed_info, CATALOG_JSON_PATH)
                
                product_description = specs["product_description"]
                
                # Determine default body color
                body_color = "Black"
                if "white" in desc_str.lower() or "white" in product_description.lower():
                    body_color = "White"
                elif "mil finish" in desc_str.lower() or "mil finish" in product_description.lower() or "silver" in desc_str.lower():
                    body_color = "Mil Finish"
                    
                # Calculate driver details
                driver_make = specs["driver_make"]
                driver_wattage = specs["driver_wattage"]
                driver_qty = qty
                
                if unit.lower() == "mtr":
                    driver_make = "Constant Voltage 24V"
                    driver_wattage = "150W"
                    try:
                        meters = float(qty)
                        driver_qty = max(1, int(meters / 12))
                    except Exception:
                        driver_qty = 1
                
                driver_details = f"{driver_make} - {driver_wattage}" if driver_make != "NA" else "NA"
                led_details = f"{specs['led_make']} - {parsed_info.get('cct', '4000K')}" if specs["led_make"] != "NA" else "NA"
                
                items.append({
                    "id": sr_no,
                    "boq_description": desc_str,
                    "boq_qty": qty,
                    "unit": unit,
                    "gls_code": desc_str, # Keep the WHOLE product code as mentioned in BOQ
                    "product_description": product_description,
                    "body_color": body_color,
                    "driver_details": driver_details,
                    "driver_qty": driver_qty,
                    "led_details": led_details,
                    "accessories": specs["accessories"],
                    "page": specs["page"],
                    "matched_by": specs["matched_by"],
                    "remarks": ""
                })
                sr_no += 1
                
        return jsonify({
            "filename": file.filename,
            "project_name": file.filename.split("__")[0].replace("BOQ", "").replace("_", " ").strip(),
            "items": items
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Failed to parse BOQ Excel: {str(e)}"}), 500

@app.route("/api/search-catalog", methods=["GET"])
def search_catalog():
    query = request.args.get("q", "").strip()
    if not query:
        return jsonify([])
        
    results = []
    if os.path.exists(CATALOG_JSON_PATH):
        try:
            with open(CATALOG_JSON_PATH, "r", encoding="utf-8") as f:
                catalog = json.load(f)
            for page_num, text in catalog.items():
                if query.lower() in text.lower():
                    idx = text.lower().find(query.lower())
                    start = max(0, idx - 80)
                    end = min(len(text), idx + 80)
                    snippet = text[start:end].replace("\n", " ").strip()
                    results.append({
                        "page": int(page_num),
                        "snippet": f"... {snippet} ..."
                    })
        except Exception as e:
            print("Error searching catalog:", e)
            
    results = sorted(results, key=lambda x: x["page"])
    return jsonify(results[:15])

@app.route("/api/generate-iwo", methods=["POST"])
def generate_iwo():
    data = request.json
    if not data:
        return jsonify({"error": "No data provided"}), 400
        
    client_name = data.get("client_name", "Food Square")
    iwo_no = data.get("iwo_no", "P176")
    iwo_date = data.get("iwo_date", datetime.date.today().strftime("%d/%m/%y"))
    delivery_date = data.get("delivery_date", "")
    sales_name = data.get("sales_name", "Sales-01")
    items = data.get("items", [])
    
    try:
        # Create styled Excel sheet
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "IWO "
        ws.views.sheetView[0].showGridLines = True
        
        # Styles
        font_title_company = Font(name="Arial", size=16, bold=True)
        font_title_iwo = Font(name="Arial", size=20, bold=True)
        font_header = Font(name="Arial", size=9, bold=True)
        font_bold_data = Font(name="Arial", size=9, bold=True)
        font_data = Font(name="Arial", size=9)
        
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        
        thin_side = Side(border_style="thin", color="000000")
        border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        
        col_widths = {
            "A": 3,   # Buffer column
            "B": 8,   # SR. NO.
            "C": 12,  # Image placeholder (kept empty)
            "D": 25,  # GLS product code
            "E": 45,  # product description
            "F": 12,  # Body Color
            "G": 10,  # Qty
            "H": 10,  # Unit
            "I": 25,  # DRIVER MAKE and Wattage
            "J": 12,  # DRIVER QTY
            "K": 20,  # LED MAKE and CCT
            "L": 20,  # Accessories
            "M": 22,  # stickering & branding details
            "N": 18,  # packing details
            "O": 25   # REMARKS
        }
        
        for col_let, width in col_widths.items():
            ws.column_dimensions[col_let].width = width
            
        ws.cell(row=1, column=2, value="GLS-SPA").font = font_title_company
        ws.cell(row=2, column=2, value="INTERNAL WORK ORDER").font = font_title_iwo
        
        ws.cell(row=3, column=3, value="Client Name :- ").font = font_bold_data
        ws.cell(row=3, column=4, value=client_name).font = font_bold_data
        ws.cell(row=3, column=5, value=f"IWO NO :- {iwo_no}").font = font_bold_data
        ws.cell(row=3, column=6, value="IWO DATE :-     ").font = font_bold_data
        ws.cell(row=3, column=7, value=iwo_date).font = font_bold_data
        ws.cell(row=3, column=9, value="PO NO. :-").font = font_bold_data
        ws.cell(row=3, column=12, value="DELIVERY DATE :-  ").font = font_bold_data
        ws.cell(row=3, column=13, value=delivery_date).font = font_bold_data
        
        ws.row_dimensions[1].height = 25
        ws.row_dimensions[2].height = 30
        ws.row_dimensions[3].height = 20
        ws.row_dimensions[4].height = 25
        
        headers = [
            "SR. NO.", "Image", "GLS product code", "product description",
            "Body Color", "Qty", "Unit", "DRIVER MAKE and Wattage ", "DRIVER QTY",
            "LED MAKE and CCT ", "Accessories", "stickering  and branding details ",
            "packing details ", "REMARKS"
        ]
        
        for c_idx, h_text in enumerate(headers):
            cell = ws.cell(row=4, column=c_idx + 2, value=h_text)
            cell.font = font_header
            cell.border = border_all
            cell.alignment = align_center
            
        # Write data rows
        curr_row = 5
        for item in items:
            ws.row_dimensions[curr_row].height = 35 # normal, non-congested height
            
            vals = [
                item.get("id"),
                "", # Image column remains empty as requested
                item.get("gls_code"),
                item.get("product_description"),
                item.get("body_color"),
                item.get("boq_qty"),
                item.get("unit"),
                item.get("driver_details"),
                item.get("driver_qty"),
                item.get("led_details"),
                item.get("accessories") or "",
                "GLS SPA", # Default stickering
                "", # packing details remains empty
                item.get("remarks") or ""
            ]
            
            for c_idx, val in enumerate(vals):
                cell = ws.cell(row=curr_row, column=c_idx + 2, value=val)
                cell.font = font_data
                cell.border = border_all
                if c_idx in [0, 1, 4, 5, 6, 8, 9, 10, 11, 12]:
                    cell.alignment = align_center
                else:
                    cell.alignment = align_left
                    
            curr_row += 1
            
        # Footer block (Signatures)
        ws.row_dimensions[curr_row + 1].height = 25
        cell_sales = ws.cell(row=curr_row + 1, column=2, value=f"SALES TEAM NAME AND SIGN  - {sales_name}")
        cell_sales.font = font_bold_data
        
        cell_prod = ws.cell(row=curr_row + 1, column=8, value="PRODUCTION SIGN.:-")
        cell_prod.font = font_bold_data
        
        cell_qc = ws.cell(row=curr_row + 1, column=12, value="QC SIGN.:-")
        cell_qc.font = font_bold_data
        
        # Save file
        filename = f"IWO-{iwo_no}-{client_name.replace(' ', '_')}.xlsx"
        save_path = os.path.join(GENERATED_FOLDER, filename)
        wb.save(save_path)
        
        return jsonify({
            "success": True,
            "filename": filename,
            "download_url": f"/api/download-iwo/{filename}"
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Failed to generate IWO: {str(e)}"}), 500

@app.route("/api/download-iwo/<filename>", methods=["GET"])
def download_iwo(filename):
    return send_from_directory(GENERATED_FOLDER, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)
