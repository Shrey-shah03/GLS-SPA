import re
import json
import os

DATA_DIR = os.environ.get("DATA_DIR", os.path.dirname(os.path.abspath(__file__)))
CATALOG_SPECS_FILE = os.path.join(DATA_DIR, "catalog_specs.json")

try:
    with open(CATALOG_SPECS_FILE, "r", encoding="utf-8") as f:
        _db = json.load(f)
        CATALOG_SPECS = _db.get("CATALOG_SPECS", {})
        ALIASES = _db.get("ALIASES", {})
except Exception as e:
    CATALOG_SPECS = {}
    ALIASES = {}

def parse_product_code(text):
    """
    Parses a string description/code from BOQ and extracts core product components.
    """
    if not text:
        return {}
    
    text = str(text).strip()
    
    # Try to extract a product code like pattern: GS-XXX-XXX-XXX
    # Support spaces in accessories (e.g. GS-MT-End Cap, GS-MT-LIVE END, GS-MT-POWER ADAPTER, GS POWER SUPPLY 48V/100W)
    match = re.search(r'\b(GS[- ][A-Z0-9\-xX/*_]+(?:\s+(?:End\s+Cap|LIVE\s+END|POWER\s+ADAPTER|POWER\s+ADAPTOR|SUPPLY\s+48V/100W|End|Cap|Live|End|Power|Adapter|Supply))?)\b', text, re.IGNORECASE)
    
    if match:
        raw_code = match.group(1).strip()
    else:
        # Fallback to simple pattern or first word
        code_match = re.search(r'(GS-[A-Z0-9\-]+(?:MTR|M|FT|W|K)?)', text, re.IGNORECASE)
        raw_code = code_match.group(1) if code_match else text.split()[0]
    
    parts = raw_code.split("-")
    
    model = ""
    size_code = ""
    wattage = ""
    cct = ""
    
    # Extract CCT (usually matches digits + K e.g. 4000K or 3000K or 6500k)
    cct_match = re.search(r'(\d{4}K)', text, re.IGNORECASE)
    if cct_match:
        cct = cct_match.group(1).upper()
        
    # Extract Wattage (usually matches digits + W or multiplier e.g. 30W or 2x25W)
    watt_match = re.search(r'((?:\d+\s*[xX]\s*)?\d+W(?:/MTR)?)', text, re.IGNORECASE)
    if watt_match:
        wattage = watt_match.group(1).upper()
        
    # Rebuild model and size
    if len(parts) >= 3:
        if parts[0].upper() == "GS":
            if parts[1].upper() in ["DE", "IP", "SF", "LINEA", "DR", "SMART", "T4", "LED", "MT"]:
                model = "-".join(parts[0:3]).upper()
                if len(parts) >= 4:
                    if not re.search(r'\d+W', parts[3]) and not re.search(r'\d+K', parts[3]):
                        size_code = parts[3].upper()
            else:
                model = "-".join(parts[0:2]).upper()
                if len(parts) >= 3:
                    if not re.search(r'\d+W', parts[2]) and not re.search(r'\d+K', parts[2]):
                        size_code = parts[2].upper()
    else:
        model = parts[0].upper()
        if len(parts) > 1:
            size_code = parts[1].upper()
            
    if not wattage:
        for p in parts:
            if "W" in p.upper() and any(c.isdigit() for c in p):
                wattage = p.upper()
                break
                
    if not cct:
        for p in parts:
            if "K" in p.upper() and any(c.isdigit() for c in p):
                cct = p.upper()
                break

    return {
        "raw_code": raw_code,
        "original_text": text,
        "model": model,
        "size_code": size_code,
        "wattage": wattage,
        "cct": cct
    }

def _lookup_catalog_database_impl(parsed_info, catalog_json_path=None):
    """
    Looks up specs in the catalog database.
    """
    model = parsed_info.get("model", "")
    size = parsed_info.get("size_code", "")
    wattage = parsed_info.get("wattage", "")
    cct = parsed_info.get("cct", "4000K")
    raw_code = parsed_info.get("raw_code", "")
    original_text = parsed_info.get("original_text", "").upper()
    
    # Try alias first
    alias_key = f"{model}-{size}" if size else model
    # Clean up alias key to check loose aliases (like GS-T4-TRACK-2 MTR -> GS-T4-TRACK)
    for k_alias in sorted(ALIASES.keys(), key=len, reverse=True):
        if alias_key.startswith(k_alias) or raw_code.upper().startswith(k_alias) or k_alias in original_text:
            model, size = ALIASES[k_alias]
            break
            
    if alias_key in ALIASES:
        model, size = ALIASES[alias_key]
        
    specs = None
    variant_info = None
    
    # Try compound key model-size check (e.g. model=GS-IP-BL, size=2X2 -> check key GS-IP-BL-2X2)
    compound_key = f"{model}-{size}" if size else model
    if compound_key in CATALOG_SPECS:
        model = compound_key
        size = "DEFAULT"
        
    # Try exact CATALOG_SPECS match
    if model in CATALOG_SPECS:
        specs = CATALOG_SPECS[model]
        variants = specs.get("variants", {})
        if size in variants:
            variant_info = variants[size]
        elif wattage in variants:
            variant_info = variants[wattage]
        elif "DEFAULT" in variants:
            variant_info = variants["DEFAULT"]
        else:
            variant_info = list(variants.values())[0] if variants else {}
            
    if specs:
        desc_temp = variant_info.get("description_template", specs["description_template"])
        placeholders = re.findall(r'\{([A-Za-z0-9_]+)\}', desc_temp)
        format_args = {}
        for ph in placeholders:
            if ph in ["L", "l", "length", "Length"]:
                len_match = re.search(r'length\s*[-:\s]*\s*(\d+)', original_text, re.IGNORECASE)
                if len_match:
                    format_args[ph] = len_match.group(1)
                else:
                    len_match2 = re.search(r'(\d+)\s*(?:mm)?\s*length', original_text, re.IGNORECASE)
                    if len_match2:
                        format_args[ph] = len_match2.group(1)
                    elif ph in variant_info:
                        format_args[ph] = variant_info[ph]
                    else:
                        format_args[ph] = "?"
            elif ph in variant_info:
                format_args[ph] = variant_info[ph]
            elif ph == "CCT":
                format_args[ph] = cct
            else:
                format_args[ph] = "?"
        
        try:
            prod_desc = desc_temp.format(**format_args)
        except Exception:
            prod_desc = desc_temp
            
        driver_watt = variant_info.get("wattage", wattage or specs.get("variants", {}).get("DEFAULT", {}).get("wattage", "10W"))
        
        return {
            "page": specs["page"],
            "gls_code": f"{model}-{size}" if size and size != "DEFAULT" else model,
            "product_description": prod_desc,
            "led_make": specs["led_make"],
            "driver_make": specs["driver_make"],
            "driver_wattage": driver_watt,
            "unit": specs["unit"],
            "accessories": specs.get("accessories", "Standard"),
            "matched_by": "exact_catalog_specs"
        }
        
    # Check if we can search the newly mapped catalog database (GLS_SPA_catalog.json)
    gls_catalog_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "GLS_SPA_catalog.json")
    if os.path.exists(gls_catalog_path) and raw_code.strip():
        try:
            with open(gls_catalog_path, "r", encoding="utf-8") as f:
                gls_catalog = json.load(f)
            
            query = raw_code.strip().upper().replace(" ", "")
            best_match = None
            
            for item in gls_catalog:
                prod_code = item.get("product_code", "").strip().upper()
                prod_code_clean = prod_code.replace(" ", "")
                if query == prod_code_clean or query in prod_code_clean or prod_code_clean in query:
                    best_match = item
                    break
                    
            if best_match:
                page_num = best_match.get("pdf_page", 0)
                code_matched = best_match.get("product_code", raw_code)
                return {
                    "page": page_num,
                    "gls_code": code_matched,
                    "product_description": f"GLS-SPA Luminaire, IP20.",
                    "led_make": "Bridgelux",
                    "driver_make": "Fulham",
                    "driver_wattage": wattage or "10W",
                    "unit": "Mtr" if "mtr" in original_text or "linear" in original_text else "Nos",
                    "accessories": "Standard",
                    "matched_by": f"gls_spa_catalog_json_page_{page_num}"
                }
        except Exception as e:
            print("Error matching GLS_SPA_catalog.json:", e)

    # Check if we can search the raw text database
    if catalog_json_path and os.path.exists(catalog_json_path):
        try:
            with open(catalog_json_path, "r", encoding="utf-8") as f:
                catalog = json.load(f)
            best_page = None
            best_score = 0
            search_query = raw_code.lower()
            
            for page_num, text in catalog.items():
                if search_query in text.lower():
                    best_page = page_num
                    break
                words = [w for w in re.split(r'\W+', search_query) if len(w) > 2]
                score = sum(1 for w in words if w in text.lower())
                if score > best_score:
                    best_score = score
                    best_page = page_num
                    
            if best_page:
                page_text = catalog[str(best_page)]
                driver_make = "Fulham" if "fulham" in page_text.lower() else "Constant Current"
                led_make = "Bridgelux" if "bridgelux" in page_text.lower() else "SMD LED"
                unit = "Mtr" if "mtr" in page_text.lower() or "linear" in page_text.lower() else "Nos"
                
                return {
                    "page": int(best_page),
                    "gls_code": raw_code,
                    "product_description": f"GLS-SPA Luminaire, IP20.",
                    "led_make": led_make,
                    "driver_make": driver_make,
                    "driver_wattage": wattage or "10W",
                    "unit": unit,
                    "accessories": "Standard",
                    "matched_by": f"catalog_text_search_page_{best_page}"
                }
        except Exception as e:
            print("Error matching catalog json:", e)
            
    # Generic fallback
    return {
        "page": 0,
        "gls_code": raw_code,
        "product_description": f"GLS-SPA Luminaire, IP20.",
        "led_make": "Bridgelux",
        "driver_make": "Fulham",
        "driver_wattage": wattage or "10W",
        "unit": "Nos",
        "accessories": "Standard",
        "matched_by": "generic_fallback"
    }

def double_verify_with_csv(matched_result, raw_code):
    import csv
    csv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "GLS_SPA_Product_Database.csv")
    if not os.path.exists(csv_path) or not raw_code.strip():
        return matched_result
        
    try:
        query_clean = raw_code.strip().upper().replace(" ", "")
        csv_match = None
        with open(csv_path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                code_val = row.get("Product Code", "").strip().upper()
                code_clean = code_val.replace(" ", "")
                if query_clean == code_clean or query_clean in code_clean or code_clean in query_clean:
                    csv_match = row
                    break
                    
        if csv_match:
            size = csv_match.get("Size", "")
            lens = csv_match.get("Lens", "")
            housing = csv_match.get("Housing", "")
            ip = csv_match.get("IP Rating", "")
            beam = csv_match.get("Beam Angle", "")
            driver_make = csv_match.get("Driver Make", "")
            driver_watt = csv_match.get("Driver Wattage", "")
            
            parts = []
            if housing:
                parts.append(f"Housing - {housing}")
            if lens:
                parts.append(lens)
            if size:
                parts.append(f"Dimensions - {size}")
            if beam:
                parts.append(f"Beam Angle - {beam}")
            if ip:
                parts.append(ip)
                
            if parts:
                desc = ", ".join(parts) + "."
                matched_result["product_description"] = desc
                
            if driver_make and driver_make != "NA" and driver_make.strip():
                matched_result["driver_make"] = driver_make
            if driver_watt and driver_watt != "NA" and driver_watt.strip():
                matched_result["driver_wattage"] = driver_watt
                
            matched_result["matched_by"] = matched_result.get("matched_by", "") + "_csv_verified"
            
    except Exception as e:
        print("Error during CSV double verification:", e)
        
    return matched_result

def lookup_catalog_database(parsed_info, catalog_json_path=None):
    res = _lookup_catalog_database_impl(parsed_info, catalog_json_path)
    return double_verify_with_csv(res, parsed_info.get("raw_code", ""))
