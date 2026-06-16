import re
import json
import os

# Database of catalog pages and specs
CATALOG_SPECS = {
    "GS-DE-FDL": {
        "page": 69,
        "name": "Deep Recessed Fixed Downlight",
        "description_template": "Housing - Die-Cast Aluminum, Reflector and Clear Glass, Dimensions - Dia - {D}mm, Height - {H}mm, Cutout - {C}mm, IP20.",
        "led_make": "Bridgelux",
        "driver_make": "Fulham",
        "unit": "Nos",
        "accessories": "Toggle clips",
        "variants": {
            "49": {"D": 64, "C": 56, "H": 39, "wattage": "7W"},
            "58": {"D": 74, "C": 68, "H": 45, "wattage": "12W"},
            "64": {"D": 79, "C": 73, "H": 60, "wattage": "18W"},
            "84": {"D": 107, "C": 97, "H": 79, "wattage": "30W"},
            "100": {"D": 123, "C": 113, "H": 83, "wattage": "40W"},
        }
    },
    "GS-DE-INT": {
        "page": 77,
        "name": "Deep Recessed Intersecting / Adjustable Downlight",
        "description_template": "Housing - Die-Cast Aluminum, Reflector and Clear Glass, Dimensions - Cutout - {C}mm, Height - {H}mm, IP20.",
        "led_make": "Bridgelux",
        "driver_make": "Fulham",
        "unit": "Nos",
        "accessories": "Toggle clips",
        "variants": {
            "84": {"C": 84, "H": 89, "wattage": "30W"},
            "100": {"C": 100, "H": 85, "wattage": "40W"},
            "4E-100": {"C": 100, "H": 50, "wattage": "20W"},
        }
    },
    "GS-DE-MGL": {
        "page": 75,
        "name": "Movable General Light",
        "description_template": "Housing - Die-Cast Aluminum, Reflector and Clear Glass, Dimensions - Cutout - 265*135mm, Outer - 300*170mm, IP20.",
        "led_make": "Bridgelux",
        "driver_make": "Fulham",
        "unit": "Nos",
        "accessories": "Toggle clips",
        "variants": {
            "84": {"wattage": "2X19W", "description_template": "Housing - Die-Cast Aluminum, Reflector and Clear Glass, Dimensions - Cutout - 265*135mm, Outer - 300*170mm, 40 Degree, IP20."},
            "100": {"wattage": "2X20W"}
        }
    },
    "GS-FDL": {
        "page": 82,
        "name": "Fixed Downlight",
        "description_template": "Housing - Die-Cast Aluminum, Reflector and Clear Glass, Dimensions - Dia - {D}mm, Height - {H}mm, Cutout - {C}mm, IP20.",
        "led_make": "Bridgelux",
        "driver_make": "Fulham",
        "unit": "Nos",
        "accessories": "Toggle clips",
        "variants": {
            "10W": {"D": 70, "C": 60, "H": 58, "wattage": "10W"},
            "15W": {"D": 82, "C": 72, "H": 68, "wattage": "15W"},
            "24W": {"D": 110, "C": 96, "H": 88, "wattage": "24W"},
            "30W": {"D": 118, "C": 107, "H": 100, "wattage": "30W"},
            "42W": {"D": 125, "C": 110, "H": 114, "wattage": "42W"}
        }
    },
    "GS-DR-SK": {
        "page": 88,
        "name": "Deep Recessed Spot Downlight",
        "description_template": "Housing - Die-Cast Aluminum, Reflector and Clear Glass, Dimensions - Dia - {D}mm, Height - {H}mm, Cutout - {C}mm, IP20.",
        "led_make": "Bridgelux",
        "driver_make": "Fulham",
        "unit": "Nos",
        "accessories": "Toggle clips",
        "variants": {
            "3W": {"D": 42, "C": 35, "H": 38, "wattage": "3W"},
            "5W": {"D": 58, "C": 50, "H": 65, "wattage": "5W"},
            "12W": {"D": 85, "C": 75, "H": 70, "wattage": "12W"},
            "15W": {"D": 85, "C": 75, "H": 83, "wattage": "15W"},
            "20W": {"D": 95, "C": 83, "H": 92, "wattage": "20W"},
            "30W": {"D": 129, "C": 113, "H": 121, "wattage": "30W"}
        }
    },
    "GS-IP-BL": {
        "page": 116,
        "name": "IP54 Backlit Panel Light",
        "description_template": "Housing - CRCA Powder Coated Body, Opal Diffuser, Dimensions - {D}mm, IP54.",
        "led_make": "SMD LED",
        "driver_make": "Constant Current",
        "unit": "Nos",
        "accessories": "Recessed Clips",
        "variants": {
            "2X2": {"D": "595 x 595 x 35", "wattage": "36W", "description_template": "Housing - CRCA Powder Coated Body, Opal Diffuser, Dimensions - 595 x 595 x 35mm, IP54."},
            "1X4": {"D": "1195 x 295 x 35", "wattage": "36W"},
            "2X4": {"D": "1195 x 595 x 35", "wattage": "72W"},
        }
    },
    "GS-SMART-BL-RD": {
        "page": 115,
        "name": "Smart Backlit Round Panel",
        "description_template": "Housing - Die-Cast Aluminum, Opal Diffuser, Dimensions - Dia - {D}mm, Height - {H}mm, Cutout - {C}mm, IP20.",
        "led_make": "SMD LED",
        "driver_make": "Constant Current",
        "unit": "Nos",
        "accessories": "Spring clips",
        "variants": {
            "15W": {"D": 140, "C": 126, "H": 30, "wattage": "15W"},
            "20W": {"D": 160, "C": 144, "H": 30, "wattage": "20W"},
            "24W": {"D": 185, "C": 172, "H": 30, "wattage": "24W"}
        }
    },
    "GS-SMART-BL-SQ": {
        "page": 115,
        "name": "Smart Backlit Square Panel",
        "description_template": "Housing - Die-Cast Aluminum, Opal Diffuser, Dimensions - {D}x{D}mm, Height - {H}mm, Cutout - {C}x{C}mm, IP20.",
        "led_make": "SMD LED",
        "driver_make": "Constant Current",
        "unit": "Nos",
        "accessories": "Spring clips",
        "variants": {
            "15W": {"D": 140, "C": 126, "H": 30, "wattage": "15W"},
            "20W": {"D": 160, "C": 144, "H": 30, "wattage": "20W"},
            "24W": {"D": 185, "C": 170, "H": 30, "wattage": "24W"}
        }
    },
    "GS-SMART-SF-RD": {
        "page": 115,
        "name": "Smart Surface Round Panel",
        "description_template": "Housing - Die-Cast Aluminum, Opal Diffuser, Dimensions - Dia - 145mm, Height - 50mm, IP20.",
        "led_make": "SMD LED",
        "driver_make": "Constant Current",
        "unit": "Nos",
        "accessories": "Surface Mounting Bracket",
        "variants": {
            "15W": {"wattage": "15W"},
            "20W": {"wattage": "20W"},
            "24W": {"wattage": "24W"},
            "DEFAULT": {"wattage": "15W"}
        }
    },
    "GS-SF-80X150": {
        "page": 149,
        "name": "Surface Cylinder Downlight",
        "description_template": "Housing - Extruded Aluminum, Reflector and Clear Glass, Dimensions - Dia - 80mm, Height - 150mm, IP20.",
        "led_make": "Bridgelux",
        "driver_make": "Fulham",
        "unit": "Nos",
        "accessories": "Surface Canopy",
        "variants": {
            "15W": {"wattage": "15W"},
            "20W": {"wattage": "20W"},
            "DEFAULT": {"wattage": "15W"}
        }
    },
    "GS-1714": {
        "page": 178,
        "name": "Aluminium Profile 17x14mm",
        "description_template": "Housing - Extruded Aluminium Profile, PMMA Opal Diffuser, Finish - Silver Anodized, with End caps and Mounting Clips, Dimensions - 17mm x 14mm, IP20.",
        "led_make": "Bridgelux SMD 2835",
        "driver_make": "Constant Voltage 24V",
        "unit": "Mtr",
        "accessories": "Mounting Clips and End caps",
        "variants": {
            "DEFAULT": {"wattage": "10W/Mtr"}
        }
    },
    "GS-LINEA-5275": {
        "page": 194,
        "name": "Linear Extruded Profile 52x75mm",
        "description_template": "Housing - Extruded Aluminum Profile, Opal / Prismatic Diffuser, with Endcaps and Suspension wires, Dimensions - Width 52mm, Height 75mm, IP20.",
        "led_make": "SMD LED",
        "driver_make": "Constant Current",
        "unit": "Nos",
        "accessories": "Suspension Kit 2MTR",
        "variants": {
            "DEFAULT": {"wattage": "20W"}
        }
    },
    "GS-LED-T5": {
        "page": 237,
        "name": "LED Batten T5",
        "description_template": "Housing - PC Body, Opal Diffuser, Dimensions - L - 1200mm, W - 20mm, H - 35mm, IP20.",
        "led_make": "SMD LED",
        "driver_make": "Constant Current",
        "unit": "Nos",
        "accessories": "Mounting Clamps",
        "variants": {
            "20W": {"wattage": "20W"},
            "DEFAULT": {"wattage": "20W"}
        }
    },
    "GS-T4-TRACK": {
        "page": 137,
        "name": "T4 Track Rail",
        "description_template": "Housing - Extruded Aluminium Profile, white/black finish, length as specified, standard accessories, IP20.",
        "led_make": "NA",
        "driver_make": "NA",
        "unit": "Nos",
        "accessories": "Live End & End Cap",
        "variants": {
            "DEFAULT": {"wattage": "NA"}
        }
    },
    "GS-T4-I": {
        "page": 137,
        "name": "T4 Track I Jointer",
        "description_template": "Housing - T4 Track I Jointer, white/black finish, for inline track connectivity, IP20.",
        "led_make": "NA",
        "driver_make": "NA",
        "unit": "Nos",
        "accessories": "None",
        "variants": {
            "DEFAULT": {"wattage": "NA"}
        }
    },
    "GS-T4-POWER": {
        "page": 137,
        "name": "T4 Track Power Adapter",
        "description_template": "Housing - T4 Track Power Adapter / Live End, white/black finish, for track electrical feed, IP20.",
        "led_make": "NA",
        "driver_make": "NA",
        "unit": "Nos",
        "accessories": "None",
        "variants": {
            "DEFAULT": {"wattage": "NA"}
        }
    },
    "GS-T4-END": {
        "page": 137,
        "name": "T4 Track End Cap",
        "description_template": "Housing - T4 Track End Cap, white/black finish, to cover track rail terminals, IP20.",
        "led_make": "NA",
        "driver_make": "NA",
        "unit": "Nos",
        "accessories": "None",
        "variants": {
            "DEFAULT": {"wattage": "NA"}
        }
    },
    "SUSPENSION": {
        "page": 118,
        "name": "Suspension Wire Kit",
        "description_template": "Suspension wire for Linear and Track lights, length 2mtr with ceiling canopy and mounting clips.",
        "led_make": "NA",
        "driver_make": "NA",
        "unit": "Nos",
        "accessories": "Ceiling Canopy",
        "variants": {
            "DEFAULT": {"wattage": "NA"}
        }
    },
    "GS-STRETCH": {
        "page": 11,
        "name": "Stretch Fabric Light Box",
        "description_template": "Housing - Aluminium Mounting Frame, Stretch Fabric Diffuser, custom dimensions, IP20.",
        "led_make": "SMD LED Modules",
        "driver_make": "Constant Voltage 24V",
        "unit": "Nos",
        "accessories": "Aluminium mounting patti",
        "variants": {
            "DEFAULT": {"wattage": "NA"}
        }
    },
    "GS-DECORATIVE": {
        "page": 144,
        "name": "GS Decorative Pendant Light",
        "description_template": "Housing - Decorative Round Ball with Surface Mounted Canopy, Opal Glass Diffuser, IP20.",
        "led_make": "LED Bulb",
        "driver_make": "NA",
        "unit": "Nos",
        "accessories": "C Clamp",
        "variants": {
            "DEFAULT": {"wattage": "10W"}
        }
    }
}

# Add aliases to point to catalog items
ALIASES = {
    "GS-DE-FDL-84": ("GS-DE-FDL", "84"),
    "GS-DE-FDL-100": ("GS-DE-FDL", "100"),
    "GS-DE-FDL-64": ("GS-DE-FDL", "64"),
    "GS-DE-FDL-58": ("GS-DE-FDL", "58"),
    "GS-DE-FDL-49": ("GS-DE-FDL", "49"),
    
    "GS-DE-INT-84": ("GS-DE-INT", "84"),
    "GS-DE-INT-100": ("GS-DE-INT", "100"),
    
    "GS-LINEA-1714": ("GS-1714", "DEFAULT"),
    "GS LINEAR 1707 CONCEL": ("GS-1714", "DEFAULT"),
    
    "GS-T4-TRACK": ("GS-T4-TRACK", "DEFAULT"),
    "GS-T4-TRACK-2": ("GS-T4-TRACK", "DEFAULT"),
    "GS-T4-TRACK-1": ("GS-T4-TRACK", "DEFAULT"),
    "GS-T4-TRACK-0": ("GS-T4-TRACK", "DEFAULT"),
    "GS-T4-I JOINTER": ("GS-T4-I", "DEFAULT"),
    "GS-T4-POWER ADAPTER": ("GS-T4-POWER", "DEFAULT"),
    "GS-T4-END CAP": ("GS-T4-END", "DEFAULT"),
    
    "SUSPENSION WIRE": ("SUSPENSION", "DEFAULT"),
    "GS-DECORATIVE-10W": ("GS-DECORATIVE", "DEFAULT"),
    "GS-DECORATIVE-LIGHT": ("GS-DECORATIVE", "DEFAULT"),
    "GS-STRETCH FABRIC": ("GS-STRETCH", "DEFAULT"),
}

def parse_product_code(text):
    """
    Parses a string description/code from BOQ and extracts core product components.
    """
    if not text:
        return {}
    
    text = str(text).strip()
    
    # Try to extract a product code like pattern: GS-XXX-XXX-XXX
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
        
    # Extract Wattage (usually matches digits + W e.g. 30W or 36W or 10W/MTR)
    watt_match = re.search(r'(\d+W(?:/MTR)?)', text, re.IGNORECASE)
    if watt_match:
        wattage = watt_match.group(1).upper()
        
    # Rebuild model and size
    if len(parts) >= 3:
        if parts[0].upper() == "GS":
            if parts[1].upper() in ["DE", "IP", "SF", "LINEA", "DR", "SMART", "T4", "LED"]:
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
        "model": model,
        "size_code": size_code,
        "wattage": wattage,
        "cct": cct
    }

def lookup_catalog_database(parsed_info, catalog_json_path=None):
    """
    Looks up specs in the catalog database.
    """
    model = parsed_info.get("model", "")
    size = parsed_info.get("size_code", "")
    wattage = parsed_info.get("wattage", "")
    cct = parsed_info.get("cct", "4000K")
    raw_code = parsed_info.get("raw_code", "")
    
    # Try alias first
    alias_key = f"{model}-{size}" if size else model
    # Clean up alias key to check loose aliases (like GS-T4-TRACK-2 MTR -> GS-T4-TRACK)
    for k_alias in list(ALIASES.keys()):
        if alias_key.startswith(k_alias) or raw_code.upper().startswith(k_alias):
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
            if ph in variant_info:
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
            "accessories": specs["accessories"],
            "matched_by": "exact_catalog_specs"
        }
        
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
