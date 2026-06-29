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
        "description_template": "Housing - Extruded Aluminium Profile, white/black finish, length as specified, standard accessories.",
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
        "description_template": "Housing - T4 Track I Jointer, white/black finish, for inline track connectivity.",
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
        "description_template": "Housing - T4 Track Power Adapter / Live End, white/black finish, for track electrical feed.",
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
        "description_template": "Housing - T4 Track End Cap, white/black finish, to cover track rail terminals.",
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
    },
    "GS-T2-TRACK": {
        "page": 137,
        "name": "T2 Track Rail",
        "description_template": "Housing - Extruded Aluminium Profile, white/black finish, length as specified, standard accessories.",
        "led_make": "NA",
        "driver_make": "NA",
        "unit": "Nos",
        "accessories": "Live End & End Cap",
        "variants": {
            "DEFAULT": {"wattage": "NA"}
        }
    },
    "GS-T2-I": {
        "page": 137,
        "name": "T2 Track I Jointer",
        "description_template": "Housing - T2 Track I Jointer, white/black finish, for inline track connectivity.",
        "led_make": "NA",
        "driver_make": "NA",
        "unit": "Nos",
        "accessories": "None",
        "variants": {
            "DEFAULT": {"wattage": "NA"}
        }
    },
    "GS-T2-POWER": {
        "page": 137,
        "name": "T2 Track Power Adapter",
        "description_template": "Housing - T2 Track Power Adapter / Live End, white/black finish, for track electrical feed.",
        "led_make": "NA",
        "driver_make": "NA",
        "unit": "Nos",
        "accessories": "None",
        "variants": {
            "DEFAULT": {"wattage": "NA"}
        }
    },
    "GS-T2-END": {
        "page": 137,
        "name": "T2 Track End Cap",
        "description_template": "Housing - T2 Track End Cap, white/black finish, to cover track rail terminals.",
        "led_make": "NA",
        "driver_make": "NA",
        "unit": "Nos",
        "accessories": "None",
        "variants": {
            "DEFAULT": {"wattage": "NA"}
        }
    },
    "GS-MDL": {
        "page": 81,
        "name": "Movable Downlight",
        "description_template": "Housing - Die-Cast Aluminum, Reflector and Clear Glass, Dimensions - Dia - {D}mm, Height - {H}mm, Cutout - {C}mm, IP20.",
        "led_make": "Bridgelux",
        "driver_make": "Fulham",
        "unit": "Nos",
        "accessories": "Toggle clips",
        "variants": {
            "5W": {"D": 60, "C": 52, "H": 35, "wattage": "5W"},
            "8W": {"D": 70, "C": 64, "H": 58, "wattage": "8W"},
            "12W": {"D": 83, "C": 75, "H": 68, "wattage": "12W"},
            "20W": {"D": 112, "C": 105, "H": 82, "wattage": "20W"},
            "30W": {"D": 119, "C": 111, "H": 100, "wattage": "30W"}
        }
    },
    "GS-GIMBLE": {
        "page": 84,
        "name": "Gimble Downlight",
        "description_template": "Housing - Die-Cast Aluminum, Reflector and Clear Glass, Dimensions - {D}, Height - {H}mm, Cutout - {C}mm, IP20.",
        "led_make": "Bridgelux",
        "driver_make": "Fulham",
        "unit": "Nos",
        "accessories": "Toggle clips",
        "variants": {
            "30W": {"D": "140 x 140mm", "C": "132 x 132", "H": 100, "wattage": "30W"},
            "2X25W": {"D": "280 x 140mm", "C": "260 x 132", "H": 110, "wattage": "2X25W"},
            "2X30W": {"D": "280 x 140mm", "C": "260 x 132", "H": 110, "wattage": "2X30W"}
        }
    },
    "GS-LED-T8": {
        "page": 237,
        "name": "LED Tube T8",
        "description_template": "Housing - PC Tube, Opal Diffuser, Dimensions - L - 1200mm, Dia - 26mm, IP20.",
        "led_make": "SMD LED",
        "driver_make": "Constant Current",
        "unit": "Nos",
        "accessories": "Mounting Clamps",
        "variants": {
            "20W": {"wattage": "20W"},
            "DEFAULT": {"wattage": "20W"}
        }
    },
    "GS-LED-AL-T5": {
        "page": 237,
        "name": "LED Batten Aluminum T5",
        "description_template": "Housing - Aluminum Body, PMMA Opal Diffuser, Dimensions - L - {L}mm, W - {W}mm, H - {H}mm, IP20.",
        "led_make": "SMD LED",
        "driver_make": "Constant Current",
        "unit": "Nos",
        "accessories": "Mounting Clamps",
        "variants": {
            "20W": {"L": "1200", "W": "22", "H": "33", "wattage": "20W"},
            "40W": {"L": "1200", "W": "22", "H": "33", "wattage": "40W"},
            "DEFAULT": {"L": "1200", "W": "22", "H": "33", "wattage": "20W"}
        }
    },
    "GS-MT-TRACK": {
        "page": 317,
        "name": "Magnetic Track Channel",
        "description_template": "Housing - Extruded Aluminium Profile Magnetic Track Channel, Black Finish, width {W}mm, length {L}mm.",
        "led_make": "NA",
        "driver_make": "NA",
        "unit": "Mtr",
        "accessories": "End caps and Suspension Kit",
        "variants": {
            "36-SUS": {"W": "36", "L": "1000", "wattage": "NA"},
            "26-SUS": {"W": "26", "L": "1000", "wattage": "NA"},
            "DEFAULT": {"W": "26", "L": "1000", "wattage": "NA"}
        }
    },
    "GS-MT-END": {
        "page": 317,
        "name": "Magnetic Track End Cap",
        "description_template": "End Cap for Magnetic Track Channel.",
        "led_make": "NA",
        "driver_make": "NA",
        "unit": "Nos",
        "accessories": "None",
        "variants": {
            "DEFAULT": {"wattage": "NA"}
        }
    },
    "GS-MT-POWER": {
        "page": 317,
        "name": "Magnetic Track Power Adapter / Live End",
        "description_template": "Power Feed / Live End for Magnetic Track.",
        "led_make": "NA",
        "driver_make": "NA",
        "unit": "Nos",
        "accessories": "None",
        "variants": {
            "DEFAULT": {"wattage": "NA"}
        }
    },
    "GS-MT-TL": {
        "page": 319,
        "name": "Magnetic Track Light",
        "description_template": "Housing - Aluminum Extruded, Reflector and Clear Glass, Dimensions - {D}mm, IP20.",
        "led_make": "Bridgelux",
        "driver_make": "Fulham",
        "unit": "Nos",
        "accessories": "Magnetic adapter",
        "variants": {
            "01": {"D": "A: 63 x B: 150", "wattage": "18W"},
            "02": {"D": "A: 63 x B: 93", "wattage": "12W"},
            "04": {"D": "A: 40 x B: 110", "wattage": "8W"},
            "05": {"D": "A: 40 x B: 120", "wattage": "2X5W"},
            "06": {"D": "A: 35 x B: 74", "wattage": "8W"},
            "07": {"D": "A: 45 x B: 100", "wattage": "12W"},
            "DEFAULT": {"D": "A: 63 x B: 150", "wattage": "18W"}
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
    
    # T2 Track & Accessories
    "GS-TRACK PATTI-T2": ("GS-T2-TRACK", "DEFAULT"),
    "GS-T2-EAN CAP": ("GS-T2-END", "DEFAULT"),
    "GS-T2-POWER ADAPTOR": ("GS-T2-POWER", "DEFAULT"),
    "GS-T2-I JOINTE": ("GS-T2-I", "DEFAULT"),
    
    # Magnetic Track & Accessories
    "GS-LED-AL-T5": ("GS-LED-AL-T5", "DEFAULT"),
    "GS-MT-36-SUS": ("GS-MT-TRACK", "36-SUS"),
    "GS-MT-26-SUS": ("GS-MT-TRACK", "26-SUS"),
    "GS-MT-36-SUS-2653": ("GS-MT-TRACK", "36-SUS"),
    "GS-MT-26-SUS-2653": ("GS-MT-TRACK", "26-SUS"),
    "GS-MT-END CAP": ("GS-MT-END", "DEFAULT"),
    "GS-MT-END": ("GS-MT-END", "DEFAULT"),
    "GS-MT-LIVE END": ("GS-MT-POWER", "DEFAULT"),
    "GS-MT-POWER ADAPTER": ("GS-MT-POWER", "DEFAULT"),
    "GS-MT-POWER ADAPTOR": ("GS-MT-POWER", "DEFAULT"),
    "GS-MT-TL-01": ("GS-MT-TL", "01"),
    "GS-MT-TL-02": ("GS-MT-TL", "02"),
    "GS-MT-TL-04": ("GS-MT-TL", "04"),
    "GS-MT-TL-05": ("GS-MT-TL", "05"),
    "GS-MT-TL-06": ("GS-MT-TL", "06"),
    "GS-MT-TL-07": ("GS-MT-TL", "07"),
    "GS-MT-TL": ("GS-MT-TL", "DEFAULT"),
    
    # MDL, Gimble, T8 Tube
    "GS-MDL": ("GS-MDL", ""),
    "GS-GIMBLE": ("GS-GIMBLE", ""),
    "GS-LED-T8": ("GS-LED-T8", ""),
    
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

def lookup_catalog_database(parsed_info, catalog_json_path=None):
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
        
    # 2. Primary fallback: raw text page search in catalog_text.json
    best_page = None
    if catalog_json_path and os.path.exists(catalog_json_path):
        try:
            with open(catalog_json_path, "r", encoding="utf-8") as f:
                catalog = json.load(f)
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
        except Exception as e:
            print("Error matching catalog json:", e)

    # 3. Double Verification & Enrichment using the new CSV database
    import csv
    csv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "GLS_SPA_Product_Database.csv")
    csv_match = None
    if os.path.exists(csv_path):
        try:
            query_clean = raw_code.strip().upper().replace(" ", "")
            with open(csv_path, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    code_val = row.get("Product Code", "").strip().upper()
                    code_clean = code_val.replace(" ", "")
                    if query_clean == code_clean or query_clean in code_clean or code_clean in query_clean:
                        csv_match = row
                        break
        except Exception as e:
            print("Error reading CSV for verification:", e)

    # Resolve page from catalog JSON fallback if not found in text search
    if not best_page:
        catalog_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "GLS_SPA_catalog.json")
        if os.path.exists(catalog_path):
            try:
                query_clean = raw_code.strip().upper().replace(" ", "")
                with open(catalog_path, "r", encoding="utf-8") as f:
                    cat_data = json.load(f)
                    for item in cat_data:
                        item_code = item.get("code", "").strip().upper().replace(" ", "")
                        if query_clean == item_code or query_clean in item_code or item_code in query_clean:
                            best_page = item.get("page", 0)
                            break
            except Exception:
                pass

    if csv_match:
        size = csv_match.get("Size", "")
        lens = csv_match.get("Lens", "Reflector and Clear Glass")
        housing = csv_match.get("Housing", "Die-Cast Aluminum")
        ip = csv_match.get("IP Rating", "IP20")
        beam = csv_match.get("Beam Angle", "")
        driver_make = csv_match.get("Driver Make", "Fulham")
        driver_watt = csv_match.get("Driver Wattage", wattage or "10W")
        
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
            
        desc = ", ".join(parts) + "."
        
        return {
            "page": int(best_page) if best_page else 0,
            "gls_code": csv_match.get("Product Code", raw_code),
            "product_description": desc,
            "led_make": "Bridgelux",
            "driver_make": driver_make,
            "driver_wattage": driver_watt,
            "unit": "Mtr" if "mtr" in original_text or "linear" in original_text else "Nos",
            "accessories": "Standard",
            "matched_by": f"csv_double_verified_page_{best_page}" if best_page else "csv_double_verified"
        }

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
