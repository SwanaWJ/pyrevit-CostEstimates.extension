# -*- coding: utf-8 -*-
"""
Material List Generator
- Reads recipes.csv
- Matches BOQ items with recipes (normalized)
- Multiplies material quantities by BOQ family quantities
- Aggregates materials across BOQ items
- Exports REAL .xlsx using Excel COM automation

STABLE version for pyRevit / IronPython
"""

import csv
import os
from collections import defaultdict


# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------

BASE_DIR = os.path.dirname(__file__)

RECIPES_CSV = os.path.abspath(os.path.join(
    BASE_DIR,
    "..",
    "..",
    "Rate.panel",
    "Rate.pushbutton",
    "recipes.csv"
))

DESKTOP = os.path.join(os.environ["USERPROFILE"], "Desktop")
OUTPUT_XLSX = os.path.join(DESKTOP, "Material_List.xlsx")


# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------

def normalize(text):
    return text.lower().replace("_", " ").strip()


# ------------------------------------------------------------
# BOQ ITEMS (PLACEHOLDER)
# ------------------------------------------------------------

def get_boq_items():
    """
    Replace later with real Revit BOQ extraction
    """
    return [
        {"boq_item": "Foundation walls_200mm", "family_qty": 10},
        {"boq_item": "Concrete_slab_100mm", "family_qty": 5},
        {"boq_item": "Pad footing 1200x1200x300mm thick", "family_qty": 3},
    ]


# ------------------------------------------------------------
# READ RECIPES.CSV (IRONPYTHON SAFE)
# ------------------------------------------------------------

def read_recipes(csv_path):
    recipes = defaultdict(list)

    with open(csv_path, "rb") as f:
        raw = f.read()

    # Remove NULL bytes
    raw = raw.replace(b"\x00", b"")

    try:
        text = raw.decode("utf-8")
    except:
        text = raw.decode("latin-1")

    reader = csv.DictReader(text.splitlines())

    for row in reader:
        boq_item = normalize(row.get("Type", ""))
        component = row.get("Component", "").strip()
        qty_raw = row.get("Quantity", "").strip()

        if not boq_item or not component or not qty_raw:
            continue

        skip_words = [
            "labour", "transport", "profit",
            "wastage", "plant", "overhead", "hours"
        ]

        if any(w in component.lower() for w in skip_words):
            continue

        if "%" in qty_raw:
            continue

        try:
            qty = float(qty_raw)
        except:
            continue

        recipes[boq_item].append({
            "material": component,
            "qty_per_family": qty
        })

    return recipes


# ------------------------------------------------------------
# AGGREGATE MATERIALS
# ------------------------------------------------------------

def generate_material_list(boq_items, recipes):
    totals = defaultdict(float)

    print("---- MATCH CHECK ----")

    for item in boq_items:
        name_raw = item["boq_item"]
        name = normalize(name_raw)
        family_qty = item["family_qty"]

        if name not in recipes:
            print("WARNING: No recipe for BOQ item ->", name_raw)
            continue

        for mat in recipes[name]:
            totals[mat["material"]] += mat["qty_per_family"] * family_qty

    return totals


# ------------------------------------------------------------
# EXPORT TO REAL XLSX (PYREVIT-SAFE)
# ------------------------------------------------------------

def export_to_xlsx(materials, output_path):
    """
    SAFE Excel COM automation for pyRevit / IronPython
    """

    import clr
    clr.AddReference("Microsoft.Office.Interop.Excel")
    from Microsoft.Office.Interop import Excel

    excel = None
    workbook = None

    try:
        excel = Excel.ApplicationClass()
        excel.Visible = False          # MUST stay False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Add()
        sheet = workbook.Worksheets[1]
        sheet.Name = "Material List"

        # Headers
        sheet.Cells[1, 1].Value2 = "Material"
        sheet.Cells[1, 2].Value2 = "Total Quantity"
        sheet.Range("A1:B1").Font.Bold = True

        row = 2
        for material, qty in sorted(materials.items()):
            sheet.Cells[row, 1].Value2 = material
            sheet.Cells[row, 2].Value2 = round(qty, 3)
            row += 1

        sheet.Columns("A:B").AutoFit()

        # 51 = xlOpenXMLWorkbook (.xlsx)
        workbook.SaveAs(output_path, 51)

    finally:
        # CRITICAL: close cleanly, no Marshal calls
        if workbook:
            workbook.Close(False)
        if excel:
            excel.Quit()


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

def main():
    if not os.path.exists(RECIPES_CSV):
        raise Exception("recipes.csv not found:\n{}".format(RECIPES_CSV))

    boq_items = get_boq_items()
    recipes = read_recipes(RECIPES_CSV)

    print("Loaded recipe types:", len(recipes))

    materials = generate_material_list(boq_items, recipes)

    if not materials:
        print("WARNING: No materials generated. Check BOQ ↔ recipe names.")

    export_to_xlsx(materials, OUTPUT_XLSX)

    print("Material list generated successfully:")
    print(OUTPUT_XLSX)


if __name__ == "__main__":
    main()
