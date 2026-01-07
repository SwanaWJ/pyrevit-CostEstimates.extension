# -*- coding: utf-8 -*-
"""
Material List Generator (pyRevit / IronPython SAFE)

- Reads recipes.csv
- Matches BOQ items with recipes
- Aggregates material quantities
- Exports REAL .xlsx via Excel COM
"""

# ------------------------------------------------------------
# FORCE PYREVIT OUTPUT VISIBILITY
# ------------------------------------------------------------

from pyrevit import script
output = script.get_output()
output.print_md("## Material List Script started")

# ------------------------------------------------------------
# IMPORTS
# ------------------------------------------------------------

import csv
import os
from collections import defaultdict

# ------------------------------------------------------------
# CONFIG (SAFE PATHS)
# ------------------------------------------------------------

try:
    BASE_DIR = os.path.dirname(__file__)
except:
    BASE_DIR = os.getcwd()

RECIPES_CSV = os.path.abspath(os.path.join(
    BASE_DIR,
    "..",
    "..",
    "Rate.panel",
    "Rate.pushbutton",
    "recipes.csv"
))

DESKTOP = os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")
OUTPUT_XLSX = os.path.join(DESKTOP, "Material_List.xlsx")

output.print_md("**CSV path:** `{}`".format(RECIPES_CSV))
output.print_md("**Output:** `{}`".format(OUTPUT_XLSX))

# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------

def normalize(text):
    if not text:
        return ""
    return text.lower().replace("_", " ").strip()

# ------------------------------------------------------------
# BOQ ITEMS (PLACEHOLDER)
# ------------------------------------------------------------

def get_boq_items():
    return [
        {"boq_item": "Foundation walls_200mm", "family_qty": 10},
        {"boq_item": "Concrete_slab_100mm", "family_qty": 5},
        {"boq_item": "Pad footing 1200x1200x300mm thick", "family_qty": 3},
    ]

# ------------------------------------------------------------
# READ RECIPES CSV (IRONPYTHON SAFE)
# ------------------------------------------------------------

def read_recipes(csv_path):
    if not os.path.exists(csv_path):
        raise Exception("recipes.csv not found:\n{}".format(csv_path))

    recipes = defaultdict(list)

    with open(csv_path, "rb") as f:
        raw = f.read()

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

        if any(w in component.lower() for w in (
            "labour", "transport", "profit",
            "wastage", "plant", "overhead", "hours"
        )):
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

    output.print_md("**Loaded recipe types:** {}".format(len(recipes)))
    return recipes

# ------------------------------------------------------------
# AGGREGATE MATERIALS
# ------------------------------------------------------------

def generate_material_list(boq_items, recipes):
    totals = defaultdict(float)
    output.print_md("### Matching BOQ → Recipes")

    for item in boq_items:
        name_raw = item["boq_item"]
        name = normalize(name_raw)
        family_qty = item["family_qty"]

        if name not in recipes:
            output.print_md("⚠️ No recipe for `{}`".format(name_raw))
            continue

        for mat in recipes[name]:
            totals[mat["material"]] += mat["qty_per_family"] * family_qty

    return totals

# ------------------------------------------------------------
# EXPORT TO EXCEL (FIXED FOR YOUR OFFICE INSTALL)
# ------------------------------------------------------------

def export_to_xlsx(materials, output_path):
    output.print_md("### Exporting to Excel...")

    import clr
    import System
    clr.AddReference("Microsoft.Office.Interop.Excel")
    clr.AddReference("System.Runtime.InteropServices")

    from Microsoft.Office.Interop.Excel import ApplicationClass
    from System.Runtime.InteropServices import Marshal

    excel = workbook = sheet = None

    try:
        # 🔑 REQUIRED on your machine
        excel = ApplicationClass()
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Add()
        sheet = workbook.Worksheets[1]
        sheet.Name = "Material List"

        # Headers
        sheet.Cells(1, 1).Value2 = "Material"
        sheet.Cells(1, 2).Value2 = "Total Quantity"
        sheet.Range("A1:B1").Font.Bold = True

        row = 2
        for material, qty in sorted(materials.items()):
            sheet.Cells(row, 1).Value2 = material
            sheet.Cells(row, 2).Value2 = round(qty, 3)
            row += 1

        sheet.Columns("A:B").AutoFit()
        workbook.SaveAs(output_path, 51)

        output.print_md("✅ Excel saved successfully")

    finally:
        # 🔴 CLEANUP ORDER IS CRITICAL
        if sheet:
            Marshal.ReleaseComObject(sheet)
            sheet = None

        if workbook:
            workbook.Close(False)
            Marshal.ReleaseComObject(workbook)
            workbook = None

        if excel:
            excel.Quit()
            Marshal.ReleaseComObject(excel)
            excel = None

        # Prevent Revit freeze
        System.GC.Collect()
        System.GC.WaitForPendingFinalizers()
        System.GC.Collect()

# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

def main():
    boq_items = get_boq_items()
    recipes = read_recipes(RECIPES_CSV)
    materials = generate_material_list(boq_items, recipes)

    if not materials:
        output.print_md("❌ No materials generated")
        return

    export_to_xlsx(materials, OUTPUT_XLSX)
    output.print_md("## 🎉 DONE")

# ------------------------------------------------------------
# ENTRY POINT
# ------------------------------------------------------------

try:
    main()
except Exception as e:
    output.print_md("## ❌ SCRIPT FAILED")
    output.print_md("```\n{}\n```".format(str(e)))
    raise
