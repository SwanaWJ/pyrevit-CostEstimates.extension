# -*- coding: utf-8 -*-
"""
Material List Generator
- Reads recipes.csv
- Matches BOQ items with recipes (normalized)
- Multiplies material quantities by BOQ family quantities
- Aggregates materials across BOQ items
- Exports shopping-ready CSV to Desktop

Compatible with pyRevit / IronPython
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
OUTPUT_CSV = os.path.join(DESKTOP, "Material_List.csv")


# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------

def normalize(text):
    """Normalize names for matching"""
    return text.lower().replace("_", " ").strip()


# ------------------------------------------------------------
# GET BOQ ITEMS (PLACEHOLDER)
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

    # --- Read binary to avoid NULL byte crash ---
    with open(csv_path, "rb") as f:
        raw = f.read()

    # Remove NULL bytes
    raw = raw.replace(b"\x00", b"")

    # Decode safely
    try:
        text = raw.decode("utf-8")
    except:
        text = raw.decode("latin-1")

    lines = text.splitlines()
    reader = csv.DictReader(lines)

    for row in reader:
        boq_item = normalize(row.get("Type", ""))
        component = row.get("Component", "").strip()
        qty_raw = row.get("Quantity", "").strip()

        if not boq_item or not component or not qty_raw:
            continue

        # Skip non-material rows
        skip_words = [
            "labour", "transport", "profit",
            "wastage", "plant", "overhead", "hours"
        ]

        if any(w in component.lower() for w in skip_words):
            continue

        # Skip percentage values
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
    material_totals = defaultdict(float)

    print("---- MATCH CHECK ----")

    for item in boq_items:
        boq_name_raw = item["boq_item"]
        boq_name = normalize(boq_name_raw)
        family_qty = item["family_qty"]

        if boq_name not in recipes:
            print("WARNING: No recipe for BOQ item ->", boq_name_raw)
            continue

        for mat in recipes[boq_name]:
            material_totals[mat["material"]] += (
                mat["qty_per_family"] * family_qty
            )

    return material_totals


# ------------------------------------------------------------
# EXPORT TO CSV (EXCEL SAFE)
# ------------------------------------------------------------

def export_to_csv(materials, output_path):
    with open(output_path, "wb") as f:
        writer = csv.writer(f)
        writer.writerow(["Material", "Total Quantity"])

        for material, qty in sorted(materials.items()):
            writer.writerow([material, round(qty, 3)])


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

    export_to_csv(materials, OUTPUT_CSV)

    print("Material list generated successfully:")
    print(OUTPUT_CSV)


if __name__ == "__main__":
    main()
