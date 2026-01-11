# -*- coding: utf-8 -*-

"""
Material List Tool - FINAL
Stage 1: Extract Revit quantities
Stage 2: Match recipes.csv (Type → Component)
Stage 3: Resolve unit costs (Item)
Stage 4: Calculate quantities GROUPED BY TYPE
Stage 5: Export grouped CSV (QS format)
"""

# ------------------------------------------------------------
# PYREVIT OUTPUT + UI
# ------------------------------------------------------------

from pyrevit import script, forms
output = script.get_output()
output.print_md("Material List script started")

# ------------------------------------------------------------
# USER INPUT
# ------------------------------------------------------------

provinces = [
    "Central", "Copperbelt", "Eastern", "Luapula", "Lusaka",
    "Muchinga", "Northern", "NorthWestern", "Southern",
    "Western", "National"
]

cost_types = ["Min", "Avg", "Max"]
choices = ["{} - {}".format(p, c) for p in provinces for c in cost_types]

selection = forms.ask_for_one_item(
    choices,
    default="Central - Avg",
    title="Select Province and Cost Type"
)

if not selection:
    script.exit()

province, cost_type = [x.strip() for x in selection.split("-")]
cost_column = "{}_{}_UnitCost".format(province, cost_type)

output.print_md("Pricing context selected")
output.print_md("- Province: {}".format(province))
output.print_md("- Cost type: {}".format(cost_type))

# ------------------------------------------------------------
# IMPORTS
# ------------------------------------------------------------

from Autodesk.Revit.DB import *
import System
import os
import csv
from collections import defaultdict

doc = __revit__.ActiveUIDocument.Document

# ------------------------------------------------------------
# FILE PATHS (FIXED)
# ------------------------------------------------------------

BASE_DIR = os.path.dirname(__file__)

RECIPES_CSV = os.path.abspath(os.path.join(
    BASE_DIR, "..", "..",
    "Rate.panel", "Rate.pushbutton",
    "recipes.csv"
))

UNIT_COSTS_CSV = os.path.abspath(os.path.join(
    BASE_DIR, "..", "..",
    "Rate.panel", "Rate.pushbutton",
    "material_unit_costs.csv"
))

if not os.path.exists(UNIT_COSTS_CSV):
    forms.alert(
        "Unit cost file not found:\n\n{}".format(UNIT_COSTS_CSV),
        title="Missing Unit Cost CSV"
    )
    script.exit()

# ------------------------------------------------------------
# CATEGORY → UNIT MAP
# ------------------------------------------------------------

CATEGORY_UNIT_MAP = {
    BuiltInCategory.OST_Walls: "m2",
    BuiltInCategory.OST_Floors: "m3",
    BuiltInCategory.OST_Roofs: "m2",
    BuiltInCategory.OST_Ceilings: "m2",
    BuiltInCategory.OST_Doors: "No",
    BuiltInCategory.OST_Windows: "No",
    BuiltInCategory.OST_StructuralColumns: "m3",
    BuiltInCategory.OST_StructuralFraming: "m",
    BuiltInCategory.OST_StructuralFoundation: "m3",
    BuiltInCategory.OST_Conduit: "m",
    BuiltInCategory.OST_PipeCurves: "m",
    BuiltInCategory.OST_ElectricalFixtures: "No",
    BuiltInCategory.OST_ElectricalEquipment: "No",
    BuiltInCategory.OST_LightingFixtures: "No",
    BuiltInCategory.OST_LightingDevices: "No",
    BuiltInCategory.OST_PlumbingFixtures: "No",
    BuiltInCategory.OST_PipeFitting: "No",
    BuiltInCategory.OST_PipeAccessory: "No",
    BuiltInCategory.OST_SpecialityEquipment: "No",
}

SUPPORTED_BICS = set(CATEGORY_UNIT_MAP.keys())

# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------

def normalize(text):
    return text.lower().strip() if text else ""

def norm_key(text):
    if not text:
        return ""
    return text.lower().replace(" ", "").replace("-", "").replace("_", "")

def get_bic(elem):
    try:
        return System.Enum.Parse(
            BuiltInCategory,
            str(elem.Category.Id.IntegerValue)
        )
    except:
        return None

def get_raw_quantity(elem, bic):
    if bic == BuiltInCategory.OST_Walls:
        p = elem.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED)
        return p.AsDouble() if p else 0.0

    if bic in (
        BuiltInCategory.OST_Floors,
        BuiltInCategory.OST_StructuralFoundation,
        BuiltInCategory.OST_StructuralColumns
    ):
        p = elem.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED)
        return p.AsDouble() if p else 0.0

    if bic in (
        BuiltInCategory.OST_StructuralFraming,
        BuiltInCategory.OST_PipeCurves,
        BuiltInCategory.OST_Conduit
    ):
        p = elem.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
        return p.AsDouble() if p else 0.0

    return 1.0

# ------------------------------------------------------------
# STAGE 1 — EXTRACT MODEL QUANTITIES
# ------------------------------------------------------------

output.print_md("Stage 1: Extracting model quantities")

model_data = {}

elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()

for elem in elements:
    if not elem.Category:
        continue

    bic = get_bic(elem)
    if bic not in SUPPORTED_BICS:
        continue

    elem_type = doc.GetElement(elem.GetTypeId())
    if not elem_type:
        continue

    type_name = elem_type.get_Parameter(
        BuiltInParameter.SYMBOL_NAME_PARAM
    ).AsString()

    unit = CATEGORY_UNIT_MAP[bic]
    raw_qty = get_raw_quantity(elem, bic)

    model_data.setdefault(type_name, {
        "unit": unit,
        "raw_qty": 0.0,
        "revit_quantity": 0.0,
        "components": {}
    })

    model_data[type_name]["raw_qty"] += raw_qty

for d in model_data.values():
    if d["unit"] == "m2":
        d["revit_quantity"] = UnitUtils.ConvertFromInternalUnits(
            d["raw_qty"], UnitTypeId.SquareMeters)
    elif d["unit"] == "m3":
        d["revit_quantity"] = UnitUtils.ConvertFromInternalUnits(
            d["raw_qty"], UnitTypeId.CubicMeters)
    elif d["unit"] == "m":
        d["revit_quantity"] = UnitUtils.ConvertFromInternalUnits(
            d["raw_qty"], UnitTypeId.Meters)
    else:
        d["revit_quantity"] = d["raw_qty"]

output.print_md("Stage 1 complete")

# ------------------------------------------------------------
# STAGE 2 — MATCH RECIPES
# ------------------------------------------------------------

output.print_md("Stage 2: Matching recipes")

recipes = defaultdict(list)

with open(RECIPES_CSV, "rb") as f:
    text = f.read().replace(b"\x00", b"").decode("utf-8", "ignore")

for r in csv.DictReader(text.splitlines()):
    try:
        recipes[normalize(r["Type"])].append(
            (r["Component"].strip(), float(r["Quantity"]))
        )
    except:
        pass

for fam, data in model_data.items():
    fam_key = normalize(fam)
    for recipe_type, comps in recipes.items():
        if recipe_type in fam_key:
            for comp, qty in comps:
                data["components"][comp] = {"recipe_qty": qty}

output.print_md("Stage 2 complete")

# ------------------------------------------------------------
# STAGE 3 — RESOLVE UNIT COSTS (FIXED PATH)
# ------------------------------------------------------------

output.print_md("Stage 3: Resolving unit costs")

costs = {}

with open(UNIT_COSTS_CSV, "rb") as f:
    text = f.read().replace(b"\x00", b"").decode("utf-8", "ignore")

for r in csv.DictReader(text.splitlines()):
    try:
        name = r.get("Item")
        if not name:
            continue
        costs[norm_key(name)] = {
            "uom": r["UoM"],
            "unit_cost": float(r[cost_column])
        }
    except:
        pass

for data in model_data.values():
    for comp, info in data["components"].items():
        key = norm_key(comp)
        if key in costs:
            info.update(costs[key])

output.print_md("Stage 3 complete")

# ------------------------------------------------------------
# STAGE 4 — FINAL QUANTITIES (GROUPED BY TYPE)
# ------------------------------------------------------------

output.print_md("Stage 4: Calculating final quantities (grouped by type)")

grouped_materials = {}

for type_name, data in model_data.items():
    revit_qty = data["revit_quantity"]
    if revit_qty <= 0 or not data["components"]:
        continue

    grouped_materials.setdefault(type_name, {})

    for comp, info in data["components"].items():
        final_qty = revit_qty * info.get("recipe_qty", 0.0)

        grouped_materials[type_name].setdefault(comp, {
            "uom": info.get("uom", ""),
            "total_qty": 0.0,
            "unit_cost": info.get("unit_cost", 0.0),
            "total_cost": 0.0
        })

        grouped_materials[type_name][comp]["total_qty"] += final_qty
        grouped_materials[type_name][comp]["total_cost"] += (
            final_qty * info.get("unit_cost", 0.0)
        )

output.print_md("Stage 4 complete")
output.print_md("Total types: {}".format(len(grouped_materials)))

# ------------------------------------------------------------
# STAGE 5 — EXPORT GROUPED CSV
# ------------------------------------------------------------

output.print_md("Stage 5: Exporting grouped material list to CSV")

desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
csv_path = os.path.join(desktop, "Material_List_Grouped.csv")

with open(csv_path, "wb") as f:
    for type_name, components in sorted(grouped_materials.items()):
        f.write("{}\n".format(type_name))
        f.write("Material,UoM,Total Quantity,Unit Cost,Total Cost\n")

        for material, data in sorted(components.items()):
            line = "{},{},{:.3f},{:.2f},{:.2f}\n".format(
                material.replace(",", " "),
                data["uom"],
                data["total_qty"],
                data["unit_cost"],
                data["total_cost"]
            )
            f.write(line)

        f.write("\n")

output.print_md("CSV export complete")
output.print_md(csv_path)
