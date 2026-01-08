# -*- coding: utf-8 -*-

"""
Material List Tool - Stages 1 to 4 (STABLE, FINAL)

Stage 1: Extract model quantities from Revit
Stage 2: Match recipes.csv (Type -> Component)
Stage 3: Resolve unit costs from material_unit_costs.csv (Item)
Stage 4: Calculate final quantities and total costs
"""

# ------------------------------------------------------------
# PYREVIT OUTPUT + SAFE UI
# ------------------------------------------------------------

from pyrevit import script, forms
output = script.get_output()
output.print_md("Material List script started")

# ------------------------------------------------------------
# SINGLE SAFE USER INPUT (ONE DIALOG ONLY)
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
# FILE PATHS
# ------------------------------------------------------------

BASE_DIR = os.path.dirname(__file__)

RECIPES_CSV = os.path.abspath(os.path.join(
    BASE_DIR, "..", "..",
    "Rate.panel", "Rate.pushbutton", "recipes.csv"
))

UNIT_COSTS_CSV = os.path.abspath(os.path.join(
    BASE_DIR, "..", "..",
    "Rate.panel", "Rate.pushbutton",
    "material_costs", "material_unit_costs.csv"
))

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
    return (
        text.lower()
        .replace(" ", "")
        .replace("-", "")
        .replace("_", "")
    )

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
# STAGE 1 — EXTRACT MODEL DATA
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

# Convert Revit internal units
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
# STAGE 2 — MATCH RECIPES (CONTAINS MATCH)
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

for family_name, family_data in model_data.items():
    fam_key = normalize(family_name)
    for recipe_type, comps in recipes.items():
        if recipe_type in fam_key:
            for comp, qty in comps:
                family_data["components"][comp] = {"recipe_qty": qty}

output.print_md("Stage 2 complete")

# ------------------------------------------------------------
# STAGE 3 — RESOLVE UNIT COSTS (Item COLUMN)
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

        key = norm_key(name)
        costs[key] = {
            "uom": r["UoM"],
            "unit_cost": float(r[cost_column])
        }
    except:
        pass

priced = 0

for family_data in model_data.values():
    for comp_name, comp in family_data["components"].items():
        lookup = norm_key(comp_name)
        if lookup in costs:
            comp["uom"] = costs[lookup]["uom"]
            comp["unit_cost"] = costs[lookup]["unit_cost"]
            priced += 1

output.print_md("Stage 3 complete")
output.print_md("Priced components: {}".format(priced))

# ------------------------------------------------------------
# STAGE 4 — FINAL QUANTITY & COST AGGREGATION
# ------------------------------------------------------------

output.print_md("Stage 4: Calculating final quantities and totals")

final_materials = {}

for family_name, family_data in model_data.items():
    revit_qty = family_data.get("revit_quantity", 0.0)
    if revit_qty <= 0:
        continue

    for comp_name, comp in family_data["components"].items():
        recipe_qty = comp.get("recipe_qty", 0.0)
        unit_cost = comp.get("unit_cost", 0.0)
        uom = comp.get("uom", "")

        final_qty = revit_qty * recipe_qty

        if comp_name not in final_materials:
            final_materials[comp_name] = {
                "uom": uom,
                "total_qty": 0.0,
                "unit_cost": unit_cost,
                "total_cost": 0.0
            }

        final_materials[comp_name]["total_qty"] += final_qty
        final_materials[comp_name]["total_cost"] += final_qty * unit_cost

output.print_md("Stage 4 complete")
output.print_md("Total unique materials: {}".format(len(final_materials)))

# ------------------------------------------------------------
# SAFE PREVIEW (FIRST 5 MATERIALS)
# ------------------------------------------------------------

output.print_md("Sample results:")

for i, (mat, data) in enumerate(final_materials.items()):
    output.print_md(
        "- {} | {} | Qty: {:.3f} | Unit: {:.2f} | Total: {:.2f}".format(
            mat,
            data["uom"],
            data["total_qty"],
            data["unit_cost"],
            data["total_cost"]
        )
    )
    if i == 4:
        break
