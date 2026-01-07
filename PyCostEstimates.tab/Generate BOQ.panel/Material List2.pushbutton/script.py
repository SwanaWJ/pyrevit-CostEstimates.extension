# -*- coding: utf-8 -*-
"""
Material List - Stage 2

- Stage 1: Extract all families from model
- Stage 2: Match recipes.csv and attach components
"""

from pyrevit import script
from Autodesk.Revit.DB import *
import System
import os
import csv
from collections import defaultdict

output = script.get_output()
doc = __revit__.ActiveUIDocument.Document

output.print_md("## Material List - Stage 2")
output.print_md("Extracting model quantities...")

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

# ------------------------------------------------------------
# CATEGORY TO UNIT MAP
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
# STAGE 1 - EXTRACT MODEL DATA
# ------------------------------------------------------------

model_data = {}

elements = (
    FilteredElementCollector(doc)
    .WhereElementIsNotElementType()
    .ToElements()
)

for elem in elements:
    if not elem.Category:
        continue

    bic = get_bic(elem)
    if bic not in SUPPORTED_BICS:
        continue

    type_id = elem.GetTypeId()
    if type_id == ElementId.InvalidElementId:
        continue

    elem_type = doc.GetElement(type_id)
    if not elem_type:
        continue

    p = elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    type_name = p.AsString() if p else "UNKNOWN"

    unit = CATEGORY_UNIT_MAP[bic]
    raw_qty = get_raw_quantity(elem, bic)

    if type_name not in model_data:
        model_data[type_name] = {
            "category": bic,
            "unit": unit,
            "raw_qty": 0.0,
            "components": {}
        }

    model_data[type_name]["raw_qty"] += raw_qty

# Convert quantities
for data in model_data.values():
    if data["unit"] == "m2":
        data["revit_quantity"] = round(
            UnitUtils.ConvertFromInternalUnits(
                data["raw_qty"], UnitTypeId.SquareMeters
            ), 3
        )
    elif data["unit"] == "m3":
        data["revit_quantity"] = round(
            UnitUtils.ConvertFromInternalUnits(
                data["raw_qty"], UnitTypeId.CubicMeters
            ), 3
        )
    elif data["unit"] == "m":
        data["revit_quantity"] = round(
            UnitUtils.ConvertFromInternalUnits(
                data["raw_qty"], UnitTypeId.Meters
            ), 3
        )
    else:
        data["revit_quantity"] = round(data["raw_qty"], 3)

# ------------------------------------------------------------
# STAGE 2 - READ RECIPES CSV
# ------------------------------------------------------------

if not os.path.exists(RECIPES_CSV):
    output.print_md("ERROR: recipes.csv not found")
    script.exit()

recipes = defaultdict(list)

with open(RECIPES_CSV, "rb") as f:
    raw = f.read().replace(b"\x00", b"")

try:
    text = raw.decode("utf-8")
except:
    text = raw.decode("latin-1")

reader = csv.DictReader(text.splitlines())

for row in reader:
    type_name = normalize(row.get("Type"))
    component = row.get("Component", "").strip()
    qty_raw = row.get("Quantity", "").strip()

    if not type_name or not component or not qty_raw:
        continue

    try:
        qty = float(qty_raw)
    except:
        continue

    recipes[type_name].append({
        "component": component,
        "qty": qty
    })

# ------------------------------------------------------------
# ATTACH COMPONENTS TO MODEL DATA
# ------------------------------------------------------------

output.print_md("### Recipe matching results")

for type_name, data in model_data.items():
    key = normalize(type_name)

    if key not in recipes:
        output.print_md("No recipe for: {}".format(type_name))
        continue

    for item in recipes[key]:
        data["components"][item["component"]] = {
            "recipe_qty": item["qty"]
        }

    output.print_md("Matched recipe for: {}".format(type_name))
    for c, v in data["components"].items():
        output.print_md("- {} x {}".format(c, v["recipe_qty"]))

output.print_md("Stage 2 complete")
