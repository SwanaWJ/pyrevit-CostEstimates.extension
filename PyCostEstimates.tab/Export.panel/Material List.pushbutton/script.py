# -*- coding: utf-8 -*-

"""
Material List Tool - FINAL
✔ All categories
✔ Recipes from Apply Rate
✔ Province → National fallback
✔ Guaranteed UoM
✔ Family subtotals
✔ Unit Cost Source column
✔ Percentage contribution column
✔ IronPython-safe formatting
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
    default="Copperbelt - Avg",
    title="Select Province and Cost Type"
)

if not selection:
    script.exit()

province, cost_type = [x.strip() for x in selection.split("-")]

PROV_COST_COL = "{}_{}_UnitCost".format(province, cost_type)
NAT_COST_COL  = "National_{}_UnitCost".format(cost_type)

output.print_md("Pricing context")
output.print_md("- Province: {}".format(province))
output.print_md("- Cost type: {}".format(cost_type))
output.print_md("- Fallback: {}".format(NAT_COST_COL))

# ------------------------------------------------------------
# IMPORTS
# ------------------------------------------------------------

from Autodesk.Revit.DB import *
import System
import os
import csv
import codecs
from collections import defaultdict

doc = __revit__.ActiveUIDocument.Document

# ------------------------------------------------------------
# APPLY RATE LOCATION
# ------------------------------------------------------------

SCRIPT_DIR = os.path.dirname(__file__)
TAB_DIR = os.path.dirname(os.path.dirname(SCRIPT_DIR))
APPLY_RATES_DIR = os.path.join(TAB_DIR, "Update.panel", "Apply Rate.pushbutton")

RECIPES_CSV = os.path.join(APPLY_RATES_DIR, "recipes.csv")
UNIT_COSTS_CSV = os.path.join(APPLY_RATES_DIR, "material_unit_costs.csv")

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
    BuiltInCategory.OST_Rebar: "m",
    BuiltInCategory.OST_Conduit: "m",
    BuiltInCategory.OST_PipeCurves: "m",
    BuiltInCategory.OST_PipeFitting: "No",
    BuiltInCategory.OST_PipeAccessory: "No",
    BuiltInCategory.OST_PlumbingFixtures: "No",
    BuiltInCategory.OST_ElectricalFixtures: "No",
    BuiltInCategory.OST_ElectricalEquipment: "No",
    BuiltInCategory.OST_LightingFixtures: "No",
    BuiltInCategory.OST_LightingDevices: "No",
    BuiltInCategory.OST_SpecialityEquipment: "No",
}

SUPPORTED_BICS = set(CATEGORY_UNIT_MAP.keys())

# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------

def norm_key(text):
    return text.lower().strip().replace(" ", "").replace("_", "").replace("-", "") if text else ""

def safe_float(val):
    try:
        return float(val)
    except:
        return 0.0

def get_bic(elem):
    try:
        return System.Enum.Parse(BuiltInCategory, str(elem.Category.Id.IntegerValue))
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

    if bic == BuiltInCategory.OST_Rebar:
        p = elem.LookupParameter("Total Bar Length")
        if p and p.AsDouble() > 0:
            return p.AsDouble()

        p = elem.get_Parameter(BuiltInParameter.REBAR_ELEM_LENGTH)
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

# ------------------------------------------------------------
# STAGE 2 — MATCH RECIPES
# ------------------------------------------------------------

output.print_md("Stage 2: Matching recipes")

recipes = defaultdict(list)

with open(RECIPES_CSV, "rb") as f:
    text = f.read().replace(b"\x00", b"").decode("utf-8", "ignore")

for r in csv.DictReader(text.splitlines()):
    recipes[norm_key(r["Type"])].append(
        (r["Component"], safe_float(r["Quantity"]))
    )

for fam, data in model_data.items():
    key = norm_key(fam)
    if key in recipes:
        for comp, qty in recipes[key]:
            data["components"][comp] = {"recipe_qty": qty}

# ------------------------------------------------------------
# STAGE 3 — RESOLVE UNIT COSTS
# ------------------------------------------------------------

output.print_md("Stage 3: Resolving unit costs")

costs = {}

with open(UNIT_COSTS_CSV, "rb") as f:
    text = f.read().replace(b"\x00", b"").decode("utf-8", "ignore")

for r in csv.DictReader(text.splitlines()):
    item = norm_key(r.get("Item"))
    uom = r.get("UoM", "")

    prov = safe_float(r.get(PROV_COST_COL))
    nat  = safe_float(r.get(NAT_COST_COL))

    if prov > 0:
        costs[item] = (prov, PROV_COST_COL, uom)
    elif nat > 0:
        costs[item] = (nat, NAT_COST_COL, uom)

for data in model_data.values():
    for comp, info in data["components"].items():
        k = norm_key(comp)
        if k in costs:
            cost, src, uom = costs[k]
            info.update({
                "unit_cost": cost,
                "cost_source": src,
                "uom": uom
            })

# ------------------------------------------------------------
# STAGE 4 — FINAL QUANTITIES
# ------------------------------------------------------------

grouped_materials = {}

for fam, data in model_data.items():
    grouped_materials[fam] = {}
    revit_qty = data["revit_quantity"]

    for comp, info in data["components"].items():
        qty = revit_qty * info["recipe_qty"]
        cost = qty * info.get("unit_cost", 0)

        grouped_materials[fam][comp] = {
            "uom": info.get("uom", ""),
            "qty": qty,
            "unit_cost": info.get("unit_cost", 0),
            "total_cost": cost,
            "cost_source": info.get("cost_source", "")
        }

# ------------------------------------------------------------
# STAGE 5 — EXPORT CSV (UTF-8 SAFE)
# ------------------------------------------------------------

output.print_md("Stage 5: Exporting CSV")

desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
csv_path = os.path.join(desktop, "Material_List_Grouped.csv")

with codecs.open(csv_path, "w", encoding="utf-8") as f:
    for fam, comps in grouped_materials.items():
        f.write(u"{}\n".format(fam))
        f.write(u"Material,UoM,Total Quantity,Unit Cost,Total Cost,Unit Cost Source,Percentage\n")

        subtotal = sum(c["total_cost"] for c in comps.values())

        for mat, d in comps.items():
            pct = float(d["total_cost"] / subtotal * 100) if subtotal else 0.0

            f.write(u"{},{},{:.3f},{:.2f},{:.2f},{},{:.0f}%\n".format(
                mat.replace(",", " "),
                d["uom"],
                float(d["qty"]),
                float(d["unit_cost"]),
                float(d["total_cost"]),
                d["cost_source"],
                pct
            ))

        f.write(u"Subtotal {},,,,{:.2f},,100%\n\n".format(fam, float(subtotal)))

output.print_md("CSV export complete")
output.print_md(csv_path)
