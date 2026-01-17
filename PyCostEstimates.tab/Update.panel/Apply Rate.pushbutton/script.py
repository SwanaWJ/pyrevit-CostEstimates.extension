# -*- coding: utf-8 -*-
import os
import csv
import traceback
from pyrevit import revit, DB, forms

doc = revit.doc

# ---------------------------------------------------------------------
# USER INPUTS
# ---------------------------------------------------------------------

province = forms.SelectFromList.show(
    [
        "Central", "Copperbelt", "Eastern", "Luapula", "Lusaka",
        "Muchinga", "Northern", "NorthWestern", "Southern",
        "Western", "National",
    ],
    title="Select Province",
    button_name="Use Selected Province"
)

if not province:
    forms.alert("No province selected. Script cancelled.")
    raise SystemExit

cost_basis = forms.SelectFromList.show(
    ["Min", "Avg", "Max"],
    title="Select Unit Cost Basis",
    button_name="Use Selected Cost"
)

if not cost_basis:
    forms.alert("No unit cost basis selected. Script cancelled.")
    raise SystemExit

cost_column = "{}_{}_UnitCost".format(province, cost_basis)
national_column = "National_{}_UnitCost".format(cost_basis)

# ---------------------------------------------------------------------
# Paths (FIXED: single CSV, no folder)
# ---------------------------------------------------------------------
script_dir = os.path.dirname(__file__)
material_costs_csv = os.path.join(script_dir, "material_unit_costs.csv")
recipes_csv = os.path.join(script_dir, "recipes.csv")

if not os.path.exists(material_costs_csv):
    forms.alert(
        "Material unit cost file not found:\n\n{}".format(material_costs_csv),
        title="Missing Material Cost File"
    )
    raise SystemExit

# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------
def is_valid_cost(value):
    try:
        return value is not None and str(value).strip() != "" and float(value) > 0
    except:
        return False

# ---------------------------------------------------------------------
# Load material prices (Province → National fallback)
# ---------------------------------------------------------------------
material_prices = {}
material_price_source = {}
loaded_files = []

with open(material_costs_csv, "r") as f:
    reader = csv.DictReader(f)
    for row in reader:
        try:
            item = row["Item"].strip()
            if not item:
                continue

            prov_val = row.get(cost_column, "")
            if is_valid_cost(prov_val):
                material_prices[item] = float(prov_val)
                material_price_source[item] = cost_column
                continue

            nat_val = row.get(national_column, "")
            if is_valid_cost(nat_val):
                material_prices[item] = float(nat_val)
                material_price_source[item] = national_column

        except:
            continue

loaded_files.append(os.path.basename(material_costs_csv))

# ---------------------------------------------------------------------
# Load recipes (UNCHANGED)
# ---------------------------------------------------------------------
recipes = {}

with open(recipes_csv, "r") as f:
    for row in csv.DictReader(f):
        try:
            rtype = row["Type"].strip()
            comp = row["Component"].strip()
            qty = float(row["Quantity"]) if row["Quantity"] else 0.0

            pct = row.get("Labour/Transport/Wastage/Profit", "").strip()
            fixed = row.get("Labour/Transport/Plant_Fixed", "").strip()
            time_dist = row.get("Time/Distance", "").strip()
            rate = row.get("Rate", "").strip()

            recipes.setdefault(rtype, {
                "materials": {},
                "labour_percent": 0.0,
                "labour_fixed": [],
                "labour_time": [],
                "transport_percent": 0.0,
                "transport_fixed": [],
                "transport_distance": [],
                "wastage_percent": 0.0,
                "plant_percent": 0.0,
                "plant_fixed": [],
                "plant_time": [],
                "overhead_percent": 0.0,
            })

            cname = comp.lower()

            if pct:
                pct_val = float(pct.replace("%", "")) / 100.0
                if "wastage" in cname or "shrinkage" in cname:
                    recipes[rtype]["wastage_percent"] = pct_val
                elif "profit" in cname or "overhead" in cname:
                    recipes[rtype]["overhead_percent"] = pct_val
                elif cname.startswith("transport"):
                    recipes[rtype]["transport_percent"] = pct_val
                elif "plant" in cname:
                    recipes[rtype]["plant_percent"] = pct_val
                else:
                    recipes[rtype]["labour_percent"] = pct_val

            if fixed:
                if cname.startswith("transport"):
                    recipes[rtype]["transport_fixed"].append(float(fixed))
                elif "plant" in cname:
                    recipes[rtype]["plant_fixed"].append(float(fixed))
                else:
                    recipes[rtype]["labour_fixed"].append(float(fixed))

            if time_dist and rate:
                cost = float(time_dist) * float(rate)
                if cname.startswith("transport"):
                    recipes[rtype]["transport_distance"].append(cost)
                elif "plant" in cname:
                    recipes[rtype]["plant_time"].append(cost)
                else:
                    recipes[rtype]["labour_time"].append(cost)

            if not pct and not fixed and not time_dist:
                recipes[rtype]["materials"][comp] = qty

        except:
            continue

# ---------------------------------------------------------------------
# Categories (UNCHANGED)
# ---------------------------------------------------------------------
CATEGORIES = [
    DB.BuiltInCategory.OST_Walls,
    DB.BuiltInCategory.OST_Floors,
    DB.BuiltInCategory.OST_Roofs,
    DB.BuiltInCategory.OST_Ceilings,
    DB.BuiltInCategory.OST_Doors,
    DB.BuiltInCategory.OST_Windows,
    DB.BuiltInCategory.OST_StructuralColumns,
    DB.BuiltInCategory.OST_StructuralFraming,
    DB.BuiltInCategory.OST_StructuralFoundation,
    DB.BuiltInCategory.OST_Conduit,
    DB.BuiltInCategory.OST_ElectricalFixtures,
    DB.BuiltInCategory.OST_ElectricalEquipment,
    DB.BuiltInCategory.OST_LightingFixtures,
    DB.BuiltInCategory.OST_LightingDevices,
    DB.BuiltInCategory.OST_PlumbingFixtures,
    DB.BuiltInCategory.OST_PipeCurves,
    DB.BuiltInCategory.OST_PipeFitting,
    DB.BuiltInCategory.OST_PipeAccessory,
    DB.BuiltInCategory.OST_GenericModel,
    DB.BuiltInCategory.OST_SpecialityEquipment,
]

type_elements = []
for cat in CATEGORIES:
    try:
        type_elements.extend(
            DB.FilteredElementCollector(doc)
            .OfCategory(cat)
            .WhereElementIsElementType()
            .ToElements()
        )
    except:
        continue

materials = list(DB.FilteredElementCollector(doc).OfClass(DB.Material))

# ---------------------------------------------------------------------
# Book-keeping (UNCHANGED)
# ---------------------------------------------------------------------
updated = {}
skipped = {}
missing_materials = set()
paint_updated = {}

labour_applied = {}
transport_applied = {}
plant_applied = {}
wastage_applied = {}
overhead_applied = {}
national_fallback_used = {}

# ---------------------------------------------------------------------
# TRANSACTION (RESTORED)
# ---------------------------------------------------------------------
try:
    with revit.Transaction(
        "Composite & Paint Cost Update [{}]".format(cost_column)
    ):

        for elem in type_elements:
            cost_param = elem.LookupParameter("Cost")
            if not cost_param or cost_param.IsReadOnly:
                continue

            name_param = elem.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if not name_param:
                continue

            tname = name_param.AsString()
            if not tname or tname not in recipes:
                continue

            r = recipes[tname]
            material_total = 0.0
            valid = True

            for mat, qty in r["materials"].items():
                if mat not in material_prices:
                    missing_materials.add(mat)
                    skipped[tname] = "missing material: {}".format(mat)
                    valid = False
                    break

                material_total += qty * material_prices[mat]

                src = material_price_source.get(mat, "")
                if src.startswith("National"):
                    national_fallback_used[tname] = src.replace("_UnitCost", "")

            if not valid:
                continue

            wastage_cost = material_total * r["wastage_percent"]

            labour_cost = (
                material_total * r["labour_percent"]
                + sum(r["labour_fixed"])
                + sum(r["labour_time"])
            )

            transport_cost = (
                material_total * r["transport_percent"]
                + sum(r["transport_fixed"])
                + sum(r["transport_distance"])
            )

            plant_cost = (
                material_total * r["plant_percent"]
                + sum(r["plant_fixed"])
                + sum(r["plant_time"])
            )

            subtotal = (
                material_total
                + wastage_cost
                + labour_cost
                + transport_cost
                + plant_cost
            )

            overhead_cost = subtotal * r["overhead_percent"]
            total_cost = subtotal + overhead_cost

            cost_param.Set(total_cost)

            updated[tname] = total_cost
            labour_applied[tname] = labour_cost > 0
            transport_applied[tname] = transport_cost > 0
            plant_applied[tname] = plant_cost > 0
            wastage_applied[tname] = wastage_cost > 0
            overhead_applied[tname] = overhead_cost > 0

        # Paint / finishes
        for mat in materials:
            if mat.Name in material_prices:
                p = mat.LookupParameter("Cost")
                if p and not p.IsReadOnly:
                    p.Set(material_prices[mat.Name])
                    paint_updated[mat.Name] = material_prices[mat.Name]

except Exception:
    forms.alert(traceback.format_exc(), title="Cost Update Failed")
    raise

# ---------------------------------------------------------------------
# SUMMARY (UNCHANGED)
# ---------------------------------------------------------------------
summary = []
summary.append("UNIT COST COLUMN USED: {}\n".format(cost_column))

if updated:
    summary.append(
        "UPDATED TYPE COSTS (INCL. LABOUR, TRANSPORT, WASTAGE, PLANT & PROFIT):"
    )
    for name in sorted(updated):
        labels = []
        if labour_applied.get(name):
            labels.append("⚠️ Labour")
        if transport_applied.get(name):
            labels.append("🚚 Transport")
        if plant_applied.get(name):
            labels.append("🚜 Plant")
        if wastage_applied.get(name):
            labels.append("♻️ Wastage")
        if overhead_applied.get(name):
            labels.append("💼 Profit")

        fallback = ""
        if name in national_fallback_used:
            fallback = " ⚠️ [{}]".format(national_fallback_used[name])

        label = "  " + ", ".join(labels) + fallback if labels or fallback else ""
        summary.append("- {} : {:.2f} ZMW{}".format(name, updated[name], label))

if paint_updated:
    summary.append("\nUPDATED PAINT / FINISH MATERIALS:")
    for name in sorted(paint_updated):
        summary.append("- {} : {:.2f} ZMW".format(name, paint_updated[name]))

if skipped:
    summary.append("\nSKIPPED TYPES:")
    for name in sorted(skipped):
        summary.append("- {} ({})".format(name, skipped[name]))

if missing_materials:
    summary.append("\nMISSING MATERIALS:")
    for m in sorted(missing_materials):
        summary.append("- " + str(m))

if loaded_files:
    summary.append("\nCSVs LOADED:")
    for f in sorted(loaded_files):
        summary.append("- " + f)

forms.alert("\n".join(summary), title="Composite & Paint Cost Update")
