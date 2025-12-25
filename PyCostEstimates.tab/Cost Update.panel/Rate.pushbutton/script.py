# -*- coding: utf-8 -*-
import os
import csv
import traceback
from pyrevit import revit, DB, forms

doc = revit.doc

# ---------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------
script_dir = os.path.dirname(__file__)
csv_folder = os.path.join(script_dir, "material_costs")
recipes_csv = os.path.join(script_dir, "recipes.csv")

# ---------------------------------------------------------------------
# Load material prices
# ---------------------------------------------------------------------
material_prices = {}
loaded_files = []

for fname in os.listdir(csv_folder):
    if not fname.endswith(".csv"):
        continue
    with open(os.path.join(csv_folder, fname), "r") as f:
        for row in csv.DictReader(f):
            try:
                material_prices[row["Item"].strip()] = float(row["UnitCost"])
            except:
                pass
    loaded_files.append(fname)

# ---------------------------------------------------------------------
# Load recipes (materials + labour %)
# ---------------------------------------------------------------------
recipes = {}

with open(recipes_csv, "r") as f:
    for row in csv.DictReader(f):
        try:
            rtype = row["Type"].strip()
            comp = row["Component"].strip()
            qty = float(row["Quantity"]) if row["Quantity"] else 0.0
            labour_val = row.get("Labour", "").strip()

            recipes.setdefault(rtype, {
                "materials": {},
                "labour_percent": 0.0
            })

            # Labour row
            if labour_val:
                labour_val = labour_val.replace("%", "")
                recipes[rtype]["labour_percent"] = float(labour_val) / 100.0
            else:
                recipes[rtype]["materials"][comp] = qty

        except:
            pass

# ---------------------------------------------------------------------
# SAFE CATEGORY COLLECTION
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
]

type_elements = []
for cat in CATEGORIES:
    try:
        elems = (
            DB.FilteredElementCollector(doc)
            .OfCategory(cat)
            .WhereElementIsElementType()
            .ToElements()
        )
        type_elements.extend(elems)
    except:
        continue

materials = list(DB.FilteredElementCollector(doc).OfClass(DB.Material))

# ---------------------------------------------------------------------
# Book-keeping
# ---------------------------------------------------------------------
updated = {}
skipped = {}
missing_materials = set()
paint_updated = {}
labour_applied = {}   # <<< NEW (tracking only)

# ---------------------------------------------------------------------
# TRANSACTION
# ---------------------------------------------------------------------
try:
    with revit.Transaction("Update Composite & Paint Costs (With Labour)"):

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

            material_total = 0.0
            labour_percent = recipes[tname]["labour_percent"]
            valid = True

            for mat, qty in recipes[tname]["materials"].items():
                if mat not in material_prices:
                    missing_materials.add(mat)
                    skipped[tname] = "missing material: {}".format(mat)
                    valid = False
                    break
                material_total += qty * material_prices[mat]

            if not valid:
                continue

            labour_cost = material_total * labour_percent
            total_cost = material_total + labour_cost

            cost_param.Set(total_cost)
            updated[tname] = total_cost
            labour_applied[tname] = labour_percent > 0   # <<< NEW

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
# SUMMARY (Labour indicated)
# ---------------------------------------------------------------------
summary = []

if updated:
    summary.append("UPDATED TYPE COSTS (INCL. LABOUR):")
    for name in sorted(updated):
        label = " > Labour Inc." if labour_applied.get(name) else ""
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
