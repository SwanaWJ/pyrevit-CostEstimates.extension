# -*- coding: utf-8 -*-
import os
import csv
import traceback
from pyrevit import revit, DB, forms

# --- Paths ------------------------------------------------------------
script_dir = os.path.dirname(__file__)
csv_folder_path = os.path.join(script_dir, "material_costs")
csv_recipes_path = os.path.join(script_dir, "recipes.csv")

# --- Load Material Unit-Costs ----------------------------------------
material_prices, loaded_files = {}, []
if os.path.isdir(csv_folder_path):
    csv_files = [f for f in os.listdir(csv_folder_path) if f.endswith(".csv")]
    if not csv_files:
        forms.alert("No CSV files found in 'material_costs' folder.", title="Missing Data")
    else:
        for fname in csv_files:
            try:
                with open(os.path.join(csv_folder_path, fname), "r") as f:
                    for row in csv.DictReader(f):
                        try:
                            material_prices[row["Item"].strip()] = float(row["UnitCost"])
                        except (KeyError, ValueError):
                            continue
                loaded_files.append(fname)
            except Exception as e:
                forms.alert("Error reading '{}': {}".format(fname, e), title="CSV Read Error")
else:
    forms.alert("Folder 'material_costs' not found next to the script.", title="Missing Folder")

# --- Load Recipes -----------------------------------------------------
recipes = {}
try:
    with open(csv_recipes_path, "r") as f:
        for row in csv.DictReader(f):
            try:
                recipes.setdefault(row["Type"].strip(), {})[row["Component"].strip()] = float(row["Quantity"])
            except (KeyError, ValueError):
                continue
except Exception as e:
    forms.alert("Error reading recipes.csv: {}".format(e), title="Recipes Load Error")

# --- Book-keeping -----------------------------------------------------
updated, skipped = [], []
missing_materials = set()
paint_updated, paint_skipped = [], []

# --- Try Import RebarBarType ------------------------------------------
try:
    rebar_type_class = DB.Structure.RebarBarType
except Exception:
    try:
        from Autodesk.Revit.DB.Structure import RebarBarType as rebar_type_class
    except Exception:
        rebar_type_class = None

# ===================== MAIN TRANSACTION ===============================
try:
    with revit.Transaction("Set Composite & Paint Costs from CSV"):

        def apply_cost_to_elements(collected, enum_value, name_param=True):
            for elem in collected:
                if not elem.Category or elem.Category.Id.IntegerValue != int(enum_value):
                    continue
                p = elem.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM) if name_param                     else elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
                if not p:
                    continue

                tname = p.AsString().strip()
                if tname not in recipes:
                    continue

                total_cost, valid = 0, True
                for mat, qty in recipes[tname].items():
                    if mat in material_prices:
                        total_cost += qty * material_prices[mat]
                    else:
                        missing_materials.add(mat)
                        skipped.append("{} (missing price for {})".format(tname, mat))
                        valid = False
                        break

                if valid:
                    cost_param = elem.LookupParameter("Cost")
                    if cost_param and cost_param.StorageType == DB.StorageType.Double and not cost_param.IsReadOnly:
                        cost_param.Set(total_cost)
                        updated.append((tname, total_cost))
                    else:
                        skipped.append("{} (no editable 'Cost' parameter)".format(tname))

        # ---------- ELEMENT TYPE COST APPLICATION -------------------------
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.WallType),              DB.BuiltInCategory.OST_Walls)
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.FloorType),             DB.BuiltInCategory.OST_Floors)
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.WallFoundationType),    DB.BuiltInCategory.OST_StructuralFoundation)
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilySymbol),          DB.BuiltInCategory.OST_StructuralFraming)
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilySymbol),          DB.BuiltInCategory.OST_GenericModel)
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.RoofType),              DB.BuiltInCategory.OST_Roofs)
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.CeilingType),           DB.BuiltInCategory.OST_Ceilings)
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilySymbol),          DB.BuiltInCategory.OST_StructuralColumns)
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilySymbol),          DB.BuiltInCategory.OST_Doors)
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilySymbol),          DB.BuiltInCategory.OST_Windows)

        if rebar_type_class:
            apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(rebar_type_class), DB.BuiltInCategory.OST_Rebar, name_param=False)

        # Electrical
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_Conduit).WhereElementIsElementType(),           DB.BuiltInCategory.OST_Conduit)
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_LightingDevices).WhereElementIsElementType(),   DB.BuiltInCategory.OST_LightingDevices)
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_LightingFixtures).WhereElementIsElementType(),  DB.BuiltInCategory.OST_LightingFixtures)
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_ElectricalFixtures).WhereElementIsElementType(),DB.BuiltInCategory.OST_ElectricalFixtures)
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment).WhereElementIsElementType(),DB.BuiltInCategory.OST_ElectricalEquipment)

        # Plumbing
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilySymbol),          DB.BuiltInCategory.OST_PlumbingFixtures)
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_PipeCurves).WhereElementIsElementType(),        DB.BuiltInCategory.OST_PipeCurves)
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_PipeFitting).WhereElementIsElementType(),       DB.BuiltInCategory.OST_PipeFitting)
        apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_PipeAccessory).WhereElementIsElementType(),     DB.BuiltInCategory.OST_PipeAccessory)

        # Paint / Finishes
        for mat in DB.FilteredElementCollector(revit.doc).OfClass(DB.Material).ToElements():
            mat_name = mat.Name.strip()
            if mat_name in material_prices:
                cost_param = mat.LookupParameter("Cost")
                if cost_param and cost_param.StorageType == DB.StorageType.Double and not cost_param.IsReadOnly:
                    cost_param.Set(material_prices[mat_name])
                    paint_updated.append((mat_name, material_prices[mat_name]))
                else:
                    paint_skipped.append(mat_name)

except Exception as e:
    forms.alert("Script crashed with error:\n{}".format(traceback.format_exc()), title="Crash in Transaction")

# ===================== SUMMARY ========================================
summary = []

if updated:
    summary.append("✅ Updated Type Costs:")
    summary.extend(["- {} : {:.2f} ZMW".format(n, c) for n, c in updated])

if paint_updated:
    summary.append("\n🎨 Updated Paint / Finish Materials:")
    summary.extend(["- {} : {:.2f} ZMW/m²".format(n, c) for n, c in paint_updated])

if skipped:
    summary.append("\n⚠️ Skipped Types:")
    summary.extend(["- " + s for s in skipped])

if paint_skipped:
    summary.append("\n⚠️ Skipped Materials (no editable Cost param):")
    summary.extend(["- " + n for n in paint_skipped])

if missing_materials:
    summary.append("\n❗ Missing materials not priced in CSVs:")
    summary.extend(sorted(missing_materials))

if loaded_files:
    summary.append("\n📂 CSVs loaded:")
    summary.extend(["- " + f for f in loaded_files])

if not summary:
    summary = ["No matching types or materials found."]

forms.alert("\n".join(summary), title="Composite & Paint Cost Update")
