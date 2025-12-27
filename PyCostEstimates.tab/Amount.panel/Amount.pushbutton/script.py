# -*- coding: utf-8 -*-
from pyrevit import revit, DB
from pyrevit import script

output = script.get_output()
doc = revit.doc

# Constants
PARAM_COST = "Cost"
PARAM_TARGET = "Test_1234"
FT3_TO_M3 = 0.0283168
FT2_TO_M2 = 0.092903
FT_TO_M = 0.3048

# Exact material names for structural columns
CONCRETE_NAME = "Concrete - Cast-in-Place Concrete"
STEEL_NAME = "Metal - Steel 43-275"

# Method of cost calculation by category
category_methods = {
    DB.BuiltInCategory.OST_Doors: "count",
    DB.BuiltInCategory.OST_Windows: "count",
    DB.BuiltInCategory.OST_StructuralFraming: "length",
    DB.BuiltInCategory.OST_StructuralFoundation: "volume",
    DB.BuiltInCategory.OST_Floors: "volume",
    DB.BuiltInCategory.OST_Walls: "area",
    DB.BuiltInCategory.OST_Roofs: "area",
    DB.BuiltInCategory.OST_Ceilings: "area",
    DB.BuiltInCategory.OST_Conduit: "length",
    DB.BuiltInCategory.OST_LightingFixtures: "count",
    DB.BuiltInCategory.OST_LightingDevices: "count",
    DB.BuiltInCategory.OST_ElectricalFixtures: "count",
    DB.BuiltInCategory.OST_ElectricalEquipment: "count",
    DB.BuiltInCategory.OST_GenericModel: "area",
    DB.BuiltInCategory.OST_Rebar: "length",
    DB.BuiltInCategory.OST_PlumbingFixtures: "count",
    DB.BuiltInCategory.OST_PipeCurves: "length",
    DB.BuiltInCategory.OST_PipeFitting: "count",
    DB.BuiltInCategory.OST_PipeAccessory: "count",  # ✅ NEW addition
}

# Collect all elements by category
elements = []
for cat in list(category_methods.keys()) + [DB.BuiltInCategory.OST_StructuralColumns]:
    elements += DB.FilteredElementCollector(doc) \
                 .OfCategory(cat) \
                 .WhereElementIsNotElementType() \
                 .ToElements()

# Begin transaction
t = DB.Transaction(doc, "Set Test_1234 using specific material logic including pipe accessories")
t.Start()

updated = 0
skipped = []

for elem in elements:
    try:
        category = elem.Category
        if not category:
            raise Exception("Missing category")

        # Structural Columns: check material to decide method
        if category.Id.IntegerValue == int(DB.BuiltInCategory.OST_StructuralColumns):
            mat_param = elem.LookupParameter("Structural Material")
            if not mat_param:
                raise Exception("No 'Structural Material' parameter")
            mat_elem = doc.GetElement(mat_param.AsElementId())
            mat_name = mat_elem.Name if mat_elem else ""

            if mat_name == CONCRETE_NAME:
                method = "volume"
            elif mat_name == STEEL_NAME:
                method = "length"
            else:
                raise Exception("Unsupported material: '{}'".format(mat_name))
        else:
            method = category_methods.get(DB.BuiltInCategory(category.Id.IntegerValue))
            if not method:
                raise Exception("Unrecognized category")

        # Retrieve parameters
        type_elem = doc.GetElement(elem.GetTypeId())
        cost_param = type_elem.LookupParameter(PARAM_COST)
        target_param = elem.LookupParameter(PARAM_TARGET)

        if not cost_param or not target_param or target_param.IsReadOnly:
            raise Exception("Missing or read-only parameter")

        cost_val = cost_param.AsDouble()
        factor = 1.0  # default for 'count'

        if method == "volume":
            vol_param = elem.LookupParameter("Volume")
            if vol_param and vol_param.HasValue:
                factor = vol_param.AsDouble() * FT3_TO_M3
            else:
                raise Exception("No volume data")

        elif method == "area":
            area_param = elem.LookupParameter("Area")
            if area_param and area_param.HasValue:
                factor = area_param.AsDouble() * FT2_TO_M2
            else:
                raise Exception("No area data")

        elif method == "length":
            if category.Id.IntegerValue == int(DB.BuiltInCategory.OST_Rebar):
                len_param = elem.LookupParameter("Total Bar Length")
            else:
                len_param = elem.LookupParameter("Length")

            if len_param and len_param.HasValue:
                factor = len_param.AsDouble() * FT_TO_M
            else:
                raise Exception("No length data")

        # Calculate and apply amount
        result = cost_val * factor
        target_param.Set(result)
        updated += 1

    except Exception as e:
        skipped.append((elem.Id, str(e)))

t.Commit()

# Output summary
output.print_md("✅ Updated {} element(s) with '{}' = Cost × Quantity.".format(updated, PARAM_TARGET))
if skipped:
    output.print_md("⚠️ Skipped {} element(s):".format(len(skipped)))
    for item in skipped:
        output.print_md("- Element ID {} | Reason: {}".format(item[0], item[1]))
