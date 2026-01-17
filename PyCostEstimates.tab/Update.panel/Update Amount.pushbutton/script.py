# -*- coding: utf-8 -*-
from pyrevit import revit, DB
from pyrevit import script

output = script.get_output()
doc = revit.doc

# ---------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------
PARAM_COST = "Cost"
PARAM_TARGET = "Amount (Qty*Rate)"   # ✅ FIX: exact name including space

FT3_TO_M3 = 0.0283168
FT2_TO_M2 = 0.092903
FT_TO_M = 0.3048

# Exact material names for structural columns
CONCRETE_NAME = "Concrete - Cast-in-Place Concrete"
STEEL_NAME = "Metal - Steel 43-275"

# ---------------------------------------------------------------------
# Method of cost calculation by category
# ---------------------------------------------------------------------
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
    DB.BuiltInCategory.OST_PipeAccessory: "count",
}

# ---------------------------------------------------------------------
# Collect all elements
# ---------------------------------------------------------------------
elements = []
for cat in list(category_methods.keys()) + [DB.BuiltInCategory.OST_StructuralColumns]:
    elements.extend(
        DB.FilteredElementCollector(doc)
        .OfCategory(cat)
        .WhereElementIsNotElementType()
        .ToElements()
    )

# ---------------------------------------------------------------------
# Transaction
# ---------------------------------------------------------------------
t = DB.Transaction(
    doc,
    "Compute Amount (Qty × Rate) using category-specific logic"
)
t.Start()

updated = 0
skipped = []

for elem in elements:
    try:
        category = elem.Category
        if not category:
            raise Exception("Missing category")

        # Structural Columns: decide method by material
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
                raise Exception("Unsupported material: {}".format(mat_name))
        else:
            method = category_methods.get(
                DB.BuiltInCategory(category.Id.IntegerValue)
            )
            if not method:
                raise Exception("Unrecognized category")

        # Retrieve parameters
        type_elem = doc.GetElement(elem.GetTypeId())
        cost_param = type_elem.LookupParameter(PARAM_COST)
        target_param = elem.LookupParameter(PARAM_TARGET)

        if not cost_param:
            raise Exception("Missing 'Cost' type parameter")

        if not target_param:
            raise Exception(
                "Missing instance parameter '{}'".format(PARAM_TARGET)
            )

        if target_param.IsReadOnly:
            raise Exception("'{}' is read-only".format(PARAM_TARGET))

        cost_val = cost_param.AsDouble()
        factor = 1.0  # default for count-based items

        # Quantity extraction
        if method == "volume":
            p = elem.LookupParameter("Volume")
            if not p or not p.HasValue:
                raise Exception("No volume data")
            factor = p.AsDouble() * FT3_TO_M3

        elif method == "area":
            p = elem.LookupParameter("Area")
            if not p or not p.HasValue:
                raise Exception("No area data")
            factor = p.AsDouble() * FT2_TO_M2

        elif method == "length":
            if category.Id.IntegerValue == int(DB.BuiltInCategory.OST_Rebar):
                p = elem.LookupParameter("Total Bar Length")
            else:
                p = elem.LookupParameter("Length")

            if not p or not p.HasValue:
                raise Exception("No length data")
            factor = p.AsDouble() * FT_TO_M

        # Calculate and write amount
        result = cost_val * factor
        target_param.Set(result)
        updated += 1

    except Exception as e:
        skipped.append((elem.Id.IntegerValue, str(e)))

t.Commit()

# ---------------------------------------------------------------------
# Output summary
# ---------------------------------------------------------------------
output.print_md(
    "✅ Updated **{}** element(s) with **{} = Quantity × Rate**."
    .format(updated, PARAM_TARGET)
)

if skipped:
    output.print_md("⚠️ Skipped **{}** element(s):".format(len(skipped)))
    for eid, reason in skipped:
        output.print_md("- Element ID {} | Reason: {}".format(eid, reason))
