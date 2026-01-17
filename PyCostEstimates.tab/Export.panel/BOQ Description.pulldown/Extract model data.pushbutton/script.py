# -*- coding: utf-8 -*-

import csv
import os
import codecs

from Autodesk.Revit.DB import *
from pyrevit import revit

doc = revit.doc

# ---------------------------------------
# COST-RELEVANT CATEGORY WHITELIST
# ---------------------------------------
ALLOWED_CATEGORIES = [
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_Roofs,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_Doors,
    BuiltInCategory.OST_Windows,
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_ElectricalEquipment,
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_GenericModel
]

ALLOWED_CAT_IDS = [int(cat) for cat in ALLOWED_CATEGORIES]

# ---------------------------------------
# OUTPUT → SAME PUSHBUTTON FOLDER
# ---------------------------------------
script_dir = os.path.dirname(__file__)
output_path = os.path.join(script_dir, "FamilyTypes_With_Comments.csv")

# ---------------------------------------
# COLLECT FAMILY TYPES (LOADABLE ONLY)
# ---------------------------------------
collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
rows = []

for symbol in collector:
    try:
        cat = symbol.Category
        if not cat:
            continue

        if cat.Id.IntegerValue not in ALLOWED_CAT_IDS:
            continue

        if hasattr(symbol, "IsActiveViewOnly") and symbol.IsActiveViewOnly:
            continue

        name_param = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if not name_param:
            continue

        type_name = name_param.AsString()
        if not type_name:
            continue

        comment_param = symbol.LookupParameter("Type Comments")
        type_comments = comment_param.AsString() if comment_param else ""

        rows.append([type_name, type_comments or ""])

    except Exception:
        continue

# ---------------------------------------
# WRITE UTF-8 CSV (FORCE OVERWRITE)
# ---------------------------------------
with codecs.open(output_path, "w", encoding="utf-8") as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(["Type", "Type Comments"])
    writer.writerows(rows)

print("CSV created successfully:")
print("Folder:", script_dir)
print("File:", output_path)
print("Exported rows:", len(rows))
