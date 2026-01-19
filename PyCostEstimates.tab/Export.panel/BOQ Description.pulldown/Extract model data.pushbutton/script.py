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
ALLOWED_CAT_IDS = {
    int(BuiltInCategory.OST_Walls),
    int(BuiltInCategory.OST_Floors),
    int(BuiltInCategory.OST_Roofs),
    int(BuiltInCategory.OST_StructuralFraming),
    int(BuiltInCategory.OST_StructuralColumns),
    int(BuiltInCategory.OST_StructuralFoundation),
    int(BuiltInCategory.OST_Doors),
    int(BuiltInCategory.OST_Windows),
    int(BuiltInCategory.OST_PlumbingFixtures),
    int(BuiltInCategory.OST_MechanicalEquipment),
    int(BuiltInCategory.OST_ElectricalEquipment),
    int(BuiltInCategory.OST_ElectricalFixtures),
    int(BuiltInCategory.OST_GenericModel),

    # ✅ ADDED (VERTICAL CIRCULATION)
    int(BuiltInCategory.OST_Stairs),
    int(BuiltInCategory.OST_Ramps),
    int(BuiltInCategory.OST_Railings),
}

# ---------------------------------------
# OUTPUT (SINGLE SOURCE OF TRUTH)
# ---------------------------------------
script_dir = os.path.dirname(__file__)
output_path = os.path.join(script_dir, "USED_TYPES_ONLY.csv")

# ---------------------------------------
# STEP 1 — COLLECT USED TYPE IDS
# ---------------------------------------
used_type_ids = set()

for el in FilteredElementCollector(doc).WhereElementIsNotElementType():
    try:
        if el.ViewSpecific:
            continue

        cat = el.Category
        if not cat or cat.Id.IntegerValue not in ALLOWED_CAT_IDS:
            continue

        type_id = el.GetTypeId()
        if type_id != ElementId.InvalidElementId:
            used_type_ids.add(type_id)

    except Exception:
        continue

# ---------------------------------------
# STEP 2 — RESOLVE TYPES
# ---------------------------------------
rows = []

for type_id in used_type_ids:
    etype = doc.GetElement(type_id)
    if not isinstance(etype, ElementType):
        continue

    cat = etype.Category
    if not cat:
        continue

    name_param = etype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    if not name_param:
        continue

    type_name = name_param.AsString()
    if not type_name:
        continue

    comment_param = etype.LookupParameter("Type Comments")
    type_comments = comment_param.AsString() if comment_param else ""

    rows.append([
        cat.Name,
        type_name,
        type_comments or ""
    ])

# ---------------------------------------
# SORT → CATEGORY A–Z, THEN TYPE A–Z
# ---------------------------------------
rows.sort(key=lambda r: (r[0].lower(), r[1].lower()))

# ---------------------------------------
# WRITE CSV (OVERWRITE INTENTIONALLY)
# ---------------------------------------
with codecs.open(output_path, "w", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["Category", "Type", "Type Comments"])
    writer.writerows(rows)

print("CSV created successfully")
print("File:", output_path)
print("Used types exported:", len(rows))
