# -*- coding: utf-8 -*-

import csv
import os
import codecs

from Autodesk.Revit.DB import *
from pyrevit import revit, forms

doc = revit.doc

# --------------------------------------------------
# CSV LOCATION (relative & safe)
# --------------------------------------------------
# Both buttons are in the SAME panel, so we navigate up
script_dir = os.path.dirname(__file__)

csv_path = os.path.normpath(os.path.join(
    script_dir,
    "..",  # back to BOQ Description.pulldown
    "Extract model data.pushbutton",
    "FamilyTypes_With_Comments.csv"
))

if not os.path.exists(csv_path):
    forms.alert(
        "CSV file not found:\n\n{}".format(csv_path),
        exitscript=True
    )

# --------------------------------------------------
# READ CSV → { type_name : type_comment }
# --------------------------------------------------
type_comment_map = {}

with codecs.open(csv_path, "r", encoding="utf-8") as csvfile:
    reader = csv.DictReader(csvfile)

    for row in reader:
        type_name = row.get("Type")
        comment = row.get("Type Comments")

        if type_name:
            type_comment_map[type_name.strip()] = (comment or "").strip()

if not type_comment_map:
    forms.alert("CSV contains no usable data.", exitscript=True)

# --------------------------------------------------
# COLLECT FAMILY TYPES (LOADABLE ONLY)
# --------------------------------------------------
collector = FilteredElementCollector(doc).OfClass(FamilySymbol)

updated = 0
skipped = 0

# --------------------------------------------------
# APPLY TYPE COMMENTS
# --------------------------------------------------
with revit.Transaction("Update Type Comments from CSV"):
    for symbol in collector:
        try:
            # Type name
            name_param = symbol.get_Parameter(
                BuiltInParameter.SYMBOL_NAME_PARAM
            )
            if not name_param:
                continue

            type_name = name_param.AsString()
            if not type_name:
                continue

            if type_name not in type_comment_map:
                continue

            # Type Comments
            comment_param = symbol.LookupParameter("Type Comments")
            if not comment_param or comment_param.IsReadOnly:
                skipped += 1
                continue

            comment_param.Set(type_comment_map[type_name])
            updated += 1

        except Exception:
            skipped += 1
            continue

# --------------------------------------------------
# REPORT
# --------------------------------------------------
forms.alert(
    "Type Comments update complete.\n\n"
    "Updated: {}\n"
    "Skipped: {}".format(updated, skipped)
)
