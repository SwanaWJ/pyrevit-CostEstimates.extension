# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms
from collections import defaultdict

# --- Settings ---
PARAM_NAME = "Test_1234"
doc = revit.doc

# --- Initialize collectors ---
elements = DB.FilteredElementCollector(doc)\
    .WhereElementIsNotElementType()\
    .ToElements()

category_totals = defaultdict(float)
category_counts = defaultdict(int)
grand_total = 0.0
total_count = 0

# --- Process elements ---
for elem in elements:
    try:
        param = elem.LookupParameter(PARAM_NAME)
        if param and param.HasValue and param.StorageType == DB.StorageType.Double:
            value = param.AsDouble()
            if value > 0:
                cat_name = elem.Category.Name if elem.Category else "Uncategorized"
                category_totals[cat_name] += value
                category_counts[cat_name] += 1
                grand_total += value
                total_count += 1
    except:
        continue

# --- Build message ---
message = "**Total of Test_1234 across {} elements:**\n\n".format(total_count)
message += "ZAR {:.2f}\n\n".format(grand_total)
message += "**Category Breakdown:**\n"
for cat in sorted(category_totals.keys()):
    message += "- {} ({}): ZAR {:.2f}\n".format(cat, category_counts[cat], category_totals[cat])

# --- Show popup ---
forms.alert(message, title="Test_1234 Totals by Category", warn_icon=True)
