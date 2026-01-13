# -*- coding: utf-8 -*-
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import *
import csv
import os

doc = revit.doc

CSV_COLUMN = "Type"

# --------------------------------------------------
# CSV
# --------------------------------------------------
this_pushbutton = os.path.dirname(script.get_bundle_file("script.py"))
tab_dir = os.path.dirname(os.path.dirname(this_pushbutton))
csv_file = os.path.join(tab_dir, "Update.panel", "Apply Rate.pushbutton", "recipes.csv")

if not os.path.exists(csv_file):
    forms.alert("CSV file not found:\n{}".format(csv_file), exitscript=True)

csv_names = []
with open(csv_file, "r") as f:
    reader = csv.DictReader(f)
    if CSV_COLUMN not in reader.fieldnames:
        forms.alert("CSV column '{}' not found".format(CSV_COLUMN), exitscript=True)
    for row in reader:
        if row.get(CSV_COLUMN):
            csv_names.append(row[CSV_COLUMN].strip())

csv_names = list(dict.fromkeys(csv_names))

# --------------------------------------------------
# COLLECT TYPES (NO OVER-FILTERING)
# --------------------------------------------------
type_classes = [
    FamilySymbol,
    WallType,
    FloorType,
    RoofType,
    WallFoundationType
]

def get_type_name(t):
    p = t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    return p.AsString() if p else t.Name

family_dict = {}

for cls in type_classes:
    for t in FilteredElementCollector(doc).OfClass(cls).WhereElementIsElementType():
        cat = t.Category
        if not cat:
            continue
        if cat.CategoryType != CategoryType.Model:
            continue
        if cat.IsTagCategory:
            continue

        key = "{} : {}".format(cat.Name, get_type_name(t))
        family_dict[key] = t

if not family_dict:
    forms.alert("No model types found.", exitscript=True)

# --------------------------------------------------
# RENAME LOOP (LET REVIT DECIDE)
# --------------------------------------------------
with revit.Transaction("Rename Model Types"):

    while csv_names and family_dict:

        csv_name = forms.SelectFromList.show(
            csv_names,
            title="Select Build-up Name",
            button_name="Use Name"
        )
        if not csv_name:
            break

        fam_key = forms.SelectFromList.show(
            sorted(family_dict.keys()),
            title="Select Model Type to Rename",
            button_name="Rename"
        )
        if not fam_key:
            break

        elem_type = family_dict[fam_key]
        old_name = get_type_name(elem_type)

        try:
            if isinstance(elem_type, FamilySymbol):
                param = elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if not param or param.IsReadOnly:
                    raise Exception("Read-only name")
                param.Set(csv_name)
            else:
                elem_type.Name = csv_name

        except Exception as e:
            forms.alert(
                "Cannot rename:\n{}\n\n{}"
                .format(old_name, str(e)),
                warn_icon=True
            )
            continue

        csv_names.remove(csv_name)
        family_dict.pop(fam_key)

        if not forms.alert(
            "Renamed:\n{}\n→ {}\n\nContinue?"
            .format(old_name, csv_name),
            ok=False, yes=True, no=True
        ):
            break

forms.alert("Renaming complete.")
