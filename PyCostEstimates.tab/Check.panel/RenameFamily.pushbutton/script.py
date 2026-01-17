# -*- coding: utf-8 -*-
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    FamilySymbol,
    WallType,
    FloorType,
    RoofType,
    WallFoundationType,
    ElementType,
    BuiltInCategory,
    BuiltInParameter,
    CategoryType,
    FilteredElementCollector
)
import csv
import os
import unicodedata

doc = revit.doc

# --------------------------------------------------
# CONFIG
# --------------------------------------------------
CSV_COLUMN = "Type"

# --------------------------------------------------
# RESOLVE SHARED CSV PATH
# --------------------------------------------------
script_dir = os.path.dirname(script.get_bundle_file("script.py"))
extension_root = os.path.abspath(os.path.join(script_dir, "..", "..", ".."))

csv_file = os.path.join(
    extension_root,
    "PyCostEstimates.tab",
    "Update.panel",
    "Apply Rate.pushbutton",
    "recipes.csv"
)

if not os.path.exists(csv_file):
    forms.alert(
        "recipes.csv not found at:\n\n{}".format(csv_file),
        exitscript=True
    )

# --------------------------------------------------
# TEXT CLEANUP
# --------------------------------------------------
def clean_text(s):
    if not s:
        return None
    s = unicodedata.normalize("NFKC", s)
    s = s.replace(u"\xa0", u" ")
    s = s.replace("\t", " ")
    return " ".join(s.split()).strip()

# --------------------------------------------------
# LOAD CSV (IRONPYTHON SAFE)
# --------------------------------------------------
csv_names = []

with open(csv_file, "rb") as f:
    raw = f.read()

try:
    text = raw.decode("utf-8-sig")
except:
    text = raw.decode("cp1252", errors="ignore")

lines = text.splitlines()
reader = csv.DictReader(lines)

if CSV_COLUMN not in reader.fieldnames:
    forms.alert(
        "CSV column '{}' not found.\nFound:\n{}"
        .format(CSV_COLUMN, ", ".join(reader.fieldnames)),
        exitscript=True
    )

for row in reader:
    val = clean_text(row.get(CSV_COLUMN))
    if val:
        csv_names.append(val)

csv_names = list(dict.fromkeys(csv_names))

# --------------------------------------------------
# COLLECT MODEL TYPES (CATEGORY-BASED, CORRECT)
# --------------------------------------------------
SUPPORTED_CATEGORIES = {
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_Roofs,
    BuiltInCategory.OST_StructuralFoundation,
    BuiltInCategory.OST_Fascia,
    BuiltInCategory.OST_Gutter,
    BuiltInCategory.OST_RoofSoffit
}

def get_type_name(t):
    p = t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    return p.AsString() if p else t.Name

family_dict = {}

for t in (
    FilteredElementCollector(doc)
    .OfClass(ElementType)
    .ToElements()
):
    cat = t.Category
    if not cat:
        continue
    if cat.CategoryType != CategoryType.Model:
        continue
    if cat.IsTagCategory:
        continue
    if cat.Id.IntegerValue not in [int(c) for c in SUPPORTED_CATEGORIES]:
        continue

    key = "{} : {}".format(cat.Name, get_type_name(t))
    family_dict[key] = t

if not family_dict:
    forms.alert("No valid model types found.", exitscript=True)

# --------------------------------------------------
# INTERACTIVE RENAME LOOP
# --------------------------------------------------
with revit.Transaction("Rename Model Types from CSV"):

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
                param = elem_type.get_Parameter(
                    BuiltInParameter.SYMBOL_NAME_PARAM
                )
                if not param or param.IsReadOnly:
                    raise Exception("Name parameter is read-only")
                param.Set(csv_name)
            else:
                elem_type.Name = csv_name

        except Exception as e:
            forms.alert(
                "Cannot rename this type:\n{}\n\n{}"
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

forms.alert(
    "Renaming complete.\n\n"
    "Build-up names applied successfully."
)
