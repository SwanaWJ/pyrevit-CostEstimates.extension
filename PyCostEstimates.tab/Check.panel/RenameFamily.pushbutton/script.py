# -*- coding: utf-8 -*-
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import *
import csv
import os
import unicodedata

doc = revit.doc

# --------------------------------------------------
# CONFIG
# --------------------------------------------------
CSV_COLUMN = "Type"   # MUST match CSV header exactly

# --------------------------------------------------
# LOCATE CSV (AUTHORITATIVE PATH)
# --------------------------------------------------
this_pushbutton = os.path.dirname(script.get_bundle_file("script.py"))
tab_dir = os.path.dirname(os.path.dirname(this_pushbutton))

csv_file = os.path.join(
    tab_dir,
    "Update.panel",
    "Apply Rate.pushbutton",
    "recipes.csv"
)

if not os.path.exists(csv_file):
    forms.alert(
        "CSV file not found:\n{}".format(csv_file),
        exitscript=True
    )

# --------------------------------------------------
# TEXT NORMALISATION (EXCEL / UNICODE SAFE)
# --------------------------------------------------
def clean_text(s):
    if not s:
        return None
    # Normalise Unicode (fix smart characters)
    s = unicodedata.normalize("NFKC", s)
    # Replace non-breaking spaces and tabs
    s = s.replace(u"\xa0", u" ")
    s = s.replace("\t", " ")
    # Collapse whitespace
    return " ".join(s.split()).strip()

# --------------------------------------------------
# LOAD CSV (IRONPYTHON-SAFE, ENCODING-ROBUST)
# --------------------------------------------------
csv_names = []

# Read raw bytes first
with open(csv_file, "rb") as f:
    raw = f.read()

# Try UTF-8 (with BOM), fallback to Windows-1252
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

# Deduplicate while preserving order
seen = set()
csv_names = [n for n in csv_names if not (n in seen or seen.add(n))]

if not csv_names:
    forms.alert("No valid build-up names found in CSV.", exitscript=True)

# --------------------------------------------------
# COLLECT MODEL TYPES (NO OVER-FILTERING)
# --------------------------------------------------
type_classes = [
    FamilySymbol,        # Loadable families
    WallType,            # Walls
    FloorType,           # Floors / slabs / pad foundations
    RoofType,            # Roofs
    WallFoundationType   # Wall / strip foundations
]

def get_type_name(t):
    p = t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    return p.AsString() if p else t.Name

family_dict = {}

for cls in type_classes:
    for t in (
        FilteredElementCollector(doc)
        .OfClass(cls)
        .WhereElementIsElementType()
        .ToElements()
    ):
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
    forms.alert("No valid model types found.", exitscript=True)

# --------------------------------------------------
# INTERACTIVE RENAME LOOP (REVIT-CORRECT)
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
            # Loadable family types
            if isinstance(elem_type, FamilySymbol):
                param = elem_type.get_Parameter(
                    BuiltInParameter.SYMBOL_NAME_PARAM
                )
                if not param or param.IsReadOnly:
                    raise Exception("Name parameter is read-only")
                param.Set(csv_name)

            # System family types
            else:
                elem_type.Name = csv_name

        except Exception as e:
            forms.alert(
                "Cannot rename this type:\n{}\n\n{}"
                .format(old_name, str(e)),
                warn_icon=True
            )
            continue

        # Update state
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
    "CSV was reloaded from disk and applied successfully."
)
