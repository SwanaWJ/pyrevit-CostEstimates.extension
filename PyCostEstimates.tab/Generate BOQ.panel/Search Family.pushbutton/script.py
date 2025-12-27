# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms

doc = revit.doc
uidoc = revit.uidoc

# Ask user for TYPE name
search = forms.ask_for_string(
    prompt="Enter TYPE name to highlight (e.g. Foundation walls_200mm):",
    title="Search by TYPE Name"
)

if not search:
    forms.alert("No TYPE name entered.")
    raise SystemExit

search = search.lower()

# ---------------------------------------------
# STEP 1: Find matching ELEMENT TYPES
# ---------------------------------------------
matching_type_ids = set()

type_collector = DB.FilteredElementCollector(doc).WhereElementIsElementType()

for t in type_collector:
    try:
        if search in t.Name.lower():
            matching_type_ids.add(t.Id)
    except:
        pass

if not matching_type_ids:
    forms.alert("No matching TYPE names found.")
    raise SystemExit

# ---------------------------------------------
# STEP 2: Find INSTANCES using those TYPES
# ---------------------------------------------
found_ids = []

inst_collector = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()

for e in inst_collector:
    try:
        if e.GetTypeId() in matching_type_ids:
            found_ids.append(e.Id)
    except:
        pass

if not found_ids:
    forms.alert("Type found, but no instances placed in model.")
else:
    uidoc.Selection.SetElementIds(found_ids)
    uidoc.ShowElements(found_ids)

    forms.alert(
        "Highlighted {} element(s).".format(len(found_ids)),
        title="Search Result"
    )
