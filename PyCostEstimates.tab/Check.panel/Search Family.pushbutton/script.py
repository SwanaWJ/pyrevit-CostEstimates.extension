# -*- coding: utf-8 -*-

from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import BuiltInParameter
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc

# --------------------------------------------------
# Ask user for TYPE name
# --------------------------------------------------
search = forms.ask_for_string(
    prompt="Enter TYPE name to highlight (contains):",
    title="Search by TYPE Name"
)

if not search:
    forms.alert("No TYPE name entered.")
    raise SystemExit

search = search.lower()

# --------------------------------------------------
# STEP 1: Find matching ELEMENT TYPES
# --------------------------------------------------
matching_type_ids = set()

type_collector = (
    DB.FilteredElementCollector(doc)
    .WhereElementIsElementType()
)

for t in type_collector:
    try:
        p = t.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if not p:
            continue

        type_name = p.AsString()
        if not type_name:
            continue

        if search in type_name.lower():
            matching_type_ids.add(t.Id)

    except:
        continue

if not matching_type_ids:
    forms.alert("No matching TYPE names found.")
    raise SystemExit

# --------------------------------------------------
# STEP 2: Find INSTANCES using those TYPES (FIXED)
# --------------------------------------------------
found_ids = List[DB.ElementId]()

inst_collector = (
    DB.FilteredElementCollector(doc)
    .WhereElementIsNotElementType()
)

for e in inst_collector:
    try:
        # Case 1: FamilyInstance (Windows, Doors, etc.)
        if isinstance(e, DB.FamilyInstance):
            if e.Symbol and e.Symbol.Id in matching_type_ids:
                found_ids.Add(e.Id)
                continue

        # Case 2: System families (Walls, Floors, etc.)
        type_id = e.GetTypeId()
        if type_id and type_id in matching_type_ids:
            found_ids.Add(e.Id)

    except:
        continue

if found_ids.Count == 0:
    forms.alert("Type found, but no instances placed in model.")
    raise SystemExit

# --------------------------------------------------
# STEP 3: Highlight + Zoom
# --------------------------------------------------
uidoc.Selection.SetElementIds(found_ids)
uidoc.ShowElements(found_ids)

forms.alert(
    "Highlighted {} element(s).".format(found_ids.Count),
    title="Search Result"
)
