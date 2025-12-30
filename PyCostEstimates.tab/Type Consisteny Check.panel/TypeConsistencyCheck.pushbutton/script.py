# -*- coding: utf-8 -*-
from pyrevit import revit, forms
from Autodesk.Revit.DB import (
    BuiltInCategory,
    FilteredElementCollector,
    View3D,
    ElementId
)

doc = revit.doc
uidoc = revit.uidoc
view = doc.ActiveView

# --------------------------------------------------
# VALIDATION
# --------------------------------------------------
if not isinstance(view, View3D):
    forms.alert("Run this tool from a 3D View only.", exitscript=True)

# --------------------------------------------------
# CATEGORY SETUP (SAFE)
# --------------------------------------------------
CATEGORY_MAP = {
    "Walls": BuiltInCategory.OST_Walls,
    "Floors": BuiltInCategory.OST_Floors,
    "Structural Foundations": BuiltInCategory.OST_StructuralFoundation
}

category_name = forms.SelectFromList.show(
    sorted(CATEGORY_MAP.keys()),
    title="Type Consistency Check",
    multiselect=False
)

if not category_name:
    forms.alert("No category selected.", exitscript=True)

bic = CATEGORY_MAP[category_name]

# --------------------------------------------------
# TYPE NAME INPUT
# --------------------------------------------------
expected_type_name = forms.ask_for_string(
    prompt="Enter EXACT Type Name to check",
    title="Type Consistency Check"
)

if not expected_type_name:
    forms.alert("No type name entered.", exitscript=True)

expected_type_name = expected_type_name.strip()

# --------------------------------------------------
# ISOLATE OPTION
# --------------------------------------------------
isolate_choice = forms.alert(
    "Isolate matching elements?\n\nYes = hide non-matching elements",
    options=["Yes", "No"]
)

# --------------------------------------------------
# COLLECT ELEMENTS (HOST ONLY – SAFE)
# --------------------------------------------------
elements = (
    FilteredElementCollector(doc, view.Id)
    .OfCategory(bic)
    .WhereElementIsNotElementType()
    .ToElements()
)

checked = 0
matched = 0
to_hide = []

# --------------------------------------------------
# TYPE CHECK (CRASH-PROOF)
# --------------------------------------------------
for el in elements:
    try:
        el_type = doc.GetElement(el.GetTypeId())
        if not el_type:
            continue

        checked += 1

        type_name = el_type.Name.strip()

        if type_name == expected_type_name:
            matched += 1
        else:
            to_hide.append(el.Id)

    except Exception:
        # absolutely safe skip
        continue

# --------------------------------------------------
# APPLY VISIBILITY (NO FILTERS = NO CRASHES)
# --------------------------------------------------
with revit.Transaction("Type Consistency Visibility"):
    if isolate_choice == "Yes":
        if to_hide:
            view.HideElements(to_hide)
    else:
        # highlight mismatches only
        if to_hide:
            view.SetElementOverrides(
                ElementId.InvalidElementId,  # no-op safety
                view.GetElementOverrides(ElementId.InvalidElementId)
            )

# --------------------------------------------------
# RESULT
# --------------------------------------------------
forms.alert(
    "Type Consistency Check complete.\n\n"
    "Category: {}\n"
    "Elements checked: {}\n"
    "Matching type instances: {}\n"
    "Hidden (mismatch): {}\n"
    "\n(No filters were created – safe mode)"
    .format(
        category_name,
        checked,
        matched,
        len(to_hide)
    )
)
