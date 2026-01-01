# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    ParameterFilterElement,
    ElementId,
    ElementParameterFilter,
    FilterStringRule,
    FilterStringEquals,
    Transaction
)

from pyrevit import revit, forms

doc = revit.doc
view = doc.ActiveView

# -------------------------------
# USER INPUT
# -------------------------------
category_map = {
    "Walls": BuiltInCategory.OST_Walls,
    "Floors": BuiltInCategory.OST_Floors,
    "Structural Foundations": BuiltInCategory.OST_StructuralFoundation
}

category_name = forms.ask_for_one_item(
    sorted(category_map.keys()),
    title="Select Category"
)

if not category_name:
    forms.alert("No category selected.", exitscript=True)

type_name = forms.ask_for_string(
    prompt="Enter EXACT Family Type Name:",
    title="Type Consistency Check"
)

if not type_name:
    forms.alert("No type name entered.", exitscript=True)

bic = category_map[category_name]

# -------------------------------
# COLLECT ELEMENTS (VALIDATION)
# -------------------------------
elements = list(
    FilteredElementCollector(doc, view.Id)
    .OfCategory(bic)
    .WhereElementIsNotElementType()
)

if not elements:
    forms.alert("No elements found in this category in the active view.", exitscript=True)

# -------------------------------
# BUILD FILTER (CORRECT API)
# -------------------------------
param_id = ElementId(BuiltInParameter.SYMBOL_NAME_PARAM)
evaluator = FilterStringEquals()
rule = FilterStringRule(param_id, evaluator, type_name)
param_filter = ElementParameterFilter(rule)

filter_name = "TCHECK_" + category_name + "_" + type_name

# -------------------------------
# TRANSACTION
# -------------------------------
t = Transaction(doc, "Type Consistency Check")
t.Start()

# Remove existing filter with same name
for f in FilteredElementCollector(doc).OfClass(ParameterFilterElement):
    if f.Name == filter_name:
        doc.Delete(f.Id)

# Create filter
filter_elem = ParameterFilterElement.Create(
    doc,
    filter_name,
    [ElementId(bic)]
)

filter_elem.SetElementFilter(param_filter)

# Apply filter to view
view.AddFilter(filter_elem.Id)
view.SetFilterVisibility(filter_elem.Id, True)

# Hide everything that does NOT match
view.SetFilterOverrides(
    filter_elem.Id,
    view.GetFilterOverrides(filter_elem.Id)
)

t.Commit()

# -------------------------------
# USER FEEDBACK
# -------------------------------
forms.alert(
    "Filter applied.\n\n"
    "Only elements with type name:\n\n"
    "'{}'\n\n"
    "are now visible.\n\n"
    "If elements disappeared, they are wrongly named."
    .format(type_name),
    title="Type Consistency Check"
)
