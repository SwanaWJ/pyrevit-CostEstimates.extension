# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, forms
from System.Collections.Generic import List

doc = revit.doc
view = doc.ActiveView

# --------------------------------------------------
# USER INPUT
# --------------------------------------------------
search_text = forms.ask_for_string(
    prompt="Enter FAMILY TYPE name (contains):",
    title="Type Consistency Check"
)

if not search_text:
    forms.alert("No text entered. Command cancelled.", exitscript=True)

# --------------------------------------------------
# FIND CATEGORIES THAT SUPPORT TYPE NAME
# --------------------------------------------------
type_name_param_id = ElementId(BuiltInParameter.ALL_MODEL_TYPE_NAME)

valid_cat_ids = List[ElementId]()

for cat in doc.Settings.Categories:
    try:
        if cat.CategoryType != CategoryType.Model:
            continue
        if not cat.AllowsBoundParameters:
            continue

        params = ParameterFilterUtilities.GetFilterableParametersInCommon(
            doc, List[ElementId]([cat.Id])
        )

        if type_name_param_id in params:
            valid_cat_ids.Add(cat.Id)

    except:
        continue

if valid_cat_ids.Count == 0:
    forms.alert(
        "No model categories support Type Name filtering in this document.",
        exitscript=True
    )

# --------------------------------------------------
# PARAMETER PROVIDER
# --------------------------------------------------
provider = ParameterValueProvider(type_name_param_id)

# --------------------------------------------------
# FILTER RULES
# --------------------------------------------------
contains_rule = FilterStringRule(
    provider,
    FilterStringContains(),
    search_text
)

# Inverted rule → everything else
not_contains_rule = FilterStringRule(
    provider,
    FilterStringContains(),
    search_text
)

match_filter = ElementParameterFilter(contains_rule)
other_filter = ElementParameterFilter(not_contains_rule, True)

# --------------------------------------------------
# FILTER NAMES
# --------------------------------------------------
safe_name = search_text.replace(" ", "_")
match_name = "BOQ_MATCH_{}".format(safe_name)
other_name = "BOQ_OTHER_{}".format(safe_name)

# --------------------------------------------------
# TRANSACTION
# --------------------------------------------------
t = Transaction(doc, "Isolate Family Type by Name")
t.Start()

def get_or_create_filter(name, elem_filter):
    for f in FilteredElementCollector(doc).OfClass(ParameterFilterElement):
        if f.Name == name:
            return f

    pf = ParameterFilterElement.Create(
        doc,
        name,
        valid_cat_ids
    )
    pf.SetElementFilter(elem_filter)
    return pf

match_pf = get_or_create_filter(match_name, match_filter)
other_pf = get_or_create_filter(other_name, other_filter)

# Apply to active view
for pf in [match_pf, other_pf]:
    if not view.IsFilterApplied(pf.Id):
        view.AddFilter(pf.Id)

# Visibility logic
view.SetFilterVisibility(match_pf.Id, True)
view.SetFilterVisibility(other_pf.Id, False)

t.Commit()

# --------------------------------------------------
# FEEDBACK
# --------------------------------------------------
TaskDialog.Show(
    "TypeConsistencyCheck",
    "Isolation complete.\n\n"
    "Only family types matching:\n\n"
    "'{}'\n\n"
    "are visible.".format(search_text)
)
