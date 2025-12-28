from pyrevit import revit, forms
from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    ParameterFilterElement,
    ElementParameterFilter,
    FilterStringRule,
    FilterStringEquals,
    ParameterValueProvider,
    FilteredElementCollector,
    View3D
)

doc = revit.doc
view = doc.ActiveView

# --------------------------------------------------
# SAFETY CHECKS
# --------------------------------------------------
if hasattr(view, 'IsTemporaryHideIsolateActive') and view.IsTemporaryHideIsolateActive():
    forms.alert(
        "Temporary Hide/Isolate is active.\n\n"
        "Please reset it before running this tool.",
        exitscript=True
    )

if not isinstance(view, View3D):
    forms.alert(
        "Please run this tool in a 3D View.",
        exitscript=True
    )

# --------------------------------------------------
# CATEGORY SELECTION
# --------------------------------------------------
categories = [
    c for c in doc.Settings.Categories
    if c.CategoryType.ToString() == "Model"
    and not c.IsTagCategory
]

cat_map = {c.Name: c for c in categories}

category_name = forms.SelectFromList.show(
    sorted(cat_map.keys()),
    title="Type Consistency Check",
    multiselect=False
)

if not category_name:
    forms.alert("No category selected.", exitscript=True)

category = cat_map[category_name]
cat_ids = [category.Id]

# --------------------------------------------------
# TYPE NAME INPUT
# --------------------------------------------------
type_name = forms.ask_for_string(
    prompt="Enter expected Type Name",
    title="Type Consistency Check"
)

if not type_name:
    forms.alert("No type name provided.", exitscript=True)

# --------------------------------------------------
# ISOLATE OPTION
# --------------------------------------------------
isolate = forms.alert(
    "Isolate this type?\n\nYes = hide all other elements in this category",
    options=["Yes", "No"]
)

# --------------------------------------------------
# PARAMETER PROVIDER
# --------------------------------------------------
provider = ParameterValueProvider(
    ElementId(BuiltInParameter.SYMBOL_NAME_PARAM)
)

# --------------------------------------------------
# UNIQUE FILTER NAMES
# --------------------------------------------------
show_filter_name = "TCC_SHOW_{}_{}".format(category.Name, type_name)
hide_filter_name = "TCC_HIDE_{}_{}".format(category.Name, type_name)

# --------------------------------------------------
# TRANSACTION
# --------------------------------------------------
with revit.Transaction("Type Consistency Check"):

    # ---------- DELETE CONFLICTING FILTERS (2023 SAFE) ----------
    existing_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement)
    for f in existing_filters:
        if f.Name in [show_filter_name, hide_filter_name]:
            doc.Delete(f.Id)

    # ---------- SHOW FILTER ----------
    show_rule = FilterStringRule(
        provider,
        FilterStringEquals(),
        type_name
    )

    show_filter = ParameterFilterElement.Create(
        doc,
        show_filter_name,
        cat_ids
    )
    show_filter.SetElementFilter(
        ElementParameterFilter(show_rule)
    )

    view.AddFilter(show_filter.Id)
    view.SetFilterVisibility(show_filter.Id, True)

    # ---------- HIDE FILTER ----------
    if isolate == "Yes":
        hide_rule = FilterStringRule(
            provider,
            FilterStringEquals(),
            type_name
        )

        hide_filter = ParameterFilterElement.Create(
            doc,
            hide_filter_name,
            cat_ids
        )
        hide_filter.SetElementFilter(
            ElementParameterFilter(hide_rule, True)  # inverted
        )

        view.AddFilter(hide_filter.Id)
        view.SetFilterVisibility(hide_filter.Id, False)

# --------------------------------------------------
# DONE
# --------------------------------------------------
forms.alert(
    "Type Consistency Check complete.\n\n"
    "Category: {}\n"
    "Expected Type: {}\n\n"
    "Hidden elements indicate naming errors."
    .format(category.Name, type_name)
)
