# -*- coding: utf-8 -*-

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    CategorySet,
    InstanceBinding,
    BuiltInParameterGroup
)

output = script.get_output()

PARAM_NAME = "xy"
SHARED_GROUP = "BOQ"
PARAM_GROUP = BuiltInParameterGroup.PG_DATA  # Group parameter under

try:
    doc = revit.doc
    app = doc.Application

    # Open shared parameter file
    sp_file = app.OpenSharedParameterFile()
    if not sp_file:
        forms.alert(
            "No shared parameter file is set.\n"
            "Go to Manage → Shared Parameters and set one.",
            exitscript=True
        )

    # Get shared parameter definition
    group = sp_file.Groups.get_Item(SHARED_GROUP)
    if not group:
        forms.alert(
            "Shared parameter group '{}' not found.".format(SHARED_GROUP),
            exitscript=True
        )

    definition = None
    for d in group.Definitions:
        if d.Name == PARAM_NAME:
            definition = d
            break

    if not definition:
        forms.alert(
            "Shared parameter '{}' not found.".format(PARAM_NAME),
            exitscript=True
        )

    # -------------------------------
    # CATEGORY SELECTION (version-safe)
    # -------------------------------
    all_categories = [
        c for c in doc.Settings.Categories
        if c.AllowsBoundParameters
    ]

    cat_names = sorted([c.Name for c in all_categories])

    selected_names = forms.SelectFromList.show(
        cat_names,
        title="Select categories for '{}'".format(PARAM_NAME),
        multiselect=True
    )

    if not selected_names:
        forms.alert("No categories selected.", exitscript=True)

    cat_set = CategorySet()
    for cat in all_categories:
        if cat.Name in selected_names:
            cat_set.Insert(cat)

    # Instance binding
    binding = InstanceBinding(cat_set)

    # Bind parameter
    with revit.Transaction("Add shared project parameter"):
        success = doc.ParameterBindings.Insert(
            definition,
            binding,
            PARAM_GROUP
        )

        if not success:
            doc.ParameterBindings.ReInsert(
                definition,
                binding,
                PARAM_GROUP
            )

    forms.alert(
        "Shared parameter '{}' added as INSTANCE project parameter.".format(PARAM_NAME)
    )

except Exception as e:
    output.print_md("## ❌ Failed to add project parameter")
    output.print_md("```")
    output.print_md(str(e))
    output.print_md("```")
    raise
