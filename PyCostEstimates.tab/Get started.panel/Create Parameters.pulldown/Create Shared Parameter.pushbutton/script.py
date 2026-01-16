# -*- coding: utf-8 -*-

from pyrevit import revit, forms, script

# Revit API (version-safe)
try:
    # Revit 2022+
    from Autodesk.Revit.DB import ExternalDefinitionCreationOptions, SpecTypeId
    USE_SPEC_TYPE = True
except ImportError:
    # Revit 2021 and earlier
    from Autodesk.Revit.DB import ExternalDefinitionCreationOptions, ParameterType
    USE_SPEC_TYPE = False


PARAM_NAME = "Amount (Qty*Rate)"
GROUP_NAME = "BOQ"

output = script.get_output()

try:
    # ✅ THIS is the correct, universal way in pyRevit
    app = revit.doc.Application

    sp_file = app.OpenSharedParameterFile()
    if not sp_file:
        forms.alert(
            "No shared parameter file is set.\n"
            "Go to Manage → Shared Parameters and set one.",
            exitscript=True
        )

    # Get or create group
    group = sp_file.Groups.get_Item(GROUP_NAME)
    if not group:
        group = sp_file.Groups.Create(GROUP_NAME)

    # Check for duplicates
    for definition in group.Definitions:
        if definition.Name == PARAM_NAME:
            forms.alert(
                "Shared parameter '{}' already exists.".format(PARAM_NAME),
                exitscript=True
            )

    # Create parameter (version-safe)
    if USE_SPEC_TYPE:
        options = ExternalDefinitionCreationOptions(
            PARAM_NAME,
            SpecTypeId.Number
        )
    else:
        options = ExternalDefinitionCreationOptions(
            PARAM_NAME,
            ParameterType.Number
        )

    options.Visible = True
    group.Definitions.Create(options)

    forms.alert(
        "Shared parameter '{}' (Number) created successfully.".format(PARAM_NAME)
    )

except Exception as e:
    output.print_md("## ❌ Failed to create shared parameter")
    output.print_md("```")
    output.print_md(str(e))
    output.print_md("```")
    raise
