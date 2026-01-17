# -*- coding: utf-8 -*-

import os
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
SHARED_PARAM_FILENAME = "PyCostEstimates_SharedParameters.txt"

output = script.get_output()

try:
    app = revit.doc.Application

    # ---------------------------------------------------------
    # 1. ENSURE SHARED PARAMETER FILE EXISTS & IS SET
    # ---------------------------------------------------------
    sp_file = app.OpenSharedParameterFile()

    if not sp_file:
        # Create shared parameter file in Documents
        documents_path = os.path.join(
            os.path.expanduser("~"),
            "Documents"
        )
        sp_file_path = os.path.join(
            documents_path,
            SHARED_PARAM_FILENAME
        )

        # Create blank file if it doesn't exist
        if not os.path.exists(sp_file_path):
            with open(sp_file_path, "w") as f:
                f.write("")  # Revit expects a plain text file

        # Tell Revit to use this file
        app.SharedParametersFilename = sp_file_path

        # Try opening again
        sp_file = app.OpenSharedParameterFile()

        if not sp_file:
            forms.alert(
                "Failed to create or open shared parameter file.",
                exitscript=True
            )

    # ---------------------------------------------------------
    # 2. GET OR CREATE GROUP
    # ---------------------------------------------------------
    group = sp_file.Groups.get_Item(GROUP_NAME)
    if not group:
        group = sp_file.Groups.Create(GROUP_NAME)

    # ---------------------------------------------------------
    # 3. CHECK FOR DUPLICATE PARAMETER
    # ---------------------------------------------------------
    for definition in group.Definitions:
        if definition.Name == PARAM_NAME:
            forms.alert(
                "Shared parameter '{}' already exists.".format(PARAM_NAME),
                exitscript=True
            )

    # ---------------------------------------------------------
    # 4. CREATE PARAMETER (VERSION SAFE)
    # ---------------------------------------------------------
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
        "Shared parameter '{}' created successfully.\n\n"
        "Shared parameter file:\n{}".format(
            PARAM_NAME,
            app.SharedParametersFilename
        )
    )

except Exception as e:
    output.print_md("## ❌ Failed to create shared parameter")
    output.print_md("```")
    output.print_md(str(e))
    output.print_md("```")
    raise
