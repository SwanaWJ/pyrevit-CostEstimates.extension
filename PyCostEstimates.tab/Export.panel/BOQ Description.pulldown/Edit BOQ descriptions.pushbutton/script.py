# -*- coding: utf-8 -*-

import os
import subprocess

from pyrevit import forms

# --------------------------------------------------
# Resolve CSV path (relative to this button)
# --------------------------------------------------
script_dir = os.path.dirname(__file__)

csv_path = os.path.normpath(os.path.join(
    script_dir,
    "..",  # back to BOQ Description.pulldown
    "Extract model data.pushbutton",
    "FamilyTypes_With_Comments.csv"
))

# --------------------------------------------------
# Validate CSV exists
# --------------------------------------------------
if not os.path.exists(csv_path):
    forms.alert(
        "BOQ CSV file not found.\n\n"
        "Please run 'Extract model data' first.\n\n"
        "Expected location:\n{}".format(csv_path),
        exitscript=True
    )

# --------------------------------------------------
# Open CSV with default application (Excel)
# --------------------------------------------------
try:
    os.startfile(csv_path)  # Windows-safe (Revit runs on Windows)
except Exception as e:
    forms.alert(
        "Failed to open CSV file.\n\n{}".format(str(e))
    )
