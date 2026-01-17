import os
from pyrevit import forms

# -----------------------------------------------------------------------------
# Resolve extension root
# -----------------------------------------------------------------------------
script_dir = os.path.dirname(__file__)

extension_root = os.path.abspath(
    os.path.join(script_dir, "..", "..", "..")
)

# -----------------------------------------------------------------------------
# Path to the SHARED recipes.csv
# -----------------------------------------------------------------------------
recipes_csv = os.path.join(
    extension_root,
    "PyCostEstimates.tab",
    "Update.panel",
    "Apply Rate.pushbutton",
    "recipes.csv"
)

# -----------------------------------------------------------------------------
# Open CSV safely
# -----------------------------------------------------------------------------
if not os.path.exists(recipes_csv):
    forms.alert(
        "recipes.csv not found at:\n\n{}".format(recipes_csv),
        warn_icon=True
    )
else:
    try:
        os.startfile(recipes_csv)
    except Exception as e:
        forms.alert(
            "Failed to open recipes.csv:\n\n{}\n\n{}"
            .format(recipes_csv, str(e)),
            warn_icon=True
        )
