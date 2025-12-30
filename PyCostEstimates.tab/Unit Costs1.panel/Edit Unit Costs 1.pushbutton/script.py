# Material Unit Costs 1 - pyRevit Pushbutton
# Robust CSV locator (no hard-coded folders)

import os
from pyrevit import script, forms

# ---------------------------------------------------------------------
# Start searching from the script directory
# ---------------------------------------------------------------------
start_dir = os.path.dirname(__file__)
target_file = "material_unit_costs.csv"

found_path = None

# Walk up directories (max 6 levels for safety)
current_dir = start_dir
for _ in range(6):
    # Walk through subfolders
    for root, dirs, files in os.walk(current_dir):
        if target_file in files:
            found_path = os.path.join(root, target_file)
            break
    if found_path:
        break
    current_dir = os.path.abspath(os.path.join(current_dir, ".."))

# ---------------------------------------------------------------------
# Validate
# ---------------------------------------------------------------------
if not found_path:
    forms.alert(
        "Material Unit Costs file not found.\n\n"
        "Searched upward from:\n{}".format(start_dir),
        title="Material Unit Costs 1"
    )
    script.exit()

# ---------------------------------------------------------------------
# Open CSV
# ---------------------------------------------------------------------
try:
    os.startfile(found_path)
except Exception as ex:
    forms.alert(
        "Failed to open Material Unit Costs file.\n\n{}".format(ex),
        title="Material Unit Costs 1"
    )
