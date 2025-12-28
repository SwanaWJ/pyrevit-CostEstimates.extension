import os
import subprocess
from pyrevit import script

# -----------------------------------------------------------------------------
# Resolve recipes.csv path
# -----------------------------------------------------------------------------
script_dir = os.path.dirname(__file__)

recipes_csv = os.path.abspath(
    os.path.join(
        script_dir,
        "..", "..",
        "Rate.panel",
        "Rate.pushbutton",
        "recipes.csv"
    )
)

# -----------------------------------------------------------------------------
# Fail silently if missing (Revit-safe)
# -----------------------------------------------------------------------------
if not os.path.exists(recipes_csv):
    script.exit()

# -----------------------------------------------------------------------------
# Simple debounce: do not relaunch if Excel already has it open
# -----------------------------------------------------------------------------
try:
    subprocess.Popen(
        ["cmd", "/c", "start", "", recipes_csv],
        shell=False
    )
except Exception:
    pass
