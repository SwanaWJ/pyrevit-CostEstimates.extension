# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ParameterFilterElement,
    Transaction
)
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit

doc = revit.doc
view = doc.ActiveView

# --------------------------------------------------
# FIND BOQ FILTERS APPLIED TO VIEW
# --------------------------------------------------
filters_to_remove = []

for fid in view.GetFilters():
    pf = doc.GetElement(fid)
    if not pf:
        continue

    if pf.Name.startswith("BOQ_MATCH_") or pf.Name.startswith("BOQ_OTHER_"):
        filters_to_remove.append(pf.Id)

# --------------------------------------------------
# TRANSACTION
# --------------------------------------------------
t = Transaction(doc, "Restore Model Visibility")
t.Start()

for fid in filters_to_remove:
    view.RemoveFilter(fid)

t.Commit()

# --------------------------------------------------
# FEEDBACK
# --------------------------------------------------
if filters_to_remove:
    TaskDialog.Show(
        "Restore Model",
        "Model visibility restored.\n\n"
        "BOQ isolation filters removed from this view."
    )
else:
    TaskDialog.Show(
        "Restore Model",
        "No BOQ isolation filters were found in this view."
    )
