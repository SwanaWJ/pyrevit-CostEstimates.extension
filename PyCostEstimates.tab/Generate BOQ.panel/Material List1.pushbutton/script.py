# -*- coding: utf-8 -*-
"""
Material List - Stage 1 (AUTO EXTRACT)

- Extract ALL family instances in the model
- Group by family type
- Resolve category to unit
- Compute total Revit quantity per type
"""

from pyrevit import script
from Autodesk.Revit.DB import *
import System
from collections import defaultdict

output = script.get_output()
doc = __revit__.ActiveUIDocument.Document

output.print_md("## Material List - Stage 1")
output.print_md("Extracting model quantities...")

# ------------------------------------------------------------
# CATEGORY TO UNIT MAP (ASCII ONLY)
# ------------------------------------------------------------

CATEGORY_UNIT_MAP = {
    BuiltInCategory.OST_Walls: "m2",
    BuiltInCategory.OST_Floors: "m3",
    BuiltInCategory.OST_Roofs: "m2",
    BuiltInCategory.OST_Ceilings: "m2",
    BuiltInCategory.OST_Doors: "No",
    BuiltInCategory.OST_Windows: "No",
    BuiltInCategory.OST_StructuralColumns: "m3",
    BuiltInCategory.OST_StructuralFraming: "m",
    BuiltInCategory.OST_StructuralFoundation: "m3",
    BuiltInCategory.OST_Conduit: "m",
    BuiltInCategory.OST_PipeCurves: "m",
    BuiltInCategory.OST_ElectricalFixtures: "No",
    BuiltInCategory.OST_ElectricalEquipment: "No",
    BuiltInCategory.OST_LightingFixtures: "No",
    BuiltInCategory.OST_LightingDevices: "No",
    BuiltInCategory.OST_PlumbingFixtures: "No",
    BuiltInCategory.OST_PipeFitting: "No",
    BuiltInCategory.OST_PipeAccessory: "No",
    BuiltInCategory.OST_SpecialityEquipment: "No",
}

SUPPORTED_BICS = set(CATEGORY_UNIT_MAP.keys())

# ------------------------------------------------------------
# COLLECT ELEMENTS
# ------------------------------------------------------------

elements = (
    FilteredElementCollector(doc)
    .WhereElementIsNotElementType()
    .ToElements()
)

# ------------------------------------------------------------
# GROUP BY FAMILY TYPE
# ------------------------------------------------------------

type_data = defaultdict(lambda: {
    "unit": None,
    "raw_qty": 0.0
})

def get_bic(elem):
    try:
        return System.Enum.Parse(
            BuiltInCategory,
            str(elem.Category.Id.IntegerValue)
        )
    except:
        return None

def get_raw_quantity(elem, bic):
    if bic == BuiltInCategory.OST_Walls:
        p = elem.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED)
        return p.AsDouble() if p else 0.0

    if bic in (
        BuiltInCategory.OST_Floors,
        BuiltInCategory.OST_StructuralFoundation,
        BuiltInCategory.OST_StructuralColumns
    ):
        p = elem.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED)
        return p.AsDouble() if p else 0.0

    if bic in (
        BuiltInCategory.OST_StructuralFraming,
        BuiltInCategory.OST_PipeCurves,
        BuiltInCategory.OST_Conduit
    ):
        p = elem.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
        return p.AsDouble() if p else 0.0

    return 1.0  # Count-based

# ------------------------------------------------------------
# PROCESS ELEMENTS
# ------------------------------------------------------------

for elem in elements:
    if not elem.Category:
        continue

    bic = get_bic(elem)
    if bic not in SUPPORTED_BICS:
        continue

    type_id = elem.GetTypeId()
    if type_id == ElementId.InvalidElementId:
        continue

    elem_type = doc.GetElement(type_id)
    if not elem_type:
        continue

    p = elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    type_name = p.AsString() if p else "UNKNOWN"

    unit = CATEGORY_UNIT_MAP[bic]
    raw_qty = get_raw_quantity(elem, bic)

    type_data[type_name]["unit"] = unit
    type_data[type_name]["raw_qty"] += raw_qty

# ------------------------------------------------------------
# CONVERT AND OUTPUT
# ------------------------------------------------------------

for type_name, data in sorted(type_data.items()):
    unit = data["unit"]
    raw_qty = data["raw_qty"]

    if unit == "m2":
        qty = UnitUtils.ConvertFromInternalUnits(
            raw_qty, UnitTypeId.SquareMeters
        )
    elif unit == "m3":
        qty = UnitUtils.ConvertFromInternalUnits(
            raw_qty, UnitTypeId.CubicMeters
        )
    elif unit == "m":
        qty = UnitUtils.ConvertFromInternalUnits(
            raw_qty, UnitTypeId.Meters
        )
    else:
        qty = raw_qty

    qty = round(qty, 3)

    output.print_md("### {}".format(type_name))
    output.print_md("**{} : {}**".format(unit, qty))

output.print_md("Stage 1 complete")
