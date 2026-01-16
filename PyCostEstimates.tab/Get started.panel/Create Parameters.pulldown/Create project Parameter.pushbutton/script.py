from Autodesk.Revit.DB import *
from pyrevit import revit, forms
import os

doc = revit.doc
app = doc.Application

PARAM_NAME = "xy"
PARAM_GROUP_NAME = "BOQ"
PARAM_GROUP = BuiltInParameterGroup.PG_DATA

# -------------------------------------------------
# DETERMINE PARAMETER SPEC (2019–2025 SAFE)
# -------------------------------------------------
if int(app.VersionNumber) >= 2022:
    PARAM_SPEC = SpecTypeId.Currency
else:
    PARAM_SPEC = ParameterType.Currency  # fallback for 2019–2021

# -------------------------------------------------
# ENSURE SHARED PARAMETER FILE EXISTS
# -------------------------------------------------
sp_path = os.path.join(
    os.environ["USERPROFILE"],
    "Documents",
    "SharedParameters_BOQ.txt"
)

if not os.path.exists(sp_path):
    with open(sp_path, "w") as f:
        f.write("# BOQ Shared Parameters\n")

app.SharedParametersFilename = sp_path
sp_file = app.OpenSharedParameterFile()

if not sp_file:
    forms.alert("Failed to open Shared Parameter file", exitscript=True)

# -------------------------------------------------
# GET OR CREATE GROUP
# -------------------------------------------------
group = next((g for g in sp_file.Groups if g.Name == PARAM_GROUP_NAME), None)
if not group:
    group = sp_file.Groups.Create(PARAM_GROUP_NAME)

# -------------------------------------------------
# GET OR CREATE DEFINITION
# -------------------------------------------------
definition = next((d for d in group.Definitions if d.Name == PARAM_NAME), None)

if not definition:
    opts = ExternalDefinitionCreationOptions(PARAM_NAME, PARAM_SPEC)
    opts.Visible = True
    opts.UserModifiable = True
    definition = group.Definitions.Create(opts)

# -------------------------------------------------
# STRICT CURRENCY-SAFE CATEGORY WHITELIST
# -------------------------------------------------
ALLOWED_BICS = [
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_Roofs,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_Doors,
    BuiltInCategory.OST_Windows,
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_ElectricalEquipment
]

cat_set = app.Create.NewCategorySet()

for bic in ALLOWED_BICS:
    cat = doc.Settings.Categories.get_Item(bic)
    if cat and cat.AllowsBoundParameters:
        cat_set.Insert(cat)

if cat_set.IsEmpty:
    forms.alert("No valid categories found", exitscript=True)

# -------------------------------------------------
# CREATE INSTANCE PROJECT PARAMETER
# -------------------------------------------------
binding = app.Create.NewInstanceBinding(cat_set)

with revit.Transaction("Create Project Parameter: xy (Currency)"):
    doc.ParameterBindings.Insert(definition, binding, PARAM_GROUP)

forms.alert("Currency instance parameter 'xy' created successfully")
