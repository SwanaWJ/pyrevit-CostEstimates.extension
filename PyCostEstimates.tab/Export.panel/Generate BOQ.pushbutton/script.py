# -*- coding: utf-8 -*-
import os
import re
import string
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import clr

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox
from pyrevit import revit, DB

# ------------------------------------------------------------------------------
# Save path
# ------------------------------------------------------------------------------
desktop = os.path.expanduser("~/Desktop")
xlsx_path = os.path.join(desktop, "BOQ_Export_From_Model.xlsx")

# ------------------------------------------------------------------------------
# Parameters / constants
# ------------------------------------------------------------------------------
PARAM_COST  = "Cost"
PARAM_TOTAL = "Amount (Qty*Rate)"   # ← FIX (was Test_1234)

TAB_COLORS = {
    "COVER":   "#A6A6A6",
    "BILL1":   "#4472C4",
    "BILL2":   "#C00000",
    "BILL3":   "#FFD966",   # EXTERNAL WORKS tab = yellow
    "SUMMARY": "#70AD47",
}

# Category order for BILL 1 + BILL 2
CATEGORY_ORDER = [
    "Cut and Fill",
    "Structural Foundations",
    "Internal Floors",
    "Internal Walls",
    "Internal Stairs",
    "Block Work in Walls",
    "Structural Columns",
    "Structural Framing",
    "Structural Rebar",
    "Roofs",
    "Ceilings",
    "Windows",
    "Doors",
    "Electrical",
    "Plumbing",
    "Painting",
    "Wall and Floor Finishes",
    "Furniture",
]

# Category order for BILL 3 (external works)
EXTERNAL_WORKS_ORDER = [
    "External Floors",
    "External Walls",
    "External Stairs",
    "Parking",
    "Planting",
    "Site Works",
    "Paving",
    "Drainage",
    "Fencing",
]

VIRTUAL_PAINT    = object()
VIRTUAL_EXTERNAL = object()  # used for manual / split / site categories

CATEGORY_MAP = {
    # CORE / BUILDING / STRUCTURE / MEP
    "Cut and Fill":            DB.BuiltInCategory.OST_Topography,  # plus graded region logic
    "Structural Foundations":  DB.BuiltInCategory.OST_StructuralFoundation,

    # split by Function (internal/external)
    "Internal Floors":         VIRTUAL_EXTERNAL,
    "External Floors":         VIRTUAL_EXTERNAL,
    "Internal Walls":          VIRTUAL_EXTERNAL,
    "External Walls":          VIRTUAL_EXTERNAL,
    "Internal Stairs":         VIRTUAL_EXTERNAL,
    "External Stairs":         VIRTUAL_EXTERNAL,

    "Block Work in Walls":     DB.BuiltInCategory.OST_Walls,
    "Structural Columns":      DB.BuiltInCategory.OST_StructuralColumns,
    "Structural Framing":      DB.BuiltInCategory.OST_StructuralFraming,
    "Structural Rebar":        DB.BuiltInCategory.OST_Rebar,
    "Roofs":                   DB.BuiltInCategory.OST_Roofs,
    "Ceilings":                DB.BuiltInCategory.OST_Ceilings,
    "Windows":                 DB.BuiltInCategory.OST_Windows,
    "Doors":                   DB.BuiltInCategory.OST_Doors,

    "Electrical": [
        DB.BuiltInCategory.OST_Conduit,
        DB.BuiltInCategory.OST_LightingFixtures,
        DB.BuiltInCategory.OST_LightingDevices,
        DB.BuiltInCategory.OST_ElectricalFixtures,
        DB.BuiltInCategory.OST_ElectricalEquipment,
    ],
    "Plumbing": [
        DB.BuiltInCategory.OST_PlumbingFixtures,
        DB.BuiltInCategory.OST_PipeCurves,
        DB.BuiltInCategory.OST_PipeFitting,
        DB.BuiltInCategory.OST_PipeAccessory,
    ],

    "Painting":                VIRTUAL_PAINT,
    "Wall and Floor Finishes": DB.BuiltInCategory.OST_GenericModel,
    "Furniture": [
        DB.BuiltInCategory.OST_Furniture,
        DB.BuiltInCategory.OST_FurnitureSystems,
    ],

    # EXTERNAL WORKS (manual / site)
    "External Floors": VIRTUAL_EXTERNAL,
    "External Walls":  VIRTUAL_EXTERNAL,
    "External Stairs": VIRTUAL_EXTERNAL,
    "Parking":         VIRTUAL_EXTERNAL,
    "Planting":        VIRTUAL_EXTERNAL,
    "Site Works":      VIRTUAL_EXTERNAL,
    "Paving":          VIRTUAL_EXTERNAL,
    "Drainage":        VIRTUAL_EXTERNAL,
    "Fencing":         VIRTUAL_EXTERNAL,
}

_missing_int  = [c for c in CATEGORY_ORDER if c not in CATEGORY_MAP]
_missing_ext  = [c for c in EXTERNAL_WORKS_ORDER if c not in CATEGORY_MAP]
_missing = _missing_int + _missing_ext
if _missing:
    from pyrevit import forms
    forms.alert("Missing in CATEGORY_MAP:\n\n- " + "\n- ".join(_missing),
                title="Category mapping error")
    raise SystemExit

CATEGORY_DESCRIPTIONS = {
    "Cut and Fill": (
        "Bulk earthworks operations including excavation (cut) and embankment (fill), "
        "measured from Revit Topography / Graded Regions, or estimated from Building "
        "Pads if no graded region exists."
    ),
    "Structural Foundations": (
        "Mass or reinforced concrete footings, hardcore bedding, DPM and formwork, "
        "conforming to BS 8000 (earthworks) and BS 8110 (concrete)."
    ),

    "Internal Floors": (
        "In-situ or suspended internal concrete floor slabs, screeds and finishes within "
        "the building footprint."
    ),
    "External Floors": (
        "External slabs, aprons, walkways, ramps and hardscape slabs cast in place, "
        "including preparation, sub-base and finishing, exposed to weather."
    ),

    "Internal Walls": (
        "Internal wall construction including blockwork, plaster, paint, finishes, framing "
        "and associated sundries within the building envelope."
    ),
    "External Walls": (
        "External / retaining walls, upstands, plinth walls and exposed walling to the "
        "perimeter and site works, including finishes and weatherproofing."
    ),

    "Internal Stairs": (
        "Internal stair flights, landings, risers and finishes within the building, "
        "including structural support and balustrades where applicable."
    ),
    "External Stairs": (
        "External stair flights, ramps or stepped access in exposed locations, including "
        "concrete, nosings, drainage slots, balustrades and associated works."
    ),

    "Block Work in Walls": (
        "Concrete block walls, load-bearing or cavity, plastered both sides and painted to "
        "BS 8000-3 masonry workmanship standards, including all mortar, bed-joint "
        "reinforcement, movement provision and finishing to BS 5628-2/-3 quality."
    ),
    "Structural Columns": (
        "Concrete/steel columns with starter bars, ties and shuttering; concrete to spec "
        "per BS 8110-1, steel primed per BS 5493."
    ),
    "Structural Framing": (
        "Mild steel beams and trusses, welded or bolted, treated with primer/paint to "
        "BS 5493 and fabricated per BS 5950."
    ),
    "Structural Rebar": (
        "High-yield deformed steel bars (BS 4449 B500B), cut, bent, fixed and supported "
        "with chairs/spacers, placed per BS 8666 & BS 8110-1."
    ),
    "Roofs": (
        "0.5 mm IBR/IT4 pre-painted roof sheeting fixed to purlins with screws, complete "
        "with ridge capping, insulation and flashings, per BS 5534 & BS 8217."
    ),
    "Ceilings": (
        "Particleboard or PVC tongue-and-groove ceilings, fixed or suspended per BS 5306 "
        "and manufacturer instructions."
    ),
    "Windows": (
        "Aluminium sliding or casement windows with glazing, mosquito nets, stays, "
        "handles and fixings; installed per BS 6262 (glazing) and BS 6375."
    ),
    "Doors": (
        "Timber or engineered doors with hardwood frames, architraves, ironmongery, seals "
        "and painting; installed and fitted as per BS 8214."
    ),
    "Wall and Floor Finishes": (
        "Tiling and screed finishes and plaster/paint to walls, following BS 5385 "
        "(tiling), BS 8203 (screed) and BS 8000 finishing workmanship standards."
    ),
    "Plumbing": (
        "Sanitary appliances (WC pans, cisterns, basins, sinks, urinals) per BS 6465-3, "
        "with associated pipework, fittings, joints, valves, traps and accessories per "
        "BS 5572 sanitary drainage."
    ),
    "Electrical": (
        "Steel conduits per BS 4568-1, armoured cables/junction boxes per SANS 1507/BS 7671, "
        "with lighting fixtures and switchgear as specified."
    ),
    "Painting": (
        "Measured areas from the Revit Paint tool on wall faces (all sides), grouped by "
        "material. Rates use the material 'Cost' if present."
    ),

    # EXTERNAL WORKS descriptions
    "Parking": (
        "External parking areas including formation, preparation, sub-base, basecourse and "
        "final wearing course (asphalt / concrete block paving), line marking, edging and "
        "any associated kerbs."
    ),
    "Planting": (
        "Planting works including topsoil preparation, supply and installation of trees, shrubs, "
        "hedges, grassing and maintenance during the defects liability period, in accordance "
        "with landscape drawings and specifications."
    ),
    "Site Works": (
        "Site preparation, grading, levelling, hardcore fill, compaction, temporary works, "
        "access routes, street furniture and other external site-related works as indicated "
        "on the site development plans."
    ),
    "Paving": (
        "Walkways and paved circulation areas using concrete blocks / pavers on sand bedding, "
        "including compacted sub-base, edge restraints and jointing sand."
    ),
    "Drainage": (
        "Surface water and site drainage including open drains, culverts, manholes, catchpits, "
        "gullies and pipework laid to falls, including bedding and surround."
    ),
    "Fencing": (
        "Site perimeter fencing including posts, rails, mesh / palisade panels, gates and "
        "associated excavation and concrete setting of posts."
    ),
}

# Units
FT3_TO_M3 = 0.0283168
FT2_TO_M2 = 0.092903
FT_TO_M   = 0.3048

# ------------------------------------------------------------------------------
# Workbook & formats
# ------------------------------------------------------------------------------
wb = xlsxwriter.Workbook(xlsx_path, {'constant_memory': True})
try:
    wb.set_calc_on_load()
except AttributeError:
    pass

font = 'Arial Narrow'
def col_fmt(bold=False, italic=False, underline=False, wrap=False, num_fmt=None):
    fmt = {
        'valign': 'top',
        'font_name': font,
        'font_size': 12,
        'border': 1
    }
    if bold: fmt['bold'] = True
    if italic: fmt['italic'] = True
    if underline: fmt['underline'] = True
    if wrap: fmt['text_wrap'] = True
    if num_fmt: fmt['num_format'] = num_fmt
    return wb.add_format(fmt)

fmt_header      = col_fmt(bold=True)
fmt_section     = col_fmt(bold=True)
fmt_description = col_fmt(italic=True, underline=True, wrap=True)
fmt_normal      = col_fmt()
fmt_italic      = col_fmt(italic=True, wrap=True)
fmt_money       = col_fmt(num_fmt='#,##0.00')
fmt_title       = wb.add_format({'bold': True, 'font_name': font, 'font_size': 12, 'align':'left'})
fmt_cover_huge  = wb.add_format({'bold': True, 'font_name': font, 'font_size': 16, 'align': 'center'})
fmt_center      = wb.add_format({'font_name': font, 'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'border': 1})
fmt_text        = wb.add_format({'font_name': font, 'font_size': 12, 'border': 1})
fmt_bold        = wb.add_format({'font_name': font, 'font_size': 12, 'border': 1, 'bold': True})
fmt_wrap        = wb.add_format({'font_name': font, 'font_size': 12, 'border': 1, 'text_wrap': True, 'valign': 'top'})
fmt_percent     = wb.add_format({'font_name': font, 'font_size': 12, 'border': 1, 'num_format': '0.00%'})
fmt_money_right = wb.add_format({'font_name': font, 'font_size': 12, 'border': 1, 'num_format': '#,##0.00', 'align': 'right'})
fmt_noborder    = wb.add_format({'font_name': font, 'font_size': 12})
fmt_text_center = wb.add_format({'font_name': font, 'font_size': 12, 'align': 'center', 'valign': 'vcenter'})

# ------------------------------------------------------------------------------
# Helpers: Project Title / Address
# ------------------------------------------------------------------------------
def _get_project_title():
    pi = revit.doc.ProjectInformation
    pname = None
    p = pi.get_Parameter(DB.BuiltInParameter.PROJECT_NAME) if pi else None
    if p and p.HasValue:
        pname = p.AsString()
    if not pname:
        try:
            import os as _os
            pname = _os.path.splitext(revit.doc.Title)[0]
        except Exception:
            pname = "PROJECT"
    return pname

def _get_project_address():
    pi = revit.doc.ProjectInformation
    addr = None
    p = pi.get_Parameter(DB.BuiltInParameter.PROJECT_ADDRESS) if pi else None
    if p and p.HasValue:
        addr = p.AsString()
    if not addr:
        addr = "PROJECT ADDRESS"
    return addr

TITLE_TEXT = "BILL OF QUANTITIES (BOQ) FOR THE CONSTRUCTION OF {}".format(_get_project_title().upper())

def _safe_sheet_name(name, used):
    s = name.replace(u"–", "-").replace(u"—", "-")
    for ch in '[]:*?/\\':
        s = s.replace(ch, "")
    s = s.strip().strip("'")[:31]
    base = s
    i = 1
    while s in used:
        suf = "({})".format(i)
        s = (base[:31-len(suf)] + suf)
        i += 1
    used.add(s)
    return s

def _is_noise(s):
    s = (s or "").strip()
    if not s:
        return True
    s2 = s.replace(".", "").replace(",", "").replace(" ", "")
    if s2.isdigit():
        return True
    if len(s) < 3:
        return True
    return False

def _item_label(idx):
    return string.ascii_uppercase[idx] if idx < 26 else str(idx + 1)

# ------------------------------------------------------------------------------
# Sheet setup helpers
# ------------------------------------------------------------------------------
def _set_portrait(ws):
    ws.set_paper(9)
    ws.set_portrait()
    ws.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.8)

def init_bill_sheet(name):
    ws = wb.add_worksheet(name)
    _set_portrait(ws)
    ws.merge_range(0, 0, 0, 5, TITLE_TEXT, fmt_title)

    headers = ["ITEM", "DESCRIPTION", "UNIT", "QTY", "RATE (EUR)", "AMOUNT (EUR)"]
    for c, h in enumerate(headers):
        ws.write(1, c, h, fmt_header)

    ws.set_column(1, 1, 45)
    ws.set_column(4, 4, 12)
    ws.set_column(5, 5, 16)
    ws.freeze_panes(2, 0)
    return ws

def init_cover_sheet(name):
    ws = wb.add_worksheet(name)
    _set_portrait(ws)

    ws.set_column("B:D", 50)
    ws.set_row(8, 28)
    ws.set_row(15, 28)
    ws.set_row(19, 28)
    ws.set_row(21, 24)

    ws.merge_range(
        "B9:D9",
        "DEPARTMENT OF HOUSING AND INFRASTRUCTURE DEVELOPMENT",
        fmt_cover_huge
    )
    ws.merge_range(
        "B15:D15",
        "BILL OF QUANTITIES",
        fmt_cover_huge
    )
    ws.merge_range(
        "B17:D17",
        "FOR THE",
        fmt_text_center
    )
    ws.merge_range(
        "B19:D19",
        TITLE_TEXT,
        fmt_cover_huge
    )
    ws.merge_range(
        "B21:D21",
        "AT {}".format(_get_project_address().upper()),
        fmt_text_center
    )

    return ws

def finalize_bill_sheet(ws, row, sheet_cat_order, cat_subtotals):
    ws.write(row, 1, "COLLECTION", fmt_section)
    row += 1
    count = 1
    for cname in sheet_cat_order:
        up = cname.upper()
        cell = cat_subtotals.get(up)
        if cell:
            ws.write(row, 0, str(count), fmt_normal)
            ws.write(row, 1, up, fmt_normal)
            ws.write_formula(row, 5, "={}".format(cell), fmt_money)
            row += 1
            count += 1

    ws.write_blank(row, 0, None, fmt_section)
    ws.write(row, 1, "GRAND TOTAL", fmt_section)
    if cat_subtotals:
        sum_cells = ",".join(
            cat_subtotals[k.upper()]
            for k in sheet_cat_order
            if k.upper() in cat_subtotals
        )
        ws.write_formula(row, 5, "=SUM({})".format(sum_cells), fmt_money)
    else:
        ws.write(row, 5, 0, fmt_money)

    return xl_rowcol_to_cell(row, 5), row

def _sheet_ref(name, cell_addr):
    return "'{}'!{}".format(name.replace("'", "''"), cell_addr)

# ------------------------------------------------------------------------------
# Workbook structure (with BILL 3 - EXTERNAL WORKS)
# ------------------------------------------------------------------------------
_USED_SHEETS = set()
COVER_NAME   = _safe_sheet_name("COVER", _USED_SHEETS)
BILL1_NAME   = _safe_sheet_name("BILL 1 - SUB & SUPERSTRUCTURE", _USED_SHEETS)
BILL2_NAME   = _safe_sheet_name("BILL 2 - MEP", _USED_SHEETS)
BILL3_NAME   = _safe_sheet_name("BILL 3 - EXTERNAL WORKS", _USED_SHEETS)
SUMMARY_NAME = _safe_sheet_name("GENERAL SUMMARY", _USED_SHEETS)

cover_ws = init_cover_sheet(COVER_NAME)
cover_ws.set_tab_color(TAB_COLORS["COVER"])

sheets = {
    BILL1_NAME: {
        "ws": init_bill_sheet(BILL1_NAME),
        "row": 2,
        "cat_counter": 1,
        "cat_subtotals": {},
        "order": []
    },
    BILL2_NAME: {
        "ws": init_bill_sheet(BILL2_NAME),
        "row": 2,
        "cat_counter": 1,
        "cat_subtotals": {},
        "order": []
    },
    BILL3_NAME: {
        "ws": init_bill_sheet(BILL3_NAME),
        "row": 2,
        "cat_counter": 1,
        "cat_subtotals": {},
        "order": []
    },
}

sheets[BILL1_NAME]["ws"].set_tab_color(TAB_COLORS["BILL1"])
sheets[BILL2_NAME]["ws"].set_tab_color(TAB_COLORS["BILL2"])
sheets[BILL3_NAME]["ws"].set_tab_color(TAB_COLORS["BILL3"])  # yellow

BILL_FOR_CATEGORY = {
    "Electrical": BILL2_NAME,
    "Plumbing":   BILL2_NAME,

    # external works live on BILL 3
    "External Floors": BILL3_NAME,
    "External Walls":  BILL3_NAME,
    "External Stairs": BILL3_NAME,
    "Parking":         BILL3_NAME,
    "Planting":        BILL3_NAME,
    "Site Works":      BILL3_NAME,
    "Paving":          BILL3_NAME,
    "Drainage":        BILL3_NAME,
    "Fencing":         BILL3_NAME,

    # internal split categories -> BILL 1
    "Internal Floors": BILL1_NAME,
    "Internal Walls":  BILL1_NAME,
    "Internal Stairs": BILL1_NAME,
}
def _bill_for(cat):
    return BILL_FOR_CATEGORY.get(cat, BILL1_NAME)

# ------------------------------------------------------------------------------
# Painting helper
# ------------------------------------------------------------------------------
def _gather_wall_painting(doc):
    grouped = {}

    def _add(material_name, rate, area_ft2):
        key = "Paint - {}".format(material_name or "Paint")
        qty_m2 = float(area_ft2) * FT2_TO_M2
        if key not in grouped:
            grouped[key] = {
                "qty": 0.0,
                "rate": float(rate or 0.0),
                "unit": "m²",
                "comment": ""
            }
        grouped[key]["qty"] += qty_m2
        if grouped[key]["rate"] == 0.0 and rate:
            grouped[key]["rate"] = float(rate)

    def _rate_from_material(mat):
        try:
            p = mat.LookupParameter(PARAM_COST) if mat else None
            return float(p.AsDouble()) if (p and p.HasValue) else 0.0
        except:
            return 0.0

    def _collect_from_faces(host_elem, faces):
        for f in faces:
            ref = f.Reference
            if not ref or not doc.IsPainted(host_elem.Id, ref):
                continue
            mid = doc.GetPaintedMaterial(host_elem.Id, ref)
            if mid == DB.ElementId.InvalidElementId:
                continue
            mat = doc.GetElement(mid)
            _add(
                mat.Name if mat else "Paint",
                _rate_from_material(mat),
                f.Area
            )

    walls = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Walls)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    opt = DB.Options()
    opt.ComputeReferences = True
    opt.IncludeNonVisibleObjects = False

    for wall in walls:
        try:
            got_any = False
            try:
                for side in (DB.ShellLayerType.Interior, DB.ShellLayerType.Exterior):
                    refs = DB.HostObjectUtils.GetSideFaces(wall, side) or []
                    for ref in refs:
                        if not doc.IsPainted(wall.Id, ref):
                            continue
                        gobj = wall.GetGeometryObjectFromReference(ref)
                        face = gobj if isinstance(gobj, DB.Face) else None
                        if not face:
                            continue
                        mid = doc.GetPaintedMaterial(wall.Id, ref)
                        if mid == DB.ElementId.InvalidElementId:
                            continue
                        mat = doc.GetElement(mid)
                        _add(
                            mat.Name if mat else "Paint",
                            _rate_from_material(mat),
                            face.Area
                        )
                        got_any = True
            except:
                pass
            if got_any:
                continue

            try:
                pids = DB.PartUtils.GetAssociatedParts(revit.doc, wall.Id, True, True)
                if pids and pids.Count > 0:
                    geom_opt = opt
                    for pid in pids:
                        part = revit.doc.GetElement(pid)
                        geom = part.get_Geometry(geom_opt)
                        if not geom:
                            continue
                        for g in geom:
                            if isinstance(g, DB.Solid) and g.Faces:
                                _collect_from_faces(part, list(g.Faces))
                            elif isinstance(g, DB.GeometryInstance):
                                inst = g.GetInstanceGeometry()
                                for gg in inst:
                                    if isinstance(gg, DB.Solid) and gg.Faces:
                                        _collect_from_faces(part, list(gg.Faces))
                    continue
            except:
                pass

            try:
                geom = wall.get_Geometry(opt)
                if geom:
                    for g in geom:
                        if isinstance(g, DB.Solid) and g.Faces:
                            _collect_from_faces(wall, list(g.Faces))
                        elif isinstance(g, DB.GeometryInstance):
                            inst = g.GetInstanceGeometry()
                            for gg in inst:
                                if isinstance(gg, DB.Solid) and gg.Faces:
                                    _collect_from_faces(wall, list(gg.Faces))
            except:
                pass

        except:
            pass

    for v in grouped.values():
        if abs(v["qty"]) < 1e-6:
            v["qty"] = 0.0

    return grouped

# ------------------------------------------------------------------------------
# Earthworks helpers
# ------------------------------------------------------------------------------
_num_pat = re.compile(r"[-+]?\d+(?:[.,]\d+)?")

def _parse_value_string_to_m3_raw(s):
    if not s:
        return 0.0
    s = s.strip()
    m = _num_pat.search(s)
    if not m:
        return 0.0
    val = float(m.group(0).replace(",", "."))
    s_low = s.lower()
    if "ft" in s_low or "ft³" in s_low or "ft^3" in s_low or "cf" in s_low:
        return val * FT3_TO_M3
    return val

def _param_to_m3(p):
    if not p or not p.HasValue:
        return 0.0
    try:
        if p.StorageType == DB.StorageType.Double:
            return p.AsDouble() * FT3_TO_M3
    except:
        pass
    try:
        return _parse_value_string_to_m3_raw(p.AsValueString())
    except:
        return 0.0

def _cutfill_from_elem(elem):
    cut = 0.0
    fill = 0.0
    try:
        cp = elem.get_Parameter(DB.BuiltInParameter.SITE_CUT_VOLUME)
        fp = elem.get_Parameter(DB.BuiltInParameter.SITE_FILL_VOLUME)
        cut  += _param_to_m3(cp)
        fill += _param_to_m3(fp)
    except:
        pass

    if cut <= 1e-9 and fill <= 1e-9:
        for name in (
            "Cut","Fill","Net cut/fill","Net Cut/Fill","Net Cut/Fill Volume",
            "Net cut/fill volume"
        ):
            try:
                p = elem.LookupParameter(name)
                if p and p.HasValue:
                    v = _param_to_m3(p)
                    nm_low = name.lower()
                    if "cut" in nm_low:
                        cut  += v
                    if "fill" in nm_low and "net" not in nm_low:
                        fill += v
            except:
                pass

        try:
            for p in elem.Parameters:
                try:
                    nm = p.Definition.Name if p.Definition else ""
                except:
                    nm = ""
                nml = (nm or "").lower()
                if ("cut" in nml or "fill" in nml) and "offset" not in nml:
                    v = _param_to_m3(p)
                    if "cut" in nml and v:
                        cut  += v
                    if "fill" in nml and "net" not in nml and v:
                        fill += v
        except:
            pass

    return max(cut, 0.0), max(fill, 0.0)

def _read_cut_fill_from_schedule_cells(doc):
    cut_total = 0.0
    fill_total = 0.0
    try:
        topo_cat_id = DB.Category.GetCategory(
            doc, DB.BuiltInCategory.OST_Topography
        ).Id
    except:
        topo_cat_id = None

    scheds = (
        DB.FilteredElementCollector(doc)
        .OfClass(DB.ViewSchedule)
        .ToElements()
    )

    for vs in scheds:
        try:
            if (
                topo_cat_id
                and vs.Definition
                and vs.Definition.CategoryId
                and vs.Definition.CategoryId.IntegerValue
                != topo_cat_id.IntegerValue
            ):
                continue

            table = vs.GetTableData()
            header = table.GetSectionData(DB.SectionType.Header)
            body   = table.GetSectionData(DB.SectionType.Body)
            if body is None:
                continue

            col_count = body.NumberOfColumns
            header_names = []

            if header and header.NumberOfRows > 0:
                hdr_row = header.NumberOfRows - 1
                for c in range(col_count):
                    header_names.append(
                        (header.GetCellText(hdr_row, c) or "")
                        .strip()
                        .lower()
                    )
            else:
                if body.NumberOfRows == 0:
                    continue
                for _ in range(col_count):
                    header_names.append("")

            cut_cols = [
                i for i, h in enumerate(header_names)
                if ("cut" in h and "net" not in h)
            ]
            fill_cols = [
                i for i, h in enumerate(header_names)
                if ("fill" in h and "net" not in h)
            ]

            if not cut_cols and not fill_cols and vs.Definition:
                try:
                    field_names = [
                        vs.Definition.GetField(i).GetName().lower()
                        for i in range(vs.Definition.GetFieldCount())
                    ]
                    for i, cap in enumerate(field_names):
                        if "cut" in cap and "net" not in cap:
                            cut_cols.append(i)
                        if "fill" in cap and "net" not in cap:
                            fill_cols.append(i)
                except:
                    pass

            if not cut_cols and not fill_cols:
                continue

            for r in range(body.NumberOfRows):
                row_is_total = False
                for c in range(col_count):
                    txt = (body.GetCellText(r, c) or "").strip().lower()
                    if "total" in txt:
                        row_is_total = True
                        break
                if row_is_total:
                    continue

                for c in cut_cols:
                    val = _parse_value_string_to_m3_raw(body.GetCellText(r, c))
                    cut_total += val
                for c in fill_cols:
                    val = _parse_value_string_to_m3_raw(body.GetCellText(r, c))
                    fill_total += val

        except:
            continue

    return cut_total, fill_total

# ------------------------------------------------------------------------------
# Helpers for splitting by Function (Interior / Exterior)
# ------------------------------------------------------------------------------
def _get_cost(o):
    if not o:
        return 0.0
    try:
        cp = o.LookupParameter(PARAM_COST)
        if cp and cp.HasValue:
            return float(cp.AsDouble())
    except:
        pass
    return 0.0

def _clean_comment(name, raw_comment):
    comment = raw_comment or ""
    if comment.strip().lower() == (name or "").strip().lower():
        comment = ""
    if len(comment.replace(" ", "")) < 3:
        comment = ""
    return comment

def _get_function_string(el_type):
    if not el_type:
        return ""
    try:
        func_param = el_type.LookupParameter("Function")
    except:
        func_param = None
    if not (func_param and func_param.HasValue):
        return ""

    try:
        val = func_param.AsString()
        if val:
            return val.strip().lower()
    except:
        pass

    try:
        val = func_param.AsValueString()
        if val:
            return val.strip().lower()
    except:
        pass

    return ""

def _is_external_function(fv_lower):
    if "exterior" in fv_lower:
        return True
    if "external" in fv_lower:
        return True
    if "outside" in fv_lower:
        return True
    return False

def _gather_floors_by_function(doc):
    internal = {}
    external = {}

    floors = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Floors)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    for el in floors:
        try:
            el_type = doc.GetElement(el.GetTypeId()) if el.GetTypeId() else None

            name = None
            if el_type:
                p_name = el_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                if p_name and p_name.HasValue:
                    name = p_name.AsString()
            if not name:
                p_ft = el.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
                if p_ft and p_ft.HasValue:
                    name = p_ft.AsValueString()
            if not name:
                name = getattr(el, "Name", None) or "Floor"

            fv = _get_function_string(el_type)

            rate = _get_cost(el_type) or _get_cost(el)

            qty = 0.0
            unit = "m²"
            a = el.LookupParameter("Area")
            if a and a.HasValue:
                qty = a.AsDouble() * FT2_TO_M2

            cmt = ""
            if el_type:
                tc = el_type.LookupParameter("Type Comments")
                if tc and tc.HasValue:
                    cmt = tc.AsString() or ""
            cmt = _clean_comment(name, cmt)

            bucket = "internal"
            if _is_external_function(fv):
                bucket = "external"

            grouped = internal if bucket == "internal" else external
            if name not in grouped:
                grouped[name] = {
                    "qty": qty,
                    "rate": rate,
                    "unit": unit,
                    "comment": cmt
                }
            else:
                grouped[name]["qty"] += qty
                if grouped[name]["rate"] == 0.0 and rate:
                    grouped[name]["rate"] = rate
                if cmt and not grouped[name].get("comment"):
                    grouped[name]["comment"] = cmt

        except:
            pass

    return internal, external

def _gather_walls_by_function(doc):
    internal = {}
    external = {}

    walls = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Walls)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    for el in walls:
        try:
            el_type = doc.GetElement(el.GetTypeId()) if el.GetTypeId() else None

            name = None
            if el_type:
                p_name = el_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                if p_name and p_name.HasValue:
                    name = p_name.AsString()
            if not name:
                famtype = el.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
                if famtype and famtype.HasValue:
                    name = famtype.AsValueString()
            if not name:
                name = getattr(el, "Name", None) or "Wall"

            fv = _get_function_string(el_type)

            rate = _get_cost(el_type) or _get_cost(el)

            qty = 0.0
            unit = "m²"
            area_param = (
                el.get_Parameter(DB.BuiltInParameter.HOST_AREA_COMPUTED)
                or el.LookupParameter("Area")
            )
            if area_param and area_param.HasValue:
                qty = area_param.AsDouble() * FT2_TO_M2

            cmt = ""
            if el_type:
                tc = el_type.LookupParameter("Type Comments")
                if tc and tc.HasValue:
                    cmt = tc.AsString() or ""
            cmt = _clean_comment(name, cmt)

            bucket = "internal"
            if _is_external_function(fv):
                bucket = "external"

            grouped = internal if bucket == "internal" else external
            if name not in grouped:
                grouped[name] = {
                    "qty": qty,
                    "rate": rate,
                    "unit": unit,
                    "comment": cmt
                }
            else:
                grouped[name]["qty"] += qty
                if grouped[name]["rate"] == 0.0 and rate:
                    grouped[name]["rate"] = rate
                if cmt and not grouped[name].get("comment"):
                    grouped[name]["comment"] = cmt

        except:
            pass

    return internal, external

def _gather_stairs_by_function(doc):
    internal = {}
    external = {}

    stairs = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Stairs)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    for el in stairs:
        try:
            el_type = doc.GetElement(el.GetTypeId()) if el.GetTypeId() else None

            name = None
            if el_type:
                p_name = el_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                if p_name and p_name.HasValue:
                    name = p_name.AsString()
            if not name:
                famtype = el.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
                if famtype and famtype.HasValue:
                    name = famtype.AsValueString()
            if not name:
                name = getattr(el, "Name", None) or "Stair"

            fv = _get_function_string(el_type)

            rate = _get_cost(el_type) or _get_cost(el)

            qty = 1.0
            unit = "No."
            area_param = (
                el.LookupParameter("Actual Tread Surface Area")
                or el.LookupParameter("Tread Surface Area")
                or el.LookupParameter("Area")
            )
            if area_param and area_param.HasValue:
                try:
                    area_val = area_param.AsDouble() * FT2_TO_M2
                    if area_val > 0:
                        qty = area_val
                        unit = "m²"
                except:
                    pass

            cmt = ""
            if el_type:
                tc = el_type.LookupParameter("Type Comments")
                if tc and tc.HasValue:
                    cmt = tc.AsString() or ""
            cmt = _clean_comment(name, cmt)

            bucket = "internal"
            if _is_external_function(fv):
                bucket = "external"

            grouped = internal if bucket == "internal" else external
            if name not in grouped:
                grouped[name] = {
                    "qty": qty,
                    "rate": rate,
                    "unit": unit,
                    "comment": cmt
                }
            else:
                grouped[name]["qty"] += qty
                if grouped[name]["rate"] == 0.0 and rate:
                    grouped[name]["rate"] = rate
                if cmt and not grouped[name].get("comment"):
                    grouped[name]["comment"] = cmt

        except:
            pass

    return internal, external

# ------------------------------------------------------------------------------
# External works collectors (Parking / Planting / Site Works etc.)
# ------------------------------------------------------------------------------
def _collect_elements_by_categories(doc, bic_list, default_unit="No."):
    """
    Group instances from multiple BuiltInCategories by type name.
    Returns { name: {qty, unit, rate, comment} }.
    - qty increments by 1 per instance
    - unit defaults to "No."
    - rate from Cost
    - comment from Type Comments
    """
    if not isinstance(bic_list, (list, tuple)):
        bic_list = [bic_list]

    grouped = {}

    for bic in bic_list:
        if bic is None:
            continue
        try:
            elems = (
                DB.FilteredElementCollector(doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .ToElements()
            )
        except:
            elems = []

        for el in elems:
            try:
                el_type = doc.GetElement(el.GetTypeId()) if el.GetTypeId() else None

                name = None
                if el_type:
                    p_name = el_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                    if p_name and p_name.HasValue:
                        name = p_name.AsString()
                if not name:
                    p_ft = el.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
                    if p_ft and p_ft.HasValue:
                        name = p_ft.AsValueString()
                if not name:
                    name = getattr(el, "Name", None) or (el.Category.Name if el.Category else "Item")

                rate = _get_cost(el_type) or _get_cost(el)

                cmt = ""
                if el_type:
                    tc = el_type.LookupParameter("Type Comments")
                    if tc and tc.HasValue:
                        cmt = tc.AsString() or ""
                cmt = _clean_comment(name, cmt)

                if name not in grouped:
                    grouped[name] = {
                        "qty": 0.0,
                        "rate": rate,
                        "unit": default_unit,
                        "comment": cmt
                    }
                grouped[name]["qty"] += 1.0
                if grouped[name]["rate"] == 0.0 and rate:
                    grouped[name]["rate"] = rate
                if cmt and not grouped[name].get("comment"):
                    grouped[name]["comment"] = cmt

            except:
                pass

    return grouped

def _gather_parking_items(doc):
    """
    Parking-related stuff: bays, bollards, markings, signs, etc.
    We'll include a few likely categories.
    """
    safe_bics = [
        DB.BuiltInCategory.OST_Parking,
        DB.BuiltInCategory.OST_ParkingComponents
        if hasattr(DB.BuiltInCategory, "OST_ParkingComponents") else None,
        DB.BuiltInCategory.OST_Site,
        DB.BuiltInCategory.OST_SpecialityEquipment,
    ]
    return _collect_elements_by_categories(doc, safe_bics, default_unit="No.")

def _gather_planting_items(doc):
    """
    Planting / trees / shrubs.
    """
    bic_list = [DB.BuiltInCategory.OST_Planting]
    return _collect_elements_by_categories(doc, bic_list, default_unit="No.")

def _gather_site_items(doc):
    """
    General site furniture, lighting poles, signs, benches, etc.
    We include:
    - OST_Site
    - OST_SpecialityEquipment
    - OST_LightingFixtures (street lights if modeled as lighting fixtures)
    - OST_GenericModel (catch-all for site furniture)
    """
    safe_bics = [
        DB.BuiltInCategory.OST_Site,
        DB.BuiltInCategory.OST_SpecialityEquipment,
        DB.BuiltInCategory.OST_LightingFixtures,
        DB.BuiltInCategory.OST_GenericModel,
    ]
    return _collect_elements_by_categories(doc, safe_bics, default_unit="No.")

# ------------------------------------------------------------------------------
# Generic writer for grouped dicts
# ------------------------------------------------------------------------------
def _dump_manual_group(cat_name, grouped, ctx):
    """
    Writes a category block (cat_name) using pre-grouped dict items:
      grouped[name] = {qty, rate, unit, comment}
    Updates ctx row counters, subtotals etc.
    """
    if not grouped:
        return

    ws           = ctx["ws"]
    row          = ctx["row"]
    cat_counter  = ctx["cat_counter"]
    cat_subtotal = ctx["cat_subtotals"]

    ws.write(row, 0, str(cat_counter), fmt_section)
    ws.write(row, 1, cat_name.upper(), fmt_section)
    row += 1
    cat_counter += 1
    ctx["order"].append(cat_name)

    if cat_name in CATEGORY_DESCRIPTIONS:
        ws.write(row, 1, CATEGORY_DESCRIPTIONS[cat_name], fmt_description)
        row += 1

    first_item_row = row
    item_idx = 0
    for name, data in grouped.items():
        ws.write(row, 0, _item_label(item_idx), fmt_normal)
        ws.write(row, 1, name, fmt_normal)
        ws.write(row, 2, data["unit"], fmt_normal)
        ws.write(row, 3, round(float(data["qty"]), 2), fmt_normal)
        ws.write(row, 4, round(float(data["rate"]), 2), fmt_money)

        ws.write_formula(
            row, 5,
            "={}*{}".format(
                xl_rowcol_to_cell(row, 3),
                xl_rowcol_to_cell(row, 4)
            ),
            fmt_money
        )
        row += 1
        item_idx += 1

        if data.get("comment"):
            ws.write(row, 1, data["comment"], fmt_italic)
            row += 1

    last_item_row = row - 1
    ws.write(row, 1, cat_name.upper() + " TO COLLECTION", fmt_section)
    ws.write_formula(
        row, 5,
        "=SUM(F{}:F{})".format(first_item_row + 1, last_item_row + 1),
        fmt_money
    )
    cat_subtotal[cat_name.upper()] = xl_rowcol_to_cell(row, 5)
    row += 2

    ctx["row"]         = row
    ctx["cat_counter"] = cat_counter

# ------------------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------------------
skipped = 0

# 0. Gather internal/external groups for Floors, Walls, Stairs
internal_floors, external_floors = _gather_floors_by_function(revit.doc)
internal_walls,  external_walls  = _gather_walls_by_function(revit.doc)
internal_stairs, external_stairs = _gather_stairs_by_function(revit.doc)

_dump_manual_group("Internal Floors", internal_floors, sheets[BILL1_NAME])
_dump_manual_group("External Floors", external_floors, sheets[BILL3_NAME])

_dump_manual_group("Internal Walls", internal_walls,   sheets[BILL1_NAME])
_dump_manual_group("External Walls",  external_walls,   sheets[BILL3_NAME])

_dump_manual_group("Internal Stairs", internal_stairs, sheets[BILL1_NAME])
_dump_manual_group("External Stairs", external_stairs, sheets[BILL3_NAME])

# 1. Process CATEGORY_ORDER (remaining categories)
for cat_name in CATEGORY_ORDER:
    if cat_name in (
        "Internal Floors", "External Floors",
        "Internal Walls",  "External Walls",
        "Internal Stairs", "External Stairs"
    ):
        continue

    bill_name = _bill_for(cat_name)
    ctx = sheets[bill_name]
    ws = ctx["ws"]
    row = ctx["row"]
    cat_counter = ctx["cat_counter"]
    cat_subtotals = ctx["cat_subtotals"]
    bic = CATEGORY_MAP.get(cat_name)
    if not bic:
        continue

    # ----- VIRTUAL: Painting -----
    if bic is VIRTUAL_PAINT:
        grouped = _gather_wall_painting(revit.doc)
        if grouped:
            ws.write(row, 0, str(cat_counter), fmt_section)
            ws.write(row, 1, cat_name.upper(), fmt_section)
            row += 1
            cat_counter += 1
            ctx["order"].append(cat_name)

            if cat_name in CATEGORY_DESCRIPTIONS:
                ws.write(row, 1, CATEGORY_DESCRIPTIONS[cat_name], fmt_description)
                row += 1

            first_item_row = row
            item_idx = 0
            for name, data in grouped.items():
                ws.write(row, 0, _item_label(item_idx), fmt_normal)
                ws.write(row, 1, name, fmt_normal)
                ws.write(row, 2, data["unit"], fmt_normal)
                ws.write(row, 3, round(float(data["qty"]), 2), fmt_normal)
                ws.write(row, 4, round(float(data["rate"]), 2), fmt_money)
                ws.write_formula(
                    row, 5,
                    "={}*{}".format(
                        xl_rowcol_to_cell(row, 3),
                        xl_rowcol_to_cell(row, 4)
                    ),
                    fmt_money
                )
                row += 1
                item_idx += 1

            last_item_row = row - 1
            ws.write(row, 1, cat_name.upper() + " TO COLLECTION", fmt_section)
            if last_item_row >= first_item_row:
                ws.write_formula(
                    row, 5,
                    "=SUM(F{}:F{})".format(first_item_row + 1, last_item_row + 1),
                    fmt_money
                )
            else:
                ws.write(row, 5, 0, fmt_money)

            cat_subtotals[cat_name.upper()] = xl_rowcol_to_cell(row, 5)
            row += 2

        ctx["row"] = row
        ctx["cat_counter"] = cat_counter
        continue

    # ----- SPECIAL: Cut and Fill -----
    if cat_name == "Cut and Fill":
        total_cut_m3  = 0.0
        total_fill_m3 = 0.0

        sc_cut, sc_fill = _read_cut_fill_from_schedule_cells(revit.doc)
        total_cut_m3  += sc_cut
        total_fill_m3 += sc_fill

        if total_cut_m3 < 1e-9 and total_fill_m3 < 1e-9:
            graded_elems = []
            try:
                import Autodesk
                Arch = Autodesk.Revit.DB.Architecture
                if hasattr(Arch, "GradedRegion"):
                    graded_elems = list(
                        DB.FilteredElementCollector(revit.doc)
                        .OfClass(Arch.GradedRegion).ToElements()
                    )
            except Exception:
                graded_elems = []
            for g in graded_elems:
                c, f = _cutfill_from_elem(g)
                total_cut_m3  += c
                total_fill_m3 += f

        if total_cut_m3 < 1e-9 and total_fill_m3 < 1e-9:
            topo_elems = list(
                DB.FilteredElementCollector(revit.doc)
                .OfCategory(DB.BuiltInCategory.OST_Topography)
                .WhereElementIsNotElementType().ToElements()
            )
            for t in topo_elems:
                c, f = _cutfill_from_elem(t)
                total_cut_m3  += c
                total_fill_m3 += f

        if total_cut_m3 < 1e-9 and total_fill_m3 < 1e-9:
            for e in DB.FilteredElementCollector(revit.doc).WhereElementIsNotElementType():
                try:
                    c, f = _cutfill_from_elem(e)
                    if c > 0 or f > 0:
                        total_cut_m3  += c
                        total_fill_m3 += f
                except:
                    pass

        if total_cut_m3 < 1e-9 and total_fill_m3 < 1e-9:
            pad_elems = list(
                DB.FilteredElementCollector(revit.doc)
                .OfCategory(DB.BuiltInCategory.OST_BuildingPad)
                .WhereElementIsNotElementType().ToElements()
            )
            pad_excav_m3 = 0.0
            for p in pad_elems:
                try:
                    v = p.LookupParameter("Volume")
                    if v and v.HasValue:
                        pad_excav_m3 += v.AsDouble() * FT3_TO_M3
                except:
                    pass

        grouped = {}
        if total_cut_m3 > 1e-9:
            grouped["Cut Volume"] = {
                "qty": round(total_cut_m3, 2),
                "rate": 0.0,
                "unit": "m³",
                "comment": ""
            }
        if total_fill_m3 > 1e-9:
            grouped["Fill Volume"] = {
                "qty": round(total_fill_m3, 2),
                "rate": 0.0,
                "unit": "m³",
                "comment": ""
            }
        if total_cut_m3 < 1e-9 and total_fill_m3 < 1e-9:
            if 'pad_excav_m3' in locals() and pad_excav_m3 > 1e-9:
                grouped["Pad Excavation (est.)"] = {
                    "qty": round(pad_excav_m3, 2),
                    "rate": 0.0,
                    "unit": "m³",
                    "comment": "Estimated from Building Pad volumes (no graded region / schedule values)."
                }

        if grouped:
            ws.write(row, 0, str(cat_counter), fmt_section)
            ws.write(row, 1, cat_name.upper(), fmt_section)
            row += 1
            cat_counter += 1
            ctx["order"].append(cat_name)

            if cat_name in CATEGORY_DESCRIPTIONS:
                ws.write(row, 1, CATEGORY_DESCRIPTIONS[cat_name], fmt_description)
                row += 1

            first_item_row = row
            item_idx = 0
            for name, data in grouped.items():
                ws.write(row, 0, _item_label(item_idx), fmt_normal)
                ws.write(row, 1, name, fmt_normal)
                ws.write(row, 2, data["unit"], fmt_normal)
                ws.write(row, 3, data["qty"], fmt_normal)
                ws.write(row, 4, round(float(data["rate"]), 2), fmt_money)
                ws.write_formula(
                    row, 5,
                    "={}*{}".format(
                        xl_rowcol_to_cell(row, 3),
                        xl_rowcol_to_cell(row, 4)
                    ),
                    fmt_money
                )
                row += 1
                item_idx += 1

            last_item_row = row - 1
            ws.write(row, 1, cat_name.upper() + " TO COLLECTION", fmt_section)
            ws.write_formula(
                row, 5,
                "=SUM(F{}:F{})".format(first_item_row + 1, last_item_row + 1),
                fmt_money
            )
            cat_subtotals[cat_name.upper()] = xl_rowcol_to_cell(row, 5)
            row += 2

        ctx["row"] = row
        ctx["cat_counter"] = cat_counter
        continue

    # ----- Default collector for standard Revit categories -----
    if bic is VIRTUAL_EXTERNAL:
        # handled later (external works sections)
        ctx["row"] = row
        ctx["cat_counter"] = cat_counter
        continue

    if isinstance(bic, list):
        elements = []
        for sub in bic:
            elements += (
                DB.FilteredElementCollector(revit.doc)
                .OfCategory(sub)
                .WhereElementIsNotElementType()
                .ToElements()
            )
    else:
        elements = (
            DB.FilteredElementCollector(revit.doc)
            .OfCategory(bic)
            .WhereElementIsNotElementType()
            .ToElements()
        )

    grouped = {}
    for el in elements:
        try:
            el_type = revit.doc.GetElement(el.GetTypeId()) if el.GetTypeId() else None

            name = None
            if el_type:
                p_name = el_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                if p_name and p_name.HasValue:
                    name = p_name.AsString()
            if not name:
                p_ft = el.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
                if p_ft and p_ft.HasValue:
                    name = p_ft.AsValueString()
            if not name:
                name = getattr(el, "Name", None) or (el.Category.Name if el.Category else "Item")

            rate = _get_cost(el_type) or _get_cost(el)

            qty = 1.0
            unit = "No."

            if cat_name in ("Block Work in Walls",):
                prm = el.get_Parameter(DB.BuiltInParameter.HOST_AREA_COMPUTED) or el.LookupParameter("Area")
                qty = prm.AsDouble() * FT2_TO_M2 if (prm and prm.HasValue) else 0.0
                unit = "m²"

            elif cat_name in ("Doors","Windows"):
                qty = 1
                unit = "No."

            elif cat_name in ("Wall and Floor Finishes","Roofs","Ceilings"):
                prm = el.LookupParameter("Area")
                if prm and prm.HasValue:
                    qty = prm.AsDouble() * FT2_TO_M2
                    unit = "m²"

            elif cat_name == "Structural Foundations":
                prm = el.LookupParameter("Volume")
                if prm and prm.HasValue:
                    qty = prm.AsDouble() * FT3_TO_M3
                    unit = "m³"


            elif cat_name == "Structural Framing":

                # Detect material

                mat_prm = el.LookupParameter("Structural Material")

                mat_elem = revit.doc.GetElement(mat_prm.AsElementId()) if mat_prm else None

                mat_name = (mat_elem.Name if mat_elem else "").lower()

                vol_prm = el.LookupParameter("Volume")

                len_prm = el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)

                if "concrete" in mat_name:

                    # Concrete beams / ring beams → volume

                    if vol_prm and vol_prm.HasValue:

                        qty = vol_prm.AsDouble() * FT3_TO_M3

                        unit = "m³"

                    else:

                        qty = 0.0

                        unit = "m³"


                else:

                    # Steel beams / channels → length

                    if len_prm and len_prm.HasValue:

                        qty = len_prm.AsDouble() * FT_TO_M

                        unit = "m"

                    else:

                        qty = 0.0

                        unit = "m"


            elif cat_name == "Structural Rebar":
                # Use TOTAL bar length (accounts for quantity of bars)
                length_param = el.get_Parameter(
                    DB.BuiltInParameter.REBAR_ELEM_TOTAL_LENGTH
                )

                if length_param and length_param.HasValue:
                    qty = length_param.AsDouble() * FT_TO_M
                    unit = "m"
                else:
                    qty = 0.0
                    unit = "m"



            elif cat_name == "Structural Columns":
                mat_prm  = el.LookupParameter("Structural Material")
                mat_elem = revit.doc.GetElement(mat_prm.AsElementId()) if mat_prm else None
                low = (
                    (mat_elem.Name if mat_elem else "") + " " +
                    (getattr(mat_elem,"MaterialClass","") if mat_elem else "")
                ).lower()

                vol_prm = (
                    el.get_Parameter(DB.BuiltInParameter.HOST_VOLUME_COMPUTED)
                    or el.LookupParameter("Volume")
                )
                len_prm = (
                    el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
                    or el.get_Parameter(DB.BuiltInParameter.INSTANCE_LENGTH_PARAM)
                    or el.get_Parameter(DB.BuiltInParameter.COLUMN_HEIGHT)
                    or el.LookupParameter("Length")
                )

                if "concrete" in low:
                    if vol_prm and vol_prm.HasValue:
                        qty  = vol_prm.AsDouble() * FT3_TO_M3
                        unit = "m³"
                    elif len_prm and len_prm.HasValue:
                        qty  = len_prm.AsDouble() * FT_TO_M
                        unit = "m"
                elif ("steel" in low) or ("metal" in low):
                    if len_prm and len_prm.HasValue:
                        qty  = len_prm.AsDouble() * FT_TO_M
                        unit = "m"
                    elif vol_prm and vol_prm.HasValue:
                        qty  = vol_prm.AsDouble() * FT3_TO_M3
                        unit = "m³"
                else:
                    if vol_prm and vol_prm.HasValue and vol_prm.AsDouble()>0:
                        qty  = vol_prm.AsDouble() * FT3_TO_M3
                        unit = "m³"
                    elif len_prm and len_prm.HasValue:
                        qty  = len_prm.AsDouble() * FT_TO_M
                        unit = "m"

            comment = ""
            if el_type:
                tc = el_type.LookupParameter("Type Comments")
                if tc and tc.HasValue:
                    comment = tc.AsString() or ""
            comment = _clean_comment(name, comment)

            if name not in grouped:
                grouped[name] = {
                    "qty": 0.0,
                    "rate": rate,
                    "unit": unit,
                    "comment": comment
                }
            grouped[name]["qty"] += qty
            if grouped[name]["rate"] == 0.0 and rate:
                grouped[name]["rate"] = rate
            if comment and not grouped[name].get("comment"):
                grouped[name]["comment"] = comment

        except:
            skipped += 1

    if grouped:
        ws.write(row, 0, str(cat_counter), fmt_section)
        ws.write(row, 1, cat_name.upper(), fmt_section)
        row += 1
        cat_counter += 1
        ctx["order"].append(cat_name)

        if cat_name in CATEGORY_DESCRIPTIONS:
            ws.write(row, 1, CATEGORY_DESCRIPTIONS[cat_name], fmt_description)
            row += 1

        first_item_row = row
        item_idx = 0
        for name, data in grouped.items():
            ws.write(row, 0, _item_label(item_idx), fmt_normal)
            ws.write(row, 1, name, fmt_normal)
            ws.write(row, 2, data["unit"], fmt_normal)
            ws.write(row, 3, round(float(data["qty"]), 2), fmt_normal)
            ws.write(row, 4, round(float(data["rate"]), 2), fmt_money)
            ws.write_formula(
                row, 5,
                "={}*{}".format(
                    xl_rowcol_to_cell(row, 3),
                    xl_rowcol_to_cell(row, 4)
                ),
                fmt_money
            )
            row += 1
            item_idx += 1

            if data.get("comment"):
                ws.write(row, 1, data["comment"], fmt_italic)
                row += 1

        last_item_row = row - 1
        ws.write(row, 1, cat_name.upper() + " TO COLLECTION", fmt_section)
        if last_item_row >= first_item_row:
            ws.write_formula(
                row, 5,
                "=SUM(F{}:F{})".format(first_item_row + 1, last_item_row + 1),
                fmt_money
            )
        else:
            ws.write(row, 5, 0, fmt_money)

        cat_subtotals[cat_name.upper()] = xl_rowcol_to_cell(row, 5)
        row += 2

    ctx["row"] = row
    ctx["cat_counter"] = cat_counter

# 2. Process EXTERNAL_WORKS_ORDER with real model data for Parking / Planting / Site Works etc.
for ext_cat in EXTERNAL_WORKS_ORDER:
    if ext_cat in ("External Floors", "External Walls", "External Stairs"):
        continue

    bill_name = _bill_for(ext_cat)
    ctx = sheets[bill_name]
    ws = ctx["ws"]
    row = ctx["row"]
    cat_counter = ctx["cat_counter"]
    cat_subtotals = ctx["cat_subtotals"]

    if CATEGORY_MAP.get(ext_cat) is not VIRTUAL_EXTERNAL:
        continue

    if ext_cat == "Parking":
        grouped = _gather_parking_items(revit.doc)
        default_desc = CATEGORY_DESCRIPTIONS.get("Parking", "")
        fallback_label = "Parking works - see site drawings / spec"

    elif ext_cat == "Planting":
        grouped = _gather_planting_items(revit.doc)
        default_desc = CATEGORY_DESCRIPTIONS.get("Planting", "")
        fallback_label = "Planting works - see site drawings / spec"

    elif ext_cat == "Site Works":
        grouped = _gather_site_items(revit.doc)
        default_desc = CATEGORY_DESCRIPTIONS.get("Site Works", "")
        fallback_label = "Site works - see site drawings / spec"

    elif ext_cat in ("Paving", "Drainage", "Fencing"):
        grouped = {}
        default_desc = CATEGORY_DESCRIPTIONS.get(ext_cat, "")
        fallback_label = ext_cat + " works - see site drawings / spec"

    else:
        grouped = {}
        default_desc = CATEGORY_DESCRIPTIONS.get(ext_cat, "")
        fallback_label = ext_cat + " works - see site drawings / spec"

    # fallback placeholder = unit "Item"
    if not grouped:
        grouped = {
            fallback_label: {
                "qty": 1.0,
                "rate": 0.0,
                "unit": "Item",  # <-- keep placeholder as Item
                "comment": ""
            }
        }

    ws.write(row, 0, str(cat_counter), fmt_section)
    ws.write(row, 1, ext_cat.upper(), fmt_section)
    row += 1
    cat_counter += 1
    ctx["order"].append(ext_cat)

    if default_desc:
        ws.write(row, 1, default_desc, fmt_description)
        row += 1

    first_item_row = row
    item_idx = 0
    for name, data in grouped.items():
        ws.write(row, 0, _item_label(item_idx), fmt_normal)
        ws.write(row, 1, name, fmt_normal)
        ws.write(row, 2, data["unit"], fmt_normal)
        ws.write(row, 3, round(float(data["qty"]), 2), fmt_normal)
        ws.write(row, 4, round(float(data["rate"]), 2), fmt_money)
        ws.write_formula(
            row, 5,
            "={}*{}".format(
                xl_rowcol_to_cell(row, 3),
                xl_rowcol_to_cell(row, 4)
            ),
            fmt_money
        )
        row += 1
        item_idx += 1

        if data.get("comment"):
            ws.write(row, 1, data["comment"], fmt_italic)
            row += 1

    last_item_row = row - 1
    ws.write(row, 1, ext_cat.upper() + " TO COLLECTION", fmt_section)
    ws.write_formula(
        row, 5,
        "=SUM(F{}:F{})".format(first_item_row + 1, last_item_row + 1),
        fmt_money
    )
    cat_subtotals[ext_cat.upper()] = xl_rowcol_to_cell(row, 5)
    row += 2

    ctx["row"] = row
    ctx["cat_counter"] = cat_counter

# ------------------------------------------------------------------------------
# Finalize bills & GENERAL SUMMARY
# ------------------------------------------------------------------------------
ORDERED_BILLS = [BILL1_NAME, BILL2_NAME, BILL3_NAME]
bill_grand_refs = []

for bill_name in ORDERED_BILLS:
    ctx = sheets[bill_name]
    ws = ctx["ws"]
    grand_addr, _ = finalize_bill_sheet(
        ws,
        ctx["row"],
        ctx["order"],
        ctx["cat_subtotals"]
    )
    bill_grand_refs.append(_sheet_ref(bill_name, grand_addr))

summary_ws = wb.add_worksheet(SUMMARY_NAME)
_set_portrait(summary_ws)
summary_ws.set_tab_color(TAB_COLORS["SUMMARY"])

summary_ws.set_column(0, 0, 6)
summary_ws.set_column(1, 1, 60)
summary_ws.set_column(2, 2, 4)
summary_ws.set_column(3, 3, 18)

summary_ws.merge_range(0, 0, 0, 3, "GENERAL SUMMARY", fmt_center)

summary_ws.write(1, 0, "ITEM", fmt_header)
summary_ws.write(1, 1, "DESCRIPTION", fmt_header)
summary_ws.write(1, 2, "", fmt_header)
summary_ws.write(1, 3, "AMOUNT (EUR)", fmt_header)

row = 2
summary_ws.merge_range(
    row, 1,
    row, 3,
    _get_project_title().upper(),
    fmt_bold
)
row += 2

CURRENCY_SYM = "€"

for idx, (bill_name, ref) in enumerate(
    zip(ORDERED_BILLS, bill_grand_refs),
    start=1
):
    if " - " in bill_name:
        label_tail = bill_name.split(" - ", 1)[-1].upper()
    else:
        label_tail = bill_name.upper()

    summary_ws.write(row, 1, "BILL No. {}: {}".format(idx, label_tail), fmt_text)
    summary_ws.write(row, 2, CURRENCY_SYM, fmt_text)
    summary_ws.write_formula(row, 3, "=" + ref, fmt_money_right)
    row += 1

sub1_row = row
summary_ws.write_blank(row, 0, None, fmt_text)
summary_ws.write(row, 1, "Sub total 1", fmt_bold)
summary_ws.write(row, 2, CURRENCY_SYM, fmt_bold)

if bill_grand_refs:
    summary_ws.write_formula(
        row, 3,
        "=SUM({})".format(",".join(bill_grand_refs)),
        fmt_money_right
    )
else:
    summary_ws.write(row, 3, 0, fmt_money_right)
row += 2

disc_text = (
    "Should the Contractor desire to make any discount on the above total, "
    "it is to be made here and the amount will be treated as a percentage of "
    "the total as above. The rates inserted by the contractor against the "
    "items throughout this tender will be adjusted accordingly by this "
    "percentage during project execution"
)

disc_top = row
disc_bottom = row + 5
summary_ws.merge_range(
    disc_top, 1,
    disc_bottom, 1,
    disc_text,
    fmt_wrap
)
summary_ws.write(disc_top, 2, "%", fmt_center)
summary_ws.write(disc_top + 1, 2, 0, fmt_percent)
discount_cell = xl_rowcol_to_cell(disc_top + 1, 2)

row = disc_bottom + 1

sub2_row = row
summary_ws.write_blank(row, 0, None, fmt_text)
summary_ws.write(row, 1, "Sub total 2", fmt_bold)
summary_ws.write(row, 2, CURRENCY_SYM, fmt_bold)
summary_ws.write_formula(
    row, 3,
    "={}*(1-{})".format(
        xl_rowcol_to_cell(sub1_row, 3),
        discount_cell
    ),
    fmt_money_right
)
row += 1

CONTINGENCY_RATE = 0.05
summary_ws.write(
    row, 1,
    "Allow for contingencies @ {}%".format(int(CONTINGENCY_RATE * 100)),
    fmt_text
)
summary_ws.write_blank(row, 2, None, fmt_text)
summary_ws.write_formula(
    row, 3,
    "={}*{}".format(
        xl_rowcol_to_cell(sub2_row, 3),
        CONTINGENCY_RATE
    ),
    fmt_money_right
)
contingency_row = row
row += 1

sub3_row = row
summary_ws.write_blank(row, 0, None, fmt_text)
summary_ws.write(row, 1, "Sub total 3", fmt_bold)
summary_ws.write(row, 2, CURRENCY_SYM, fmt_bold)
summary_ws.write_formula(
    row, 3,
    "={}+{}".format(
        xl_rowcol_to_cell(sub2_row, 3),
        xl_rowcol_to_cell(contingency_row, 3)
    ),
    fmt_money_right
)
row += 1

summary_ws.write(
    row, 1,
    "Add VAT OR TOT, whichever is applicable",
    fmt_text
)
summary_ws.write(row, 2, "", fmt_text)
summary_ws.write(row, 3, "Inclusive", fmt_text)
row += 1

summary_ws.write(
    row, 1,
    "GRAND TOTAL CARRIED TO FORM OF TENDER",
    fmt_bold
)
summary_ws.write(row, 2, CURRENCY_SYM, fmt_bold)
summary_ws.write_formula(
    row, 3,
    "={}".format(xl_rowcol_to_cell(sub3_row, 3)),
    fmt_money_right
)
row += 1

FIRST_PAGE_LAST_ROW = 47
SIG_BLOCK_HEIGHT = 4
sig_top_row_1based = FIRST_PAGE_LAST_ROW - SIG_BLOCK_HEIGHT + 1
sig_top_row_0based = sig_top_row_1based - 1

while row < sig_top_row_0based:
    summary_ws.write_blank(row, 0, None, fmt_noborder)
    summary_ws.write_blank(row, 1, None, fmt_noborder)
    summary_ws.write_blank(row, 2, None, fmt_noborder)
    summary_ws.write_blank(row, 3, None, fmt_noborder)
    row += 1

summary_ws.write(
    row, 1,
    "Signature of Contractor .................................................................",
    fmt_text
)
row += 1
summary_ws.write(
    row, 1,
    "Name of Firm: ..............................................................................",
    fmt_text
)
row += 1
summary_ws.write(
    row, 1,
    "Address: ...................................................................................",
    fmt_text
)
row += 1
summary_ws.write(
    row, 1,
    "Date: ......................................................................................",
    fmt_text
)
row += 1

summary_ws.set_h_pagebreaks([FIRST_PAGE_LAST_ROW])

# ------------------------------------------------------------------------------
# Close and notify
# ------------------------------------------------------------------------------
wb.close()
MessageBox.Show(
    "BOQ export (multi-sheet) complete!\nSaved to Desktop:\n{}\nSkipped: {}".format(
        xlsx_path, skipped
    ),
    "✅ XLSX Export"
)
