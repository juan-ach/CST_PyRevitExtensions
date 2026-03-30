# -*- coding: utf-8 -*-

__title__ = "Check Presto"
__author__ = "Juan Achenbach & OpenAI"
__version__ = "Version: 2.0"
__doc__ = """Version: 2.0
_____________________________________________________________________
Description:

Replicates the Dynamo graph "antes de mandar a presto_v2.dyn" in pyRevit.

Implemented blocks:
- room cleanup and chamber volume transfer
- comments / line measurement fields
- PRESTO chapter assignment
- Codigo_Presto assignment
- duct / fitting gross areas
- panel gross areas
- gross pipe lengths
- refrigerant labeling
- wall sweep (zocalo) recreation
_____________________________________________________________________
Author: Juan Manuel Achenbach Anguita & OpenAI"""

import clr
from collections import defaultdict

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import (  # noqa: E401
    BuiltInCategory,
    BuiltInParameter,
    Element,
    ElementId,
    FilteredElementCollector,
    GeometryInstance,
    Options,
    Solid,
    StorageType,
    Transaction,
    TransactionGroup,
    WallSide,
    WallSweep,
    WallSweepInfo,
    WallSweepType,
)
from pyrevit import DB, revit, script

import System
from System.Windows.Forms import Form, Label, Timer
import System.Drawing


doc = revit.doc
output = script.get_output()

FT_TO_M = 0.3048
M_TO_FT = 1.0 / FT_TO_M
GROSS_FACTOR = 1.05
PIPE_WASTE_ADD_M = 4.0
PIPE_BAR_LENGTH_MM = 4000.0
PIPE_BAR_LENGTH_M = 4.0
PIPE_DIAMETER_TOL_MM = 0.5

PARAM_PARTIDAS = u"Partidas_PRESTO"
PARAM_CODIGO_PRESTO = u"Codigo_Presto"
PARAM_COMENTARIOS = u"Comentarios"
PARAM_COMENTARIOS2 = u"Comentarios2"
PARAM_UBICACION = u"ubicación"
PARAM_VOL_CAMARA = u"Vol.Cámara"
PARAM_SUP_BRUTA_PANEL = u"sup.bruta.panel"
PARAM_LONG_BRUTA_TUB = u"long.bruta.tub"
PARAM_LEE_REFRIGERANTE = u"Lee_Refrigerante"
PARAM_DUCT_FITTING_AREA = u"Duct Fitting Area"
PARAM_DUCT_CONNECTION_AREA = u"Duct Connection Area"
PARAM_CODIGO_MONTAJE = u"Código de montaje"

ROOM_NAME_PARAM_NAMES = [u"Nombre", u"Name"]

AUTO_CO2_ABBREVIATIONS = [
    u"A1-AUTO",
    u"A2-AUTO",
    u"A1+AUTO",
    u"A2+AUTO",
    u"L1-AUTO",
    u"L2-AUTO",
    u"L1+AUTO",
    u"L2+AUTO",
]

CENTRAL_PLUS_CO2_TYPES = [
    u"A1+",
    u"A2+",
    u"A3+",
    u"L1",
    u"L2",
    u"L3",
    u"L4",
    u"L5",
]

CENTRAL_MINUS_CO2_TYPES = [u"A1-", u"A2-"]

DRC_CO2_TYPE_CONTAINS = [
    u"DRC Compensación CO2",
    u"DRC Descarga CO2",
    u"DRC Retorno de Líquido CO2",
    u"Desrecalen. IDA A",
    u"Desrecalen. RETORNO A",
]

SAFETY_VALVE_TYPE_CONTAINS = [
    u"Conducción V.S. ACN_A",
    u"Conducción V.S. CN_A",
    u"Conducción V.S. CP_A",
]

ALL_NON_CO2_SYSTEM_TYPES = [
    u"DRC Compensación",
    u"DRC Descarga",
    u"DRC Retorno de Líquido",
    u"L1+_ASPIRACIÓN",
    u"L1+_LÍQUIDO",
    u"L1-_ASPIRACIÓN",
    u"L1-_LÍQUIDO",
    u"L2+_ASPIRACIÓN",
    u"L2+_LÍQUIDO",
    u"L2-_ASPIRACIÓN",
    u"L2-_LÍQUIDO",
    u"L3+_ASPIRACIÓN",
    u"L3+_LÍQUIDO",
    u"L4+_ASPIRACIÓN",
    u"L4+_LÍQUIDO",
    u"L5+ ASPIRACIÓN",
    u"L5+ LÍQUIDO",
    u"L+1_AUTONOMO_ASPIRACIÓN",
    u"L+1_AUTONOMO_LÍQUIDO",
    u"L-1_AUTONOMO_ASPIRACIÓN",
    u"L-1_AUTONOMO_LÍQUIDO",
    u"L+2_AUTONOMO_ASPIRACIÓN",
    u"L+2_AUTONOMO_LÍQUIDO",
    u"L-2_AUTONOMO_ASPIRACIÓN",
    u"L-2_AUTONOMO_LÍQUIDO",
]

DUCT_TYPE_CODE_BY_KEYWORD = [
    (u"condensador", u"COND.EXT.COND"),
    (u"turbina", u"COND.EXT.TURB"),
    (u"gascooler", u"COND.EXT.GASCOOLER"),
]

PRESTO_MM_MAP_STANDARD = [
    (9.525, u"02.TI.CU.38"),
    (12.7, u"03.TI.CU.12"),
    (15.875, u"04.TI.CU.58"),
    (19.05, u"05.TI.CU.34"),
    (22.225, u"06.TI.CU.78"),
    (28.575, u"07.TI.CU.118"),
    (34.925, u"08.TI.CU.138"),
    (41.275, u"09.TI.CU.158"),
    (53.975, u"10.TI.CU.218"),
    (66.675, u"11.TI.CU.258"),
]

PRESTO_MM_MAP_AUTONOMO_120 = [
    (9.525, u"23.TI.CU.38-K65-120B"),
    (12.7, u"24.TI.CU.12-K65-120B"),
    (15.875, u"25.TI.CU.58-K65-120B"),
    (19.05, u"26.TI.CU.34-K65-120B"),
    (22.225, u"27.TI.CU.78-K65-120B"),
    (28.575, u"28.TI.CU.118-K65-120B"),
    (34.925, u"29.TI.CU.138-K65-120B"),
    (41.275, u"30.TI.CU.158-K65-120B"),
    (53.975, u"31.TI.CU.218-K65-120B"),
]

PRESTO_MM_MAP_SERVICIOS_120 = [
    (9.525, u"02.TI.CU.38"),
    (12.7, u"24.TI.CU.12-K65-120B"),
    (15.875, u"25.TI.CU.58-K65-120B"),
    (19.05, u"26.TI.CU.34-K65-120B"),
    (22.225, u"27.TI.CU.78-K65-120B"),
    (28.575, u"28.TI.CU.118-K65-120B"),
    (34.925, u"29.TI.CU.138-K65-120B"),
    (41.275, u"30.TI.CU.158-K65-120B"),
    (53.975, u"31.TI.CU.218-K65-120B"),
]

PRESTO_MM_MAP_DRC_130 = [
    (9.525, u"32.TI.CU.38-K65-130B"),
    (12.7, u"33.TI.CU.12-K65-130B"),
    (15.875, u"34.TI.CU.58-K65-130B"),
    (19.05, u"35.TI.CU.34-K65-130B"),
    (22.225, u"35.TI.CU.78-K65-130B"),
    (28.575, u"36.TI.CU.118-K65-130B"),
    (34.925, u"37.TI.CU.138-K65-130B"),
    (41.275, u"38.TI.CU.158-K65-130B"),
    (53.975, u"39.TI.CU.218-K65-130B"),
]

INSULATION_CODE_BY_TYPE_NAME = {
    u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 1/4 - 19mm": u"40.AI.14.19mm",
    u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 3/8 - 19mm": u"41.AI.38.19mm",
    u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 1/2 - 19mm": u"42.AI.12.19mm",
    u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 5/8 - 19mm": u"43.AI.58.19mm",
    u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 3/4 - 19mm": u"44.AI.34.19mm",
    u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 7/8 - 19mm": u"45.AI.78.19mm",
    u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 1 1/8 - 19mm": u"46.AI.118.19mm",
    u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 1 3/8 - 19mm": u"47.AI.138.19mm",
    u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 1 5/8 - 19mm": u"48.AI.158.19mm",
    u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 2 1/8 - 19mm": u"49.AI.218.19mm",
    u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 2 5/8 - 19mm": u"50.AI.258.19mm",
    u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 3 1/8 - 19mm": u"51.AI.318.19mm",
    u"AISLAMIENTO INSTAL. TUBERÍA COBRE 1/4 - 32mm": u"52.AI.14.32mm",
    u"AISLAMIENTO INSTAL. TUBERÍA COBRE 3/8 - 32mm": u"53.AI.38.32mm",
    u"AISLAMIENTO INSTAL. TUBERÍA COBRE 1/2 - 32mm": u"54.AI.12.32mm",
    u"AISLAMIENTO INSTAL. TUBERÍA COBRE 5/8 - 32mm": u"55.AI.58.32mm",
    u"AISLAMIENTO INSTAL. TUBERÍA COBRE 3/4 - 32mm": u"56.AI.34.32mm",
    u"AISLAMIENTO INSTAL. TUBERÍA COBRE 7/8 - 32mm": u"57.AI.78.32mm",
    u"AISLAMIENTO INSTAL. TUBERÍA COBRE 1 1/8 - 32mm": u"58.AI.118.32mm",
    u"AISLAMIENTO INSTAL. TUBERÍA COBRE 1 3/8 - 32mm": u"59.AI.138.32mm",
    u"AISLAMIENTO INSTAL. TUBERÍA COBRE 1 5/8 - 32mm": u"60.AI.158.32mm",
    u"AISLAMIENTO INSTAL. TUBERÍA COBRE 2 1/8 - 32mm": u"61.AI.218.32mm",
    u"AISLAMIENTO INSTAL. TUBERÍA COBRE 2 5/8 - 32mm": u"62.AI.258.32mm",
    u"AISLAMIENTO INSTAL. TUBERÍA COBRE 3 1/8 - 32mm": u"63.AI.318.32mm",
}

PIPE_TYPE_RULES = [
    {"name": u"no_co2", "filter_mode": "equals", "filter_field": "system_type_name", "terms": ALL_NON_CO2_SYSTEM_TYPES, "type_match": "contains", "target_type": u"Cu Standar", "presto_map": PRESTO_MM_MAP_STANDARD},
    {"name": u"safety_valve", "filter_mode": "contains", "filter_field": "system_type_name", "terms": SAFETY_VALVE_TYPE_CONTAINS, "type_match": "exact", "target_type": u"Cu Standar", "presto_map": PRESTO_MM_MAP_STANDARD},
    {"name": u"autonomo_co2", "filter_mode": "contains", "filter_field": "system_abbreviation", "terms": AUTO_CO2_ABBREVIATIONS, "type_match": "exact", "target_type": u"Cu_K65 120 bar +", "presto_map": PRESTO_MM_MAP_AUTONOMO_120},
    {"name": u"central_plus_co2", "filter_mode": "equals", "filter_field": "system_type_name", "terms": CENTRAL_PLUS_CO2_TYPES, "type_match": "exact", "target_type": u"Cu_K65 120 bar +", "presto_map": PRESTO_MM_MAP_SERVICIOS_120},
    {"name": u"central_minus_co2", "filter_mode": "equals", "filter_field": "system_type_name", "terms": CENTRAL_MINUS_CO2_TYPES, "type_match": "exact", "target_type": u"Cu_K65 120 bar -", "presto_map": PRESTO_MM_MAP_SERVICIOS_120},
    {"name": u"drc_co2", "filter_mode": "contains", "filter_field": "system_type_name", "terms": DRC_CO2_TYPE_CONTAINS, "type_match": "exact", "target_type": u"Cu_K65 130 bar", "presto_map": PRESTO_MM_MAP_DRC_130},
]


class AutoClosePopup(Form):
    def __init__(self, message, duration_ms=3000):
        self.Text = "Info"
        self.Width = 430
        self.Height = 170
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen

        label = Label()
        label.Text = message
        label.Dock = System.Windows.Forms.DockStyle.Fill
        label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        label.Font = System.Drawing.Font("Arial", 10)
        self.Controls.Add(label)

        timer = Timer()
        timer.Interval = duration_ms
        timer.Tick += self.close_popup
        timer.Start()

    def close_popup(self, sender, args):
        self.Close()


def as_list(value):
    if value is None:
        return []
    if isinstance(value, (list, tuple, set)):
        return list(value)
    return [value]


def ft_to_m(value_ft):
    return value_ft * FT_TO_M if value_ft is not None else None


def m_to_ft(value_m):
    return value_m * M_TO_FT if value_m is not None else None


def ft_to_mm(value_ft):
    return value_ft * 304.8 if value_ft is not None else None


def safe_text(value):
    if value is None:
        return u""
    try:
        return u"{}".format(value)
    except Exception:
        return str(value)


def unique_elements(elements):
    seen = set()
    result = []
    for elem in elements:
        if elem is None:
            continue
        elem_id = elem.Id.IntegerValue
        if elem_id in seen:
            continue
        seen.add(elem_id)
        result.append(elem)
    return result


def get_parameter(elem, names):
    if elem is None:
        return None
    for name in as_list(names):
        try:
            param = elem.LookupParameter(name)
        except Exception:
            param = None
        if param:
            return param
    return None


def get_parameter_value(param):
    if param is None:
        return None
    if param.StorageType == StorageType.String:
        return param.AsString() or param.AsValueString() or u""
    if param.StorageType == StorageType.Double:
        return param.AsDouble()
    if param.StorageType == StorageType.Integer:
        return param.AsInteger()
    if param.StorageType == StorageType.ElementId:
        elem_id = param.AsElementId()
        if elem_id and elem_id != ElementId.InvalidElementId and elem_id.IntegerValue >= 0:
            return doc.GetElement(elem_id)
        return None
    return param.AsValueString()


def get_param_text(elem, names, default=u""):
    param = get_parameter(elem, names)
    if not param:
        return default
    value = get_parameter_value(param)
    if isinstance(value, Element):
        value = get_element_name(value)
    if value is None:
        return default
    if isinstance(value, (int, float)):
        text = param.AsValueString()
        return text if text else safe_text(value)
    return safe_text(value)


def get_param_double(elem, names, default=None):
    param = get_parameter(elem, names)
    if not param:
        return default
    if param.StorageType == StorageType.Double:
        return param.AsDouble()
    value = get_parameter_value(param)
    if isinstance(value, (int, float)):
        return float(value)
    return default


def get_param_element(elem, names):
    param = get_parameter(elem, names)
    if not param or param.StorageType != StorageType.ElementId:
        return None
    elem_id = param.AsElementId()
    if not elem_id or elem_id == ElementId.InvalidElementId or elem_id.IntegerValue < 0:
        return None
    return doc.GetElement(elem_id)


def get_element_name(elem):
    if elem is None:
        return u""
    for candidate in [u"Nombre de tipo", u"Nombre", u"Name"]:
        text = get_param_text(elem, [candidate], default=u"")
        if text:
            return text
    try:
        if elem.Name:
            return safe_text(elem.Name)
    except Exception:
        pass
    try:
        name_param = elem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if name_param and name_param.AsString():
            return name_param.AsString()
    except Exception:
        pass
    try:
        sym_param = elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if sym_param and sym_param.AsString():
            return sym_param.AsString()
    except Exception:
        pass
    return u""


def get_element_type(elem):
    if elem is None:
        return None
    try:
        type_id = elem.GetTypeId()
    except Exception:
        type_id = ElementId.InvalidElementId
    if not type_id or type_id == ElementId.InvalidElementId:
        return None
    return doc.GetElement(type_id)


def set_param_safe(elem, param_name, value):
    param = get_parameter(elem, [param_name])
    if not param or param.IsReadOnly:
        return False
    try:
        if param.StorageType == StorageType.String:
            param.Set(safe_text(value))
            return True
        if param.StorageType == StorageType.Double:
            if value is None:
                return False
            param.Set(float(value))
            return True
        if param.StorageType == StorageType.Integer:
            param.Set(int(round(float(value))))
            return True
        if param.StorageType == StorageType.ElementId:
            if isinstance(value, ElementId):
                param.Set(value)
                return True
            if isinstance(value, Element):
                param.Set(value.Id)
                return True
    except Exception:
        return False
    return False


def collect_elements(built_in_category):
    return list(
        FilteredElementCollector(doc)
        .OfCategory(built_in_category)
        .WhereElementIsNotElementType()
        .ToElements()
    )


def collect_pipe_insulations():
    return list(
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_PipeInsulations)
        .WhereElementIsNotElementType()
        .ToElements()
    )


def collect_pipe_types():
    return list(
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_PipeCurves)
        .WhereElementIsElementType()
        .ToElements()
    )


def find_pipe_type_by_exact_name(type_name):
    for pipe_type in collect_pipe_types():
        if get_element_name(pipe_type) == type_name:
            return pipe_type
    return None


def find_pipe_type_by_contains(text):
    text = safe_text(text)
    for pipe_type in collect_pipe_types():
        if text in get_element_name(pipe_type):
            return pipe_type
    return None


def run_block(name, action, errors):
    tx = Transaction(doc, name)
    try:
        tx.Start()
        result = action()
        tx.Commit()
        return result
    except Exception as err:
        if tx.HasStarted():
            tx.RollBack()
        errors.append(u"{}: {}".format(name, safe_text(err)))
        return None


def get_pipe_length_ft(pipe):
    try:
        param = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
        if param:
            return param.AsDouble()
    except Exception:
        pass
    return get_param_double(pipe, [u"Longitud"])


def get_pipe_diameter_ft(pipe):
    try:
        return pipe.Diameter
    except Exception:
        return get_param_double(pipe, [u"Diámetro", u"Diametro"])


def get_wall_length_ft(wall):
    try:
        param = wall.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
        if param:
            return param.AsDouble()
    except Exception:
        pass
    return get_param_double(wall, [u"Longitud"])


def get_wall_unconnected_height_ft(wall):
    try:
        param = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
        if param:
            return param.AsDouble()
    except Exception:
        pass
    return get_param_double(wall, [u"Altura desconectada"])


def get_floor_area_sqft(floor):
    try:
        param = floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED)
        if param:
            return param.AsDouble()
    except Exception:
        pass
    return get_param_double(floor, [u"Área", u"Area"])


def get_room_perimeter_ft(room):
    try:
        param = room.get_Parameter(BuiltInParameter.ROOM_PERIMETER)
        if param:
            return param.AsDouble()
    except Exception:
        pass
    return get_param_double(room, [u"Perímetro", u"Perimetro"])


def get_room_volume_cuft(room):
    try:
        param = room.get_Parameter(BuiltInParameter.ROOM_VOLUME)
        if param:
            return param.AsDouble()
    except Exception:
        pass
    return get_param_double(room, [u"Volumen"])


def get_pipe_system_type_element(pipe):
    try:
        param = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)
        if param and param.StorageType == StorageType.ElementId:
            elem_id = param.AsElementId()
            if elem_id and elem_id != ElementId.InvalidElementId and elem_id.IntegerValue >= 0:
                return doc.GetElement(elem_id)
    except Exception:
        pass
    return get_param_element(pipe, [u"Tipo de sistema"])


def get_pipe_system_type_name(pipe):
    system_type = get_pipe_system_type_element(pipe)
    if system_type:
        name = get_param_text(system_type, [u"Nombre de tipo"], default=u"")
        if name:
            return name
        name = get_element_name(system_type)
        if name:
            return name
    return get_param_text(pipe, [u"Tipo de sistema"], default=u"")


def get_pipe_system_abbreviation(pipe):
    text = get_param_text(pipe, [u"Abreviatura de sistema"], default=u"")
    if text:
        return text
    system_type = get_pipe_system_type_element(pipe)
    if system_type:
        return get_param_text(system_type, [u"Abreviatura de sistema"], default=u"")
    return u""


def get_pipe_system_classification(pipe):
    system_type = get_pipe_system_type_element(pipe)
    if not system_type:
        return u""
    return get_param_text(system_type, [u"Clasificación de sistema"], default=u"")


def get_pipe_fluid_type_name(pipe):
    system_type = get_pipe_system_type_element(pipe)
    if not system_type:
        return u""
    fluid_type = get_param_element(system_type, [u"Tipo de fluido"])
    if fluid_type:
        name = get_param_text(fluid_type, [u"Nombre de tipo"], default=u"")
        return name if name else get_element_name(fluid_type)
    return u""


def contains_any(text, terms):
    text = safe_text(text)
    return any(term in text for term in terms)


def equals_any(text, terms):
    text = safe_text(text)
    return any(term == text for term in terms)


def find_mm_code(diameter_mm, mm_map):
    if diameter_mm is None:
        return None
    best_code = None
    best_diff = None
    for target_mm, code in mm_map:
        diff = abs(diameter_mm - target_mm)
        if best_diff is None or diff < best_diff:
            best_diff = diff
            best_code = code
    if best_diff is not None and best_diff <= PIPE_DIAMETER_TOL_MM:
        return best_code
    return None


def get_total_face_area_sqft(element):
    options = Options()
    options.DetailLevel = DB.ViewDetailLevel.Fine
    options.IncludeNonVisibleObjects = True

    def iter_solids(geometry):
        if geometry is None:
            return
        for geom_obj in geometry:
            if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
                yield geom_obj
            elif isinstance(geom_obj, GeometryInstance):
                for nested in iter_solids(geom_obj.GetInstanceGeometry()):
                    yield nested

    total = 0.0
    geometry = element.get_Geometry(options)
    for solid in iter_solids(geometry):
        for face in solid.Faces:
            total += face.Area
    return total


def delete_elements_by_ids(element_ids):
    ids = [elem_id for elem_id in element_ids if elem_id and elem_id != ElementId.InvalidElementId]
    if not ids:
        return 0
    deleted = doc.Delete(System.Collections.Generic.List[ElementId](ids))
    return len(deleted)


def get_hosted_pipe_insulations_by_pipe():
    result = defaultdict(list)
    for insulation in collect_pipe_insulations():
        host_id = getattr(insulation, "HostElementId", None)
        if not host_id or host_id == ElementId.InvalidElementId:
            continue
        result[host_id.IntegerValue].append(insulation)
    return result


def get_insulation_type_name(insulation):
    return get_element_name(get_element_type(insulation))


def get_duct_type_name(duct):
    return get_element_name(get_element_type(duct))


def get_revit_round(value):
    return float(System.Math.Round(float(value)))


def get_revit_round_digits(value, digits):
    return float(System.Math.Round(float(value), int(digits)))


def get_revit_ceiling(value):
    return float(System.Math.Ceiling(float(value)))


def delete_orphan_rooms(summary):
    rooms = collect_elements(BuiltInCategory.OST_Rooms)
    orphan_ids = []
    valid_rooms = []
    for room in rooms:
        perimeter = get_room_perimeter_ft(room)
        if perimeter is None or perimeter <= 0:
            orphan_ids.append(room.Id)
        else:
            valid_rooms.append(room)
    summary["rooms_deleted"] += delete_elements_by_ids(orphan_ids)
    summary["rooms_valid"] = len(valid_rooms)
    return valid_rooms


def build_room_volume_map(valid_rooms):
    room_map = {}
    for room in valid_rooms:
        room_name = get_param_text(room, ROOM_NAME_PARAM_NAMES, default=u"")
        if not room_name:
            room_name = get_element_name(room)
        if room_name:
            room_map[room_name] = get_room_volume_cuft(room)
    return room_map


def transfer_room_volume_to_equipment(summary, room_volume_map):
    changed = 0
    for equipment in collect_elements(BuiltInCategory.OST_MechanicalEquipment):
        room_name = get_param_text(equipment, [PARAM_UBICACION], default=u"")
        volume_cuft = room_volume_map.get(room_name)
        if volume_cuft is None:
            continue
        if set_param_safe(equipment, PARAM_VOL_CAMARA, volume_cuft):
            changed += 1
    summary["vol_camara"] += changed


def set_comments_for_electrical_devices(summary):
    electrical = collect_elements(BuiltInCategory.OST_ElectricalFixtures)
    if not electrical:
        try:
            electrical = collect_elements(BuiltInCategory.OST_ElectricalEquipment)
        except Exception:
            electrical = []
    changed = 0
    for elem in electrical:
        ubicacion = get_param_text(elem, [PARAM_UBICACION], default=u"")
        comentarios2 = get_param_text(elem, [PARAM_COMENTARIOS2], default=u"")
        value = u"{}{}{}".format(ubicacion, u"_", comentarios2)
        if set_param_safe(elem, PARAM_COMENTARIOS, value):
            changed += 1
    summary["comentarios_electrico"] += changed


def set_comments_from_location(summary, elements, summary_key):
    changed = 0
    for elem in elements:
        ubicacion = get_param_text(elem, [PARAM_UBICACION], default=u"")
        if set_param_safe(elem, PARAM_COMENTARIOS, ubicacion):
            changed += 1
    summary[summary_key] += changed


def set_door_comments(summary):
    changed = 0
    for door in collect_elements(BuiltInCategory.OST_Doors):
        ubicacion = get_param_text(door, [PARAM_UBICACION], default=u"")
        if not ubicacion:
            continue
        comentarios2 = get_param_text(door, [PARAM_COMENTARIOS2], default=u"")
        value = u"{}{}{}".format(ubicacion, u"_", comentarios2)
        if set_param_safe(door, PARAM_COMENTARIOS, value):
            changed += 1
    summary["comentarios_puertas"] += changed


def assign_partidas_by_category(summary):
    changed = 0
    category_map = [
        (BuiltInCategory.OST_PipeCurves, u"04"),
        (BuiltInCategory.OST_PipeInsulations, u"04"),
        (BuiltInCategory.OST_Walls, u"09"),
        (BuiltInCategory.OST_Floors, u"09"),
        (BuiltInCategory.OST_DuctCurves, u"01.09"),
        (BuiltInCategory.OST_DuctTerminal, u"01.09"),
        (BuiltInCategory.OST_DuctFitting, u"01.09"),
        (BuiltInCategory.OST_DuctAccessory, u"01.09"),
    ]
    for category, value in category_map:
        for elem in collect_elements(category):
            if set_param_safe(elem, PARAM_PARTIDAS, value):
                changed += 1
    summary["partidas_base"] += changed


def get_sala_maquinas_pipes(all_pipes):
    result = []
    seen = set()
    for pipe in all_pipes:
        system_type_name = get_pipe_system_type_name(pipe)
        abbreviation = get_pipe_system_abbreviation(pipe)
        matches = contains_any(system_type_name, [u"DRC", u"Conducción"]) or contains_any(abbreviation, [u"DESR"])
        if matches and pipe.Id.IntegerValue not in seen:
            seen.add(pipe.Id.IntegerValue)
            result.append(pipe)
    return result


def assign_sala_maquinas_partidas(summary, pipes):
    changed = 0
    for pipe in pipes:
        if set_param_safe(pipe, PARAM_PARTIDAS, u"01.07"):
            changed += 1
    summary["partidas_sala_maquinas"] += changed


def assign_panel_gross_area(summary):
    changed = 0
    for wall in collect_elements(BuiltInCategory.OST_Walls):
        length_ft = get_wall_length_ft(wall)
        height_ft = get_wall_unconnected_height_ft(wall)
        if length_ft is None or height_ft is None:
            continue
        if set_param_safe(wall, PARAM_SUP_BRUTA_PANEL, length_ft * height_ft * GROSS_FACTOR):
            changed += 1
    for floor in collect_elements(BuiltInCategory.OST_Floors):
        area_sqft = get_floor_area_sqft(floor)
        if area_sqft is None:
            continue
        if set_param_safe(floor, PARAM_SUP_BRUTA_PANEL, area_sqft * GROSS_FACTOR):
            changed += 1
    summary["sup_bruta_panel"] += changed


def assign_duct_areas(summary):
    changed = 0
    for duct in collect_elements(BuiltInCategory.OST_DuctCurves):
        width_ft = get_param_double(duct, [u"Anchura"])
        height_ft = get_param_double(duct, [u"Altura"])
        length_ft = get_param_double(duct, [u"Longitud"])
        if width_ft is None or height_ft is None or length_ft is None:
            continue
        gross_area_sqft = 2.0 * (width_ft + height_ft) * length_ft * GROSS_FACTOR
        if set_param_safe(duct, PARAM_DUCT_FITTING_AREA, gross_area_sqft):
            changed += 1
    for fitting in collect_elements(BuiltInCategory.OST_DuctFitting):
        total_face_area_sqft = get_total_face_area_sqft(fitting)
        connection_area_sqft = get_param_double(fitting, [PARAM_DUCT_CONNECTION_AREA], default=0.0) or 0.0
        gross_area_sqft = max(total_face_area_sqft - connection_area_sqft, 0.0) * GROSS_FACTOR
        if set_param_safe(fitting, PARAM_DUCT_FITTING_AREA, gross_area_sqft):
            changed += 1
    summary["duct_areas"] += changed


def get_target_pipe_type(rule):
    if rule["type_match"] == "exact":
        return find_pipe_type_by_exact_name(rule["target_type"])
    return find_pipe_type_by_contains(rule["target_type"])


def get_rule_pipe_subset(all_pipes, rule):
    subset = []
    seen = set()
    for pipe in all_pipes:
        if rule["filter_field"] == "system_abbreviation":
            field_value = get_pipe_system_abbreviation(pipe)
        else:
            field_value = get_pipe_system_type_name(pipe)
        if rule["filter_mode"] == "contains":
            matched = contains_any(field_value, rule["terms"])
        else:
            matched = equals_any(field_value, rule["terms"])
        if matched and pipe.Id.IntegerValue not in seen:
            seen.add(pipe.Id.IntegerValue)
            subset.append(pipe)
    return subset


def change_pipe_types(summary, all_pipes):
    changed_groups = {}
    for rule in PIPE_TYPE_RULES:
        subset = get_rule_pipe_subset(all_pipes, rule)
        changed_groups[rule["name"]] = subset
        target_type = get_target_pipe_type(rule)
        if not target_type:
            summary["pipe_type_missing_" + rule["name"]] += len(subset)
            continue
        changed = 0
        for pipe in subset:
            try:
                param = pipe.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
                if param and not param.IsReadOnly:
                    param.Set(target_type.Id)
                    changed += 1
            except Exception:
                continue
        summary["pipe_type_changed_" + rule["name"]] += changed
    return changed_groups


def assign_pipe_presto_codes(summary, changed_groups):
    changed = 0
    for rule in PIPE_TYPE_RULES:
        for pipe in changed_groups.get(rule["name"], []):
            diameter_ft = get_pipe_diameter_ft(pipe)
            code = find_mm_code(ft_to_mm(diameter_ft), rule["presto_map"])
            if code and set_param_safe(pipe, PARAM_CODIGO_PRESTO, code):
                changed += 1
    summary["codigo_presto_tuberias"] += changed


def assign_wall_floor_presto_codes(summary):
    changed = 0
    all_elements = []
    all_elements.extend(collect_elements(BuiltInCategory.OST_Walls))
    all_elements.extend(collect_elements(BuiltInCategory.OST_Floors))
    for elem in all_elements:
        code = get_param_text(get_element_type(elem), [PARAM_CODIGO_MONTAJE], default=u"")
        if code and set_param_safe(elem, PARAM_CODIGO_PRESTO, code):
            changed += 1
    summary["codigo_presto_cerramientos"] += changed


def assign_duct_presto_codes(summary):
    changed = 0
    for duct in collect_elements(BuiltInCategory.OST_DuctCurves):
        code = u"SIN ETIQUETA"
        try:
            duct_type_name = safe_text(get_duct_type_name(duct)).lower()
            for keyword, keyword_code in DUCT_TYPE_CODE_BY_KEYWORD:
                if keyword in duct_type_name:
                    code = keyword_code
                    break
        except Exception as err:
            code = u"Error: {}".format(safe_text(err))
        if set_param_safe(duct, PARAM_CODIGO_PRESTO, code):
            changed += 1
    summary["codigo_presto_conductos"] += changed


def assign_insulation_presto_codes(summary):
    changed = 0
    for insulation in collect_pipe_insulations():
        type_name = get_insulation_type_name(insulation)
        code = INSULATION_CODE_BY_TYPE_NAME.get(type_name)
        if code and set_param_safe(insulation, PARAM_CODIGO_PRESTO, code):
            changed += 1
    summary["codigo_presto_aislamientos"] += changed


def get_gross_total_pipe_length_m(total_net_m):
    gross_plus = get_revit_round(total_net_m + PIPE_WASTE_ADD_M)
    gross_bars = get_revit_round_digits(gross_plus / PIPE_WASTE_ADD_M, 0)
    return gross_bars * PIPE_WASTE_ADD_M


def assign_pipe_gross_lengths(summary, all_pipes, sala_maquinas_pipes):
    sala_ids = set(pipe.Id.IntegerValue for pipe in sala_maquinas_pipes)
    non_sala = [pipe for pipe in all_pipes if pipe.Id.IntegerValue not in sala_ids]

    def write_group(pipes, summary_key):
        groups = defaultdict(list)
        for pipe in pipes:
            size_key = get_param_text(pipe, [u"Tamaño"], default=u"")
            length_ft = get_pipe_length_ft(pipe)
            if length_ft is not None and length_ft > 0:
                groups[size_key].append((pipe, length_ft))
        if not groups:
            return
        changed_local = 0
        for _, items in groups.items():
            total_net_m = sum(ft_to_m(length_ft) for _, length_ft in items)
            if total_net_m <= 0:
                continue
            gross_total_m = get_gross_total_pipe_length_m(total_net_m)
            for pipe, length_ft in items:
                gross_length_m = ft_to_m(length_ft) * gross_total_m / total_net_m
                if set_param_safe(pipe, PARAM_LONG_BRUTA_TUB, m_to_ft(gross_length_m)):
                    changed_local += 1
        summary[summary_key] += changed_local

    write_group(sala_maquinas_pipes, "long_bruta_tub_sala_maquinas")
    write_group(non_sala, "long_bruta_tub_resto")


def assign_pipe_insulation_gross_lengths(summary):
    # Dynamo builds the gross-length groups from every pipe with a filled
    # "Tipo de aislamiento", then writes only to the insulation elements that
    # actually exist on those pipes. That means pipes with an insulation type
    # but no modeled insulation still affect the divisor and the total bars.
    insulations_by_pipe = get_hosted_pipe_insulations_by_pipe()
    groups = defaultdict(list)
    for pipe in collect_elements(BuiltInCategory.OST_PipeCurves):
        insulation_type_name = get_param_text(pipe, [u"Tipo de aislamiento"], default=u"")
        if u" " not in insulation_type_name:
            continue
        length_ft = get_pipe_length_ft(pipe)
        if length_ft is None or length_ft <= 0:
            continue
        groups[insulation_type_name].append(
            (length_ft, insulations_by_pipe.get(pipe.Id.IntegerValue, []))
        )

    changed = 0
    for _, items in groups.items():
        total_length_mm = sum(ft_to_mm(length_ft) for length_ft, _ in items)
        if total_length_mm <= 0:
            continue
        gross_total_m = get_revit_ceiling((total_length_mm + PIPE_BAR_LENGTH_MM) / PIPE_BAR_LENGTH_MM) * PIPE_BAR_LENGTH_M
        divisor = get_revit_ceiling(total_length_mm)
        if divisor <= 0:
            continue
        for length_ft, insulations in items:
            gross_length_m = gross_total_m * ft_to_mm(length_ft) / divisor
            for insulation in insulations:
                if set_param_safe(insulation, PARAM_LONG_BRUTA_TUB, m_to_ft(gross_length_m)):
                    changed += 1
    summary["long_bruta_tub_aislamientos"] += changed


def assign_refrigerant_labels(summary, all_pipes):
    changed = 0
    for pipe in all_pipes:
        classification = get_pipe_system_classification(pipe)
        if u"Sanitario" in safe_text(classification):
            continue
        fluid_type_name = get_pipe_fluid_type_name(pipe)
        if not fluid_type_name:
            continue
        label = None
        if u"R-744" in fluid_type_name:
            label = u"R-744"
        elif u"R-448A" in fluid_type_name:
            label = u"R-448A"
        elif u"R-134" in fluid_type_name:
            label = u"R-134a"
        if label and set_param_safe(pipe, PARAM_LEE_REFRIGERANTE, label):
            changed += 1
    summary["lee_refrigerante"] += changed


def delete_existing_wall_sweeps(summary):
    wall_sweeps = list(FilteredElementCollector(doc).OfClass(WallSweep).ToElements())
    summary["zocalos_borrados"] += delete_elements_by_ids([ws.Id for ws in wall_sweeps])


def delete_non_pipe_hosted_insulations(summary):
    ids_to_delete = []
    for insulation in collect_pipe_insulations():
        host_id = getattr(insulation, "HostElementId", None)
        host = doc.GetElement(host_id) if host_id and host_id != ElementId.InvalidElementId else None
        if host is None:
            ids_to_delete.append(insulation.Id)
            continue
        host_category = getattr(host, "Category", None)
        host_category_id = host_category.Id.IntegerValue if host_category else None
        if host_category_id != int(BuiltInCategory.OST_PipeCurves):
            ids_to_delete.append(insulation.Id)
    summary["aislamientos_fittings_borrados"] += delete_elements_by_ids(ids_to_delete)


def find_wall_sweep_type(type_name):
    for sweep_type in FilteredElementCollector(doc).WhereElementIsElementType():
        if get_element_name(sweep_type) == type_name:
            return sweep_type
    return None


def create_zocalos(summary):
    sweep_type = find_wall_sweep_type(u"CST_Zocalo")
    if sweep_type is None:
        summary["zocalos_sin_tipo"] += 1
        return

    created = []
    for wall in collect_elements(BuiltInCategory.OST_Walls):
        wall_type_name = get_element_name(get_element_type(wall))
        if not wall_type_name:
            continue
        if u"zocalo 2 lados" in wall_type_name:
            ext_info = WallSweepInfo(WallSweepType.Sweep, False)
            ext_info.Distance = 0.0
            ext_info.WallSide = WallSide.Exterior
            created.append(WallSweep.Create(wall, sweep_type.Id, ext_info))

            int_info = WallSweepInfo(WallSweepType.Sweep, False)
            int_info.Distance = 0.0
            int_info.WallSide = WallSide.Interior
            created.append(WallSweep.Create(wall, sweep_type.Id, int_info))

        if u"zocalo 1 lado" in wall_type_name:
            one_side = WallSweepInfo(WallSweepType.Sweep, False)
            one_side.Distance = 0.0
            created.append(WallSweep.Create(wall, sweep_type.Id, one_side))

    parametrized = 0
    for sweep in created:
        ok_partida = set_param_safe(sweep, PARAM_PARTIDAS, u"09")
        ok_codigo = set_param_safe(sweep, PARAM_CODIGO_PRESTO, u"ZOC.PP500.300")
        if ok_partida or ok_codigo:
            parametrized += 1

    summary["zocalos_creados"] += len(created)
    summary["zocalos_parametrizados"] += parametrized


def print_summary(summary, errors):
    rows = [
        [u"Habitaciones borradas", summary.get("rooms_deleted", 0)],
        [u"Vol.Cámara en equipos", summary.get("vol_camara", 0)],
        [u"Comentarios eléctricos", summary.get("comentarios_electrico", 0)],
        [u"Comentarios accesorios tubería", summary.get("comentarios_pipe_accessories", 0)],
        [u"Comentarios modelos genéricos", summary.get("comentarios_generic_models", 0)],
        [u"Comentarios muebles", summary.get("comentarios_furniture_systems", 0)],
        [u"Comentarios puertas", summary.get("comentarios_puertas", 0)],
        [u"Partidas base", summary.get("partidas_base", 0)],
        [u"Partidas sala máquinas", summary.get("partidas_sala_maquinas", 0)],
        [u"sup.bruta.panel", summary.get("sup_bruta_panel", 0)],
        [u"Áreas conductos/uniones", summary.get("duct_areas", 0)],
        [u"Código PRESTO tuberías", summary.get("codigo_presto_tuberias", 0)],
        [u"Código PRESTO conductos", summary.get("codigo_presto_conductos", 0)],
        [u"Código PRESTO cerramientos", summary.get("codigo_presto_cerramientos", 0)],
        [u"Código PRESTO aislamientos", summary.get("codigo_presto_aislamientos", 0)],
        [u"long.bruta.tub sala máquinas", summary.get("long_bruta_tub_sala_maquinas", 0)],
        [u"long.bruta.tub resto", summary.get("long_bruta_tub_resto", 0)],
        [u"long.bruta.tub aislamientos", summary.get("long_bruta_tub_aislamientos", 0)],
        [u"Lee_Refrigerante", summary.get("lee_refrigerante", 0)],
        [u"Aislamientos fittings borrados", summary.get("aislamientos_fittings_borrados", 0)],
        [u"Barridos de muro borrados", summary.get("zocalos_borrados", 0)],
        [u"Zócalos creados", summary.get("zocalos_creados", 0)],
        [u"Zócalos parametrizados", summary.get("zocalos_parametrizados", 0)],
    ]

    output.print_md("## ANTES DE MANDAR A PRESTO")
    if hasattr(output, "print_table"):
        output.print_table(table_data=rows, columns=[u"Bloque", u"Elementos"])
    else:
        output.print_md("| Bloque | Elementos |")
        output.print_md("|---|---|")
        for row in rows:
            output.print_md("| {} | {} |".format(row[0], row[1]))

    if errors:
        output.print_md("## Incidencias")
        for err in errors:
            output.print_md("- {}".format(err))


def main():
    summary = defaultdict(int)
    errors = []

    all_pipes = collect_elements(BuiltInCategory.OST_PipeCurves)
    pipe_accessories = collect_elements(BuiltInCategory.OST_PipeAccessory)
    generic_models = collect_elements(BuiltInCategory.OST_GenericModel)
    furniture_systems = collect_elements(BuiltInCategory.OST_FurnitureSystems)

    tx_group = TransactionGroup(doc, u"Antes de mandar a PRESTO (Dynamo)")
    tx_group.Start()

    try:
        valid_rooms = run_block(
            u"Dynamo - Borrar habitaciones huerfanas",
            lambda: delete_orphan_rooms(summary),
            errors,
        ) or []
        room_volume_map = build_room_volume_map(valid_rooms)

        run_block(u"Dynamo - Vol.Cámara en equipos", lambda: transfer_room_volume_to_equipment(summary, room_volume_map), errors)
        run_block(u"Dynamo - Comentarios eléctricos", lambda: set_comments_for_electrical_devices(summary), errors)
        run_block(
            u"Dynamo - Comentarios auxiliares",
            lambda: (
                set_comments_from_location(summary, pipe_accessories, "comentarios_pipe_accessories"),
                set_comments_from_location(summary, generic_models, "comentarios_generic_models"),
                set_comments_from_location(summary, furniture_systems, "comentarios_furniture_systems"),
            ),
            errors,
        )
        run_block(u"Dynamo - Comentarios puertas", lambda: set_door_comments(summary), errors)
        run_block(u"Dynamo - Partidas PRESTO base", lambda: assign_partidas_by_category(summary), errors)

        sala_maquinas_pipes = get_sala_maquinas_pipes(all_pipes)
        run_block(u"Dynamo - Partidas sala de máquinas", lambda: assign_sala_maquinas_partidas(summary, sala_maquinas_pipes), errors)
        run_block(u"Dynamo - sup.bruta.panel", lambda: assign_panel_gross_area(summary), errors)
        run_block(u"Dynamo - Áreas conductos", lambda: assign_duct_areas(summary), errors)

        changed_groups = run_block(u"Dynamo - Cambio de tipo de tuberías", lambda: change_pipe_types(summary, all_pipes), errors) or {}

        run_block(u"Dynamo - Codigo_Presto tuberías", lambda: assign_pipe_presto_codes(summary, changed_groups), errors)
        run_block(u"Dynamo - Codigo_Presto cerramientos", lambda: assign_wall_floor_presto_codes(summary), errors)
        run_block(u"Dynamo - Codigo_Presto conductos", lambda: assign_duct_presto_codes(summary), errors)
        run_block(u"Dynamo - Codigo_Presto aislamientos", lambda: assign_insulation_presto_codes(summary), errors)

        run_block(u"Dynamo - long.bruta.tub tuberías", lambda: assign_pipe_gross_lengths(summary, all_pipes, sala_maquinas_pipes), errors)
        run_block(u"Dynamo - Borrar aislamientos fittings", lambda: delete_non_pipe_hosted_insulations(summary), errors)
        run_block(u"Dynamo - long.bruta.tub aislamientos", lambda: assign_pipe_insulation_gross_lengths(summary), errors)
        run_block(u"Dynamo - Lee_Refrigerante", lambda: assign_refrigerant_labels(summary, all_pipes), errors)

        run_block(u"Dynamo - Borrar barridos de muro", lambda: delete_existing_wall_sweeps(summary), errors)
        run_block(u"Dynamo - Crear zócalos", lambda: create_zocalos(summary), errors)

        tx_group.Assimilate()
    except Exception as fatal_error:
        tx_group.RollBack()
        errors.append(u"FATAL: {}".format(safe_text(fatal_error)))

    print_summary(summary, errors)

    message = (
        u"ANTES DE MANDAR A PRESTO\n\n"
        u"Errores: {}\n"
        u"Rooms borradas: {}\n"
        u"Vol.Cámara: {}\n"
        u"Partidas base: {}\n"
        u"Código tuberías: {}\n"
        u"Zócalos creados: {}"
    ).format(
        len(errors),
        summary.get("rooms_deleted", 0),
        summary.get("vol_camara", 0),
        summary.get("partidas_base", 0),
        summary.get("codigo_presto_tuberias", 0),
        summary.get("zocalos_creados", 0),
    )
    AutoClosePopup(message, duration_ms=3000).ShowDialog()


if __name__ == "__main__":
    main()
