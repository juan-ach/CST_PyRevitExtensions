# -*- coding: utf-8 -*-
"""

Bloques cubiertos:
1. Alta/verificacion de parametros compartidos.
2. Calculo de "Pot.Frigorifica" en tuberias.
3. Copia de "Temperatura de fluido" desde el tipo de sistema a cada tuberia.
4. Agrupacion por Tramo + Nombre de tipo para escribir "P.Carga_Tramo" y "LongTramo".
5. Calculo acumulado de longitud y perdida de carga desde las tuberias iniciales del Dynamo.
6. Propagacion de perdidas acumuladas a equipos mecanicos.

"""

import os
import traceback

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit import DB

BuiltInCategory = DB.BuiltInCategory
BuiltInParameter = DB.BuiltInParameter
FilteredElementCollector = DB.FilteredElementCollector
InstanceBinding = DB.InstanceBinding
StorageType = DB.StorageType
Transaction = DB.Transaction
TypeBinding = DB.TypeBinding

from pyrevit import revit, script

try:
    clr.AddReference("RevitNodes")
    import Revit
    clr.ImportExtensions(Revit.Elements)
    DYNAMO_WRAPPERS_AVAILABLE = True
except Exception:
    DYNAMO_WRAPPERS_AVAILABLE = False


try:
    basestring
except NameError:
    basestring = str

try:
    unicode
except NameError:
    unicode = str


doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
output = script.get_output()

try:
    output.set_title("PERDIDA CARGA POR TRAMO")
except Exception:
    pass


def insert_binding_with_group(binding_map, definition, binding):
    try:
        group_type = getattr(DB, "GroupTypeId", None)
        if group_type is not None:
            return binding_map.Insert(definition, binding, group_type.Data)
    except Exception:
        pass

    try:
        return binding_map.Insert(definition, binding, DB.BuiltInParameterGroup.PG_DATA)
    except Exception:
        return binding_map.Insert(definition, binding)


def resolve_project_helper_file(helper_dir_name, filename):
    """
    Busca un helper partiendo desde la carpeta del script y subiendo niveles.
    Si encuentra la raiz *.extension, resuelve tambien desde Helpers como en
    otros scripts del proyecto.
    """
    this_dir = os.path.dirname(os.path.abspath(__file__))
    cursor = this_dir
    extension_root = None

    while cursor and cursor != os.path.dirname(cursor):
        candidate_paths = [
            os.path.join(cursor, filename),
            os.path.join(cursor, helper_dir_name, filename),
            os.path.join(cursor, "Helpers", helper_dir_name, filename),
        ]
        for candidate in candidate_paths:
            if os.path.isfile(candidate):
                return candidate

        if cursor.lower().endswith(".extension"):
            extension_root = cursor
            break
        cursor = os.path.dirname(cursor)

    if extension_root:
        candidate = os.path.join(
            extension_root,
            "Helpers",
            helper_dir_name,
            filename,
        )
        if os.path.isfile(candidate):
            return candidate

    return ""


SHARED_PARAMS_FILENAME = "parametros compartidos revit .txt"
SHARED_PARAMS_HELPER_DIR = "Percargaxtramoparamcomp"

INITIAL_PIPE_UNIQUE_IDS = [
    "2bef12d8-aa12-4497-b5a7-c7bfc574fc2c-01152c93",
    "2bef12d8-aa12-4497-b5a7-c7bfc574fc2c-01152cfd",
    "e077c18c-9f33-4287-801e-fac741d7a510-01188e45",
    "e077c18c-9f33-4287-801e-fac741d7a510-01188f39",
]
INITIAL_SOURCE_SYSTEM_ORDER = [u"A1+", u"A1-", u"L1", u"L2"]
INITIAL_PIPE_PROXIMITY_TOLS_MM = [300.0, 500.0, 800.0]
SOURCE_EQUIPMENT_NAME_HINTS = (
    u"central",
    u"frigor",
    u"transcrit",
    u"rack",
    u"booster",
)

FILTER_FAMILY_KEYWORDS = (u"evap", u"mueble")
ASPIRATION_SYSTEM_NAMES = set([
    u"A1+",
    u"A1-",
    u"L1+_ASPIRACIÓN",
    u"L1-_ASPIRACIÓN",
    u"L2+_ASPIRACIÓN",
])
DIRECT_ASP_SYSTEM_NAMES = set([u"A1+", u"A1-"])
DIRECT_LIQ_SYSTEM_NAMES = set([u"L1", u"L2"])

PARAM_POT_FRIGO = u"Pot.Frigorifica"
PARAM_LONG_ACUM = u"longitud_acumulada"
PARAM_PCARGA_ACUM = u"P.Carga_Acumulada"
PARAM_DIF_ENTALPIA = u"diferencia_entalpia"
PARAM_PCARGA_ASP = u"p.carga.acum_asp"
PARAM_PCARGA_LIQ = u"p.carga.acum_liq"
PARAM_FLOW = u"Flujo"
PARAM_SYSTEM_TYPE = u"Tipo de sistema"
PARAM_DENSITY = u"Densidad de fluido"
PARAM_TEMP_FLUID = u"Temperatura de fluido"
PARAM_TYPE_NAME = u"Nombre de tipo"
PARAM_LONGITUD = u"Longitud"
PARAM_LONGTRAMO = u"LongTramo"
PARAM_PCARGA = u"Pérdida de carga"
PARAM_PCARGA_TRAMO = u"P.Carga_Tramo"
PARAM_TRAMO = u"Tramo"
PARAM_CATEGORY = u"Categoría"
PARAM_DIF_PRES_EVAP = u"Dif_Pres_Evap"
PARAM_DIF_PRES_LIQ = u"Dif_Pres_Liq"


REQUIRED_SHARED_PARAMS = [
    {
        "name": PARAM_POT_FRIGO,
        "categories": [BuiltInCategory.OST_PipeCurves],
        "is_type": False,
    },
    {
        "name": PARAM_LONG_ACUM,
        "categories": [BuiltInCategory.OST_PipeCurves],
        "is_type": False,
    },
    {
        "name": PARAM_PCARGA_ACUM,
        "categories": [BuiltInCategory.OST_PipeCurves],
        "is_type": False,
    },
    {
        "name": PARAM_DIF_ENTALPIA,
        "categories": [BuiltInCategory.OST_PipingSystem],
        "is_type": True,
    },
    {
        "name": PARAM_PCARGA_ASP,
        "categories": [BuiltInCategory.OST_MechanicalEquipment],
        "is_type": False,
    },
    {
        "name": PARAM_PCARGA_LIQ,
        "categories": [BuiltInCategory.OST_MechanicalEquipment],
        "is_type": False,
    },
]


def as_list(value):
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return list(value)
    return [value]


def safe_float(value, default=0.0):
    if value is None:
        return default
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = None
    try:
        text = unicode(value)
    except Exception:
        try:
            text = str(value)
        except Exception:
            return default
    if text is None:
        return default
    text = text.strip().replace(",", ".")
    if not text:
        return default
    try:
        return float(text)
    except Exception:
        return default


def safe_int(value, default=0):
    try:
        return int(round(safe_float(value, default)))
    except Exception:
        return default


def mm_to_internal_feet(value_mm):
    return safe_float(value_mm, 0.0) / 304.8


def internal_feet_to_mm(value_ft):
    return safe_float(value_ft, 0.0) * 304.8


def to_text(value):
    if value is None:
        return u""
    try:
        if isinstance(value, basestring):
            return value
        return unicode(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return u""


def is_category(elem, bic):
    try:
        return (
            elem is not None
            and elem.Category is not None
            and elem.Category.Id.IntegerValue == int(bic)
        )
    except Exception:
        return False


def get_param(elem, name):
    if elem is None:
        return None
    try:
        return elem.LookupParameter(name)
    except Exception:
        return None


def get_param_value(elem, name, default=None):
    param = get_param(elem, name)
    if param is None:
        return default

    try:
        if param.StorageType == StorageType.String:
            return param.AsString()
        if param.StorageType == StorageType.Double:
            return param.AsDouble()
        if param.StorageType == StorageType.Integer:
            return param.AsInteger()
        if param.StorageType == StorageType.ElementId:
            elem_id = param.AsElementId()
            if elem_id and elem_id.IntegerValue > 0:
                linked = doc.GetElement(elem_id)
                return linked if linked is not None else elem_id.IntegerValue
            return None
    except Exception:
        pass

    try:
        return param.AsValueString()
    except Exception:
        return default


def get_param_text(elem, name, default=u""):
    param = get_param(elem, name)
    if param is None:
        return default

    try:
        value = param.AsString()
        if value not in (None, u""):
            return to_text(value)
    except Exception:
        pass

    try:
        value = param.AsValueString()
        if value not in (None, u""):
            return to_text(value)
    except Exception:
        pass

    value = get_dynamo_param_value(elem, name, None)
    if value not in (None, u""):
        return to_text(value)

    return default


def get_linked_param_element(elem, name):
    param = get_param(elem, name)
    if param is None:
        return None

    try:
        if param.StorageType == StorageType.ElementId:
            elem_id = param.AsElementId()
            if elem_id and elem_id.IntegerValue > 0:
                linked = doc.GetElement(elem_id)
                if linked is not None:
                    return linked
    except Exception:
        pass

    return None


def to_dynamo_element(elem):
    if not DYNAMO_WRAPPERS_AVAILABLE or elem is None:
        return None

    if hasattr(elem, "GetParameterValueByName"):
        return elem

    try:
        return elem.ToDSType(True)
    except Exception:
        pass

    try:
        from Revit.Elements import ElementWrapper
        return ElementWrapper.Wrap(elem, True)
    except Exception:
        pass

    return None


def unwrap_revit_element(value):
    if value is None:
        return None

    try:
        internal = value.InternalElement
        if internal is not None:
            return internal
    except Exception:
        pass

    try:
        if isinstance(value, DB.Element):
            return value
    except Exception:
        pass

    try:
        if hasattr(value, "Id") and hasattr(value.Id, "IntegerValue"):
            linked = doc.GetElement(value.Id)
            if linked is not None:
                return linked
    except Exception:
        pass

    return value


def get_dynamo_param_value(elem, name, default=None):
    wrapped = to_dynamo_element(elem)
    if wrapped is not None:
        try:
            value = wrapped.GetParameterValueByName(name)
            if value is not None:
                return value
        except Exception as exc:
            logger.debug("Dynamo GetParameterValueByName fallo en '{}': {}".format(name, exc))

    param = get_param(elem, name)
    if param is None:
        return default

    try:
        if param.StorageType == StorageType.String:
            return param.AsString()
        if param.StorageType == StorageType.Double:
            return convert_from_internal_param_value(param, param.AsDouble())
        if param.StorageType == StorageType.Integer:
            return param.AsInteger()
        if param.StorageType == StorageType.ElementId:
            elem_id = param.AsElementId()
            if elem_id and elem_id.IntegerValue > 0:
                linked = doc.GetElement(elem_id)
                return linked if linked is not None else elem_id.IntegerValue
            return None
    except Exception:
        pass

    try:
        return param.AsValueString()
    except Exception:
        return default


def get_param_unit_token(param):
    if param is None:
        return None

    try:
        token = param.GetUnitTypeId()
        if token is not None:
            return token
    except Exception:
        pass

    try:
        data_type = param.Definition.GetDataType()
        if data_type is not None:
            fmt = doc.GetUnits().GetFormatOptions(data_type)
            if fmt is not None:
                return fmt.GetUnitTypeId()
    except Exception:
        pass

    try:
        return param.DisplayUnitType
    except Exception:
        pass

    try:
        unit_type = param.Definition.UnitType
        fmt = doc.GetUnits().GetFormatOptions(unit_type)
        if fmt is not None:
            return fmt.DisplayUnits
    except Exception:
        pass

    return None


def convert_from_internal_param_value(param, value):
    token = get_param_unit_token(param)
    if token is None:
        return safe_float(value, value)
    try:
        return DB.UnitUtils.ConvertFromInternalUnits(float(value), token)
    except Exception:
        return safe_float(value, value)


def convert_to_internal_param_value(param, value):
    token = get_param_unit_token(param)
    if token is None:
        return safe_float(value)
    try:
        return DB.UnitUtils.ConvertToInternalUnits(float(safe_float(value)), token)
    except Exception:
        return safe_float(value)


def set_param_value(elem, name, value):
    param = get_param(elem, name)
    if param is None or param.IsReadOnly:
        return False

    try:
        if param.StorageType == StorageType.String:
            param.Set(to_text(value))
            return True
        if param.StorageType == StorageType.Double:
            param.Set(float(safe_float(value)))
            return True
        if param.StorageType == StorageType.Integer:
            param.Set(safe_int(value))
            return True
        if param.StorageType == StorageType.ElementId:
            if hasattr(value, "Id"):
                param.Set(value.Id)
                return True
    except Exception as exc:
        logger.debug("No se pudo escribir '{}': {}".format(name, exc))
        return False

    return False


def set_dynamo_param_value(elem, name, value):
    wrapped = to_dynamo_element(elem)
    if wrapped is not None:
        try:
            wrapped.SetParameterByName(name, value)
            return True
        except Exception as exc:
            logger.debug("Dynamo SetParameterByName fallo en '{}': {}".format(name, exc))

    param = get_param(elem, name)
    if param is None or param.IsReadOnly:
        return False

    try:
        if param.StorageType == StorageType.String:
            param.Set(to_text(value))
            return True
        if param.StorageType == StorageType.Double:
            param.Set(float(convert_to_internal_param_value(param, value)))
            return True
        if param.StorageType == StorageType.Integer:
            param.Set(safe_int(value))
            return True
        if param.StorageType == StorageType.ElementId and hasattr(value, "Id"):
            param.Set(value.Id)
            return True
    except Exception as exc:
        logger.debug("Fallback SetParameterByName fallo en '{}': {}".format(name, exc))
        return False

    return False


def set_dynamep_param_value(elem, name, value):
    param = get_param(elem, name)
    if param is None or param.IsReadOnly:
        return False

    try:
        if param.StorageType == StorageType.String:
            param.Set(to_text(value))
            return True
        if param.StorageType == StorageType.Double:
            internal_value = convert_to_internal_param_value(param, value)
            param.Set(float(internal_value))
            return True
        if param.StorageType == StorageType.Integer:
            param.Set(safe_int(value))
            return True
        if param.StorageType == StorageType.ElementId and hasattr(value, "Id"):
            param.Set(value.Id)
            return True
    except Exception as exc:
        logger.debug("DynaMEP-style set fallo en '{}': {}".format(name, exc))
        return False

    return False


def get_type_name(elem):
    if elem is None:
        return u""

    for bip in (BuiltInParameter.ALL_MODEL_TYPE_NAME, BuiltInParameter.SYMBOL_NAME_PARAM):
        try:
            param = elem.get_Parameter(bip)
            if param:
                value = param.AsString() or param.AsValueString()
                if value:
                    return value
        except Exception:
            pass

    try:
        if hasattr(elem, "Name") and elem.Name:
            return elem.Name
    except Exception:
        pass

    return u""


def get_family_name(elem):
    if elem is None:
        return u""
    try:
        if hasattr(elem, "Symbol") and elem.Symbol and elem.Symbol.Family:
            return elem.Symbol.Family.Name or u""
    except Exception:
        pass
    return u""


def get_connectors(elem):
    if elem is None:
        return []

    try:
        if hasattr(elem, "MEPModel") and elem.MEPModel is not None:
            manager = elem.MEPModel.ConnectorManager
            if manager is not None:
                return list(manager.Connectors)
    except Exception:
        pass

    try:
        if hasattr(elem, "ConnectorManager") and elem.ConnectorManager is not None:
            return list(elem.ConnectorManager.Connectors)
    except Exception:
        pass

    return []


def get_connector_points(elem):
    points = []
    for connector in get_connectors(elem):
        try:
            origin = connector.Origin
        except Exception:
            origin = None
        if origin is not None:
            points.append(origin)
    return points


def get_element_point(elem):
    if elem is None:
        return None

    loc = getattr(elem, "Location", None)
    if loc is not None:
        try:
            point = getattr(loc, "Point", None)
            if point is not None:
                return point
        except Exception:
            pass
        try:
            curve = getattr(loc, "Curve", None)
            if curve is not None:
                return curve.Evaluate(0.5, True)
        except Exception:
            pass

    try:
        bbox = elem.get_BoundingBox(None)
        if bbox is not None:
            return DB.XYZ(
                (bbox.Min.X + bbox.Max.X) * 0.5,
                (bbox.Min.Y + bbox.Max.Y) * 0.5,
                (bbox.Min.Z + bbox.Max.Z) * 0.5,
            )
    except Exception:
        pass

    return None


def get_pipe_curve(pipe):
    if pipe is None:
        return None
    try:
        loc = pipe.Location
        curve = getattr(loc, "Curve", None)
        if curve is not None:
            return curve
    except Exception:
        pass
    return None


def get_pipe_endpoints(pipe):
    curve = get_pipe_curve(pipe)
    if curve is None:
        return []
    try:
        return [curve.GetEndPoint(0), curve.GetEndPoint(1)]
    except Exception:
        return []


def xyz_distance(a, b):
    if a is None or b is None:
        return float("inf")
    try:
        return a.DistanceTo(b)
    except Exception:
        dx = safe_float(a.X, 0.0) - safe_float(b.X, 0.0)
        dy = safe_float(a.Y, 0.0) - safe_float(b.Y, 0.0)
        dz = safe_float(a.Z, 0.0) - safe_float(b.Z, 0.0)
        return (dx * dx + dy * dy + dz * dz) ** 0.5


def distance_point_to_segment(pt, a, b):
    if pt is None or a is None or b is None:
        return float("inf")

    abx = b.X - a.X
    aby = b.Y - a.Y
    abz = b.Z - a.Z
    apx = pt.X - a.X
    apy = pt.Y - a.Y
    apz = pt.Z - a.Z
    ab2 = abx * abx + aby * aby + abz * abz
    if ab2 <= 1e-12:
        return xyz_distance(pt, a)

    t = (apx * abx + apy * aby + apz * abz) / ab2
    if t < 0.0:
        t = 0.0
    elif t > 1.0:
        t = 1.0

    proj = DB.XYZ(
        a.X + abx * t,
        a.Y + aby * t,
        a.Z + abz * t,
    )
    return xyz_distance(pt, proj)


def get_equipment_anchor_points(equipment):
    points = get_connector_points(equipment)
    if points:
        return points

    point = get_element_point(equipment)
    if point is not None:
        return [point]

    return []


def get_pipe_anchor_points(pipe):
    points = get_connector_points(pipe)
    if points:
        return points
    return get_pipe_endpoints(pipe)


def get_pipe_distance_to_points(pipe, points):
    if pipe is None or not points:
        return float("inf")

    best = float("inf")
    pipe_points = get_pipe_anchor_points(pipe)
    for point in points:
        for pipe_point in pipe_points:
            dist = xyz_distance(point, pipe_point)
            if dist < best:
                best = dist

        endpoints = get_pipe_endpoints(pipe)
        if len(endpoints) >= 2:
            dist = distance_point_to_segment(point, endpoints[0], endpoints[1])
            if dist < best:
                best = dist

    return best


def get_pipe_system_type(pipe):
    if pipe is None:
        return None

    linked = get_linked_param_element(pipe, PARAM_SYSTEM_TYPE)
    if linked is not None:
        return linked

    linked = unwrap_revit_element(get_dynamo_param_value(pipe, PARAM_SYSTEM_TYPE, None))
    if linked is not None:
        try:
            if isinstance(linked, DB.Element):
                return linked
        except Exception:
            pass

        try:
            if hasattr(linked, "Id") and hasattr(linked.Id, "IntegerValue"):
                elem = doc.GetElement(linked.Id)
                if elem is not None:
                    return elem
        except Exception:
            pass

    try:
        if pipe.MEPSystem is not None:
            type_id = pipe.MEPSystem.GetTypeId()
            if type_id and type_id.IntegerValue > 0:
                return doc.GetElement(type_id)
    except Exception:
        pass

    linked = get_param_value(pipe, PARAM_SYSTEM_TYPE, None)
    if linked is None:
        return None

    try:
        if hasattr(linked, "GetTypeId"):
            type_id = linked.GetTypeId()
            if type_id and type_id.IntegerValue > 0:
                type_elem = doc.GetElement(type_id)
                if type_elem is not None:
                    return type_elem
    except Exception:
        pass

    return linked if hasattr(linked, "Id") else None


def get_pipe_system_name(pipe):
    system_type = get_pipe_system_type(pipe)
    value = get_dynamo_param_value(system_type, PARAM_TYPE_NAME, None)
    if value not in (None, u""):
        return to_text(value)
    return get_type_name(system_type)


def get_pipe_group_type_name(pipe):
    system_elem = get_linked_param_element(pipe, PARAM_SYSTEM_TYPE)
    if system_elem is not None:
        value = get_param_text(system_elem, PARAM_TYPE_NAME, u"")
        if value:
            return value
        type_name = get_type_name(system_elem)
        if type_name:
            return type_name

    system_obj = get_dynamo_param_value(pipe, PARAM_SYSTEM_TYPE, None)
    value = get_dynamo_param_value(system_obj, PARAM_TYPE_NAME, None)
    if value not in (None, u""):
        return to_text(value)

    system_elem = unwrap_revit_element(system_obj)
    if system_elem is not None:
        value = get_param_text(system_elem, PARAM_TYPE_NAME, u"")
        if value:
            return value
        type_name = get_type_name(system_elem)
        if type_name:
            return type_name

    return get_pipe_system_name(pipe)


def get_pipe_group_system_text(pipe):
    value = get_pipe_group_type_name(pipe)
    if value:
        return value
    return get_param_text(pipe, PARAM_SYSTEM_TYPE, u"")


def get_pipe_group_tramo_text(pipe):
    value = get_param_text(pipe, PARAM_TRAMO, u"")
    if value:
        return value
    raw_value = get_dynamo_param_value(pipe, PARAM_TRAMO, u"")
    if isinstance(raw_value, float):
        rounded = round(raw_value)
        if abs(raw_value - rounded) < 1e-9:
            return to_text(int(rounded))
    return to_text(raw_value)


def get_dynamo_group_key(pipe):
    """
    Replica literalmente la clave del Dynamo:
        a + "_" + b
    donde:
        a = GetParameterValueByName("Tramo") en la tuberia
        b = GetParameterValueByName("Nombre de tipo") sobre el resultado de
            GetParameterValueByName("Tipo de sistema") en la tuberia
    """
    tramo_value = get_dynamo_param_value(pipe, PARAM_TRAMO, u"")
    tramo_text = to_text(tramo_value)

    system_obj = get_dynamo_param_value(pipe, PARAM_SYSTEM_TYPE, None)
    system_name = get_dynamo_param_value(system_obj, PARAM_TYPE_NAME, u"")
    system_text = to_text(system_name)

    if not system_text:
        # Fallback defensivo por si el wrapper de Dynamo no devuelve el tipo de
        # sistema como elemento enlazado en un proyecto concreto.
        linked = unwrap_revit_element(system_obj) or get_linked_param_element(pipe, PARAM_SYSTEM_TYPE)
        if linked is not None:
            system_text = get_param_text(linked, PARAM_TYPE_NAME, u"")
            if not system_text:
                system_text = get_type_name(linked)

    if not tramo_text:
        tramo_text = get_pipe_group_tramo_text(pipe)
    if not system_text:
        system_text = get_pipe_group_type_name(pipe)

    return u"{}_{}".format(tramo_text, system_text)


def get_connected_pipes_for_fitting(fitting):
    pipes = []
    seen = set()
    for connector in get_connectors(fitting):
        try:
            refs = list(connector.AllRefs)
        except Exception:
            refs = []
        for ref in refs:
            owner = getattr(ref, "Owner", None)
            if owner is None or owner.Id == fitting.Id:
                continue
            if owner.Id.IntegerValue in seen:
                continue
            if is_category(owner, BuiltInCategory.OST_PipeCurves):
                pipes.append(owner)
                seen.add(owner.Id.IntegerValue)
    return pipes


def get_selected_pipe_elements():
    pipes = []
    seen = set()
    try:
        for elem_id in uidoc.Selection.GetElementIds():
            elem = doc.GetElement(elem_id)
            if elem is None:
                continue
            if not is_category(elem, BuiltInCategory.OST_PipeCurves):
                continue
            if elem.Id.IntegerValue in seen:
                continue
            pipes.append(elem)
            seen.add(elem.Id.IntegerValue)
    except Exception:
        pass
    return pipes


def resolve_dynamo_selected_element(unique_id, candidate_elements=None):
    elem = None

    try:
        elem = doc.GetElement(unique_id)
    except Exception:
        elem = None

    if elem is not None:
        return elem

    if candidate_elements:
        target = unique_id.lower()
        for candidate in candidate_elements:
            try:
                if candidate.UniqueId and candidate.UniqueId.lower() == target:
                    return candidate
            except Exception:
                pass

    try:
        hex_id = unique_id.rsplit("-", 1)[-1]
        int_id = int(hex_id, 16)
        elem = doc.GetElement(DB.ElementId(int_id))
        if elem is not None:
            return elem
    except Exception:
        pass

    return None


def get_equipment_type_element(equipment):
    if equipment is None:
        return None
    try:
        type_id = equipment.GetTypeId()
        if type_id and type_id.IntegerValue > 0:
            return doc.GetElement(type_id)
    except Exception:
        pass
    return None


def get_equipment_label(equipment):
    parts = []
    try:
        family_name = get_family_name(equipment)
        if family_name:
            parts.append(family_name)
    except Exception:
        pass

    try:
        type_name = get_type_name(get_equipment_type_element(equipment))
        if type_name:
            parts.append(type_name)
    except Exception:
        pass

    if not parts:
        try:
            if hasattr(equipment, "Name") and equipment.Name:
                parts.append(to_text(equipment.Name))
        except Exception:
            pass

    return u" | ".join(parts)


def unique_elements_by_id(elements):
    output = []
    seen = set()
    for elem in as_list(elements):
        if elem is None:
            continue
        try:
            elem_id = elem.Id.IntegerValue
        except Exception:
            continue
        if elem_id in seen:
            continue
        seen.add(elem_id)
        output.append(elem)
    return output


def infer_initial_pipes_from_equipment(pipe_candidates, all_equipment):
    pipe_candidates = unique_elements_by_id(pipe_candidates)
    all_equipment = unique_elements_by_id(all_equipment)
    if not pipe_candidates or not all_equipment:
        return [], u""

    candidate_pipe_ids = set(pipe.Id.IntegerValue for pipe in pipe_candidates)
    excluded_ids = set(
        elem.Id.IntegerValue for elem in filter_target_equipment(all_equipment)
    )
    evaluated = []

    equipment_pool = [eq for eq in all_equipment if eq.Id.IntegerValue not in excluded_ids]
    if not equipment_pool:
        equipment_pool = list(all_equipment)

    for equipment in equipment_pool:
        connected_pipes = [
            pipe for pipe in get_equipment_pipe_lists(equipment)
            if pipe is not None and pipe.Id.IntegerValue in candidate_pipe_ids
        ]
        connected_pipes = unique_elements_by_id(connected_pipes)
        if not connected_pipes:
            continue

        pipes_by_system = {}
        all_system_names = []
        for pipe in connected_pipes:
            system_name = get_pipe_system_name(pipe)
            if system_name:
                all_system_names.append(system_name)
                pipes_by_system.setdefault(system_name, pipe)

        preferred_pipes = []
        preferred_names = []
        for system_name in INITIAL_SOURCE_SYSTEM_ORDER:
            pipe = pipes_by_system.get(system_name)
            if pipe is not None:
                preferred_pipes.append(pipe)
                preferred_names.append(system_name)

        selected_pipes = preferred_pipes if preferred_pipes else connected_pipes
        label = get_equipment_label(equipment)
        search_text = label.lower()
        keyword_score = 0
        for index, hint in enumerate(SOURCE_EQUIPMENT_NAME_HINTS):
            if hint in search_text:
                keyword_score += len(SOURCE_EQUIPMENT_NAME_HINTS) - index

        evaluated.append({
            "equipment": equipment,
            "label": label,
            "all_system_count": len(set(all_system_names)),
            "preferred_system_count": len(preferred_names),
            "keyword_score": keyword_score,
            "connected_pipe_count": len(connected_pipes),
            "selected_pipes": selected_pipes,
            "selected_systems": preferred_names or list(dict.fromkeys(all_system_names)),
        })

    if not evaluated:
        return [], u""

    evaluated.sort(
        key=lambda item: (
            item["preferred_system_count"],
            item["keyword_score"],
            item["connected_pipe_count"],
            item["all_system_count"],
            len(item["selected_pipes"]),
        ),
        reverse=True,
    )
    best = evaluated[0]
    systems_text = u", ".join(best["selected_systems"]) if best["selected_systems"] else u"sin_sistema"
    source_label = best["label"] or u"equipo_sin_nombre"
    source = u"inferido:{} [{}]".format(source_label, systems_text)
    return unique_elements_by_id(best["selected_pipes"]), source


def infer_initial_pipes_by_proximity(pipe_candidates, all_equipment, tolerance_ft):
    pipe_candidates = unique_elements_by_id(pipe_candidates)
    all_equipment = unique_elements_by_id(all_equipment)
    if not pipe_candidates or not all_equipment:
        return [], u""

    excluded_ids = set(
        elem.Id.IntegerValue for elem in filter_target_equipment(all_equipment)
    )
    equipment_pool = [eq for eq in all_equipment if eq.Id.IntegerValue not in excluded_ids]
    if not equipment_pool:
        equipment_pool = list(all_equipment)

    evaluated = []
    for equipment in equipment_pool:
        anchor_points = get_equipment_anchor_points(equipment)
        if not anchor_points:
            continue

        nearest_by_system = {}
        all_hits = []
        for pipe in pipe_candidates:
            dist = get_pipe_distance_to_points(pipe, anchor_points)
            if dist > tolerance_ft:
                continue

            system_name = get_pipe_system_name(pipe)
            all_hits.append((dist, pipe, system_name))
            if system_name:
                current = nearest_by_system.get(system_name)
                if current is None or dist < current[0]:
                    nearest_by_system[system_name] = (dist, pipe)

        if not all_hits:
            continue

        selected_pipes = []
        selected_systems = []
        total_dist = 0.0

        for system_name in INITIAL_SOURCE_SYSTEM_ORDER:
            info = nearest_by_system.get(system_name)
            if info is None:
                continue
            selected_pipes.append(info[1])
            selected_systems.append(system_name)
            total_dist += info[0]

        if not selected_pipes:
            seen_pipe_ids = set()
            for dist, pipe, system_name in sorted(all_hits, key=lambda item: item[0]):
                pipe_id = pipe.Id.IntegerValue
                if pipe_id in seen_pipe_ids:
                    continue
                seen_pipe_ids.add(pipe_id)
                selected_pipes.append(pipe)
                total_dist += dist
                if system_name and system_name not in selected_systems:
                    selected_systems.append(system_name)
                if len(selected_pipes) >= 4:
                    break

        label = get_equipment_label(equipment)
        search_text = label.lower()
        keyword_score = 0
        for index, hint in enumerate(SOURCE_EQUIPMENT_NAME_HINTS):
            if hint in search_text:
                keyword_score += len(SOURCE_EQUIPMENT_NAME_HINTS) - index

        avg_dist = total_dist / float(max(len(selected_pipes), 1))
        evaluated.append({
            "equipment": equipment,
            "label": label,
            "selected_pipes": unique_elements_by_id(selected_pipes),
            "selected_systems": selected_systems,
            "preferred_system_count": len(selected_systems),
            "keyword_score": keyword_score,
            "all_hit_count": len(all_hits),
            "avg_dist": avg_dist,
        })

    if not evaluated:
        return [], u""

    evaluated.sort(
        key=lambda item: (
            item["preferred_system_count"],
            item["keyword_score"],
            item["all_hit_count"],
            -item["avg_dist"],
        ),
        reverse=True,
    )
    best = evaluated[0]
    systems_text = u", ".join(best["selected_systems"]) if best["selected_systems"] else u"sin_sistema"
    source_label = best["label"] or u"equipo_sin_nombre"
    source = u"proximidad:{} [{}] <= {:.0f}mm".format(
        source_label,
        systems_text,
        internal_feet_to_mm(tolerance_ft),
    )
    return unique_elements_by_id(best["selected_pipes"]), source


def get_initial_pipes(pipe_candidates=None, all_equipment=None):
    found = []
    seen = set()

    for unique_id in INITIAL_PIPE_UNIQUE_IDS:
        elem = resolve_dynamo_selected_element(unique_id, pipe_candidates)
        if elem is None or not is_category(elem, BuiltInCategory.OST_PipeCurves):
            continue
        if elem.Id.IntegerValue in seen:
            continue
        found.append(elem)
        seen.add(elem.Id.IntegerValue)

    if found:
        return found, "dynamo"

    inferred, source = infer_initial_pipes_from_equipment(pipe_candidates, all_equipment)
    if inferred:
        return inferred, source

    for tol_mm in INITIAL_PIPE_PROXIMITY_TOLS_MM:
        inferred, source = infer_initial_pipes_by_proximity(
            pipe_candidates,
            all_equipment,
            mm_to_internal_feet(tol_mm),
        )
        if inferred:
            return inferred, source

    selected = get_selected_pipe_elements()
    if selected:
        return selected, "seleccion_actual"

    return [], "sin_inicio"


def ensure_shared_parameters():
    report = {
        "present": [],
        "updated": [],
        "added": [],
        "missing_in_file": [],
        "errors": [],
        "executed": False,
    }

    shared_params_file = resolve_project_helper_file(
        SHARED_PARAMS_HELPER_DIR,
        SHARED_PARAMS_FILENAME,
    )

    if not shared_params_file:
        report["errors"].append(
            u"No existe el archivo de parametros compartidos en Helpers\\{}\\{}".format(
                SHARED_PARAMS_HELPER_DIR,
                SHARED_PARAMS_FILENAME,
            )
        )
        return report

    app = doc.Application
    app.SharedParametersFilename = shared_params_file
    shared_file = app.OpenSharedParameterFile()
    if shared_file is None:
        report["errors"].append(u"No se pudo abrir el archivo de parametros compartidos.")
        return report

    binding_map = doc.ParameterBindings
    existing = {}

    iterator = binding_map.ForwardIterator()
    iterator.Reset()
    while iterator.MoveNext():
        definition = iterator.Key
        binding = iterator.Current
        name = getattr(definition, "Name", None)
        if not name:
            continue
        cats = set()
        try:
            for cat in binding.Categories:
                cats.add(cat.Id.IntegerValue)
        except Exception:
            pass
        existing[name] = {
            "definition": definition,
            "binding": binding,
            "categories": cats,
            "is_type": isinstance(binding, TypeBinding),
        }

    trans = Transaction(doc, "PARAMETROS COMPARTIDOS - PERDIDA CARGA POR TRAMO")
    trans.Start()
    try:
        for item in REQUIRED_SHARED_PARAMS:
            name = item["name"]
            required_cats = item["categories"]
            is_type_required = item["is_type"]

            if name in existing:
                info = existing[name]
                missing_cat_ids = [
                    cat for cat in required_cats if int(cat) not in info["categories"]
                ]
                if missing_cat_ids or info["is_type"] != is_type_required:
                    category_set = app.Create.NewCategorySet()
                    try:
                        for cat in info["binding"].Categories:
                            category_set.Insert(cat)
                    except Exception:
                        pass

                    for bic in missing_cat_ids:
                        cat = doc.Settings.Categories.get_Item(bic)
                        if cat is not None:
                            category_set.Insert(cat)

                    if is_type_required:
                        new_binding = app.Create.NewTypeBinding(category_set)
                    else:
                        new_binding = app.Create.NewInstanceBinding(category_set)

                    binding_map.ReInsert(info["definition"], new_binding)
                    report["updated"].append(name)
                else:
                    report["present"].append(name)
                continue

            external_definition = None
            for group in shared_file.Groups:
                candidate = group.Definitions.get_Item(name)
                if candidate is not None:
                    external_definition = candidate
                    break

            if external_definition is None:
                report["missing_in_file"].append(name)
                continue

            category_set = app.Create.NewCategorySet()
            for bic in required_cats:
                cat = doc.Settings.Categories.get_Item(bic)
                if cat is not None:
                    category_set.Insert(cat)

            if is_type_required:
                binding = app.Create.NewTypeBinding(category_set)
            else:
                binding = app.Create.NewInstanceBinding(category_set)

            insert_binding_with_group(binding_map, external_definition, binding)
            report["added"].append(name)

        trans.Commit()
        report["executed"] = True
    except Exception as exc:
        report["errors"].append(to_text(exc))
        try:
            trans.RollBack()
        except Exception:
            pass

    return report


def calculate_pot_frigorifica(flow, density, delta_h):
    flow_value = safe_float(flow)
    density_value = safe_float(density)
    delta_h_value = safe_float(delta_h)

    flow_m3_s = flow_value * 0.001
    delta_h_j_kg = delta_h_value * 1000.0
    watts = flow_m3_s * density_value * delta_h_j_kg
    return watts / 3.600


def feet_to_millimeters(length_in_feet):
    return safe_float(length_in_feet) / 304.8


def bar_to_pa(bar_value):
    return safe_float(bar_value) * 30480.0


def build_pipe_adjacency(pipes, fittings):
    adjacency = {}
    pipe_ids = set(pipe.Id.IntegerValue for pipe in pipes)

    for fitting in fittings:
        fitting_id = fitting.Id.IntegerValue
        connected_pipe_ids = [
            pipe.Id.IntegerValue
            for pipe in get_connected_pipes_for_fitting(fitting)
            if pipe.Id.IntegerValue in pipe_ids
        ]
        for pipe_id in connected_pipe_ids:
            adjacency.setdefault(fitting_id, []).append(pipe_id)
            adjacency.setdefault(pipe_id, []).append(fitting_id)

    return adjacency


def compute_cumulative_lengths(pipes, fittings, initial_pipes):
    pipe_ids = [pipe.Id.IntegerValue for pipe in pipes]
    pipe_length_dict = {}
    for pipe in pipes:
        pipe_length_dict[pipe.Id.IntegerValue] = feet_to_millimeters(
            get_dynamo_param_value(pipe, PARAM_LONGITUD, 0.0)
        )

    adjacency = build_pipe_adjacency(pipes, fittings)
    cumulative_lengths = {}
    pipe_paths = {}
    visited = set()

    def bfs(start_pipe_id):
        queue = [(start_pipe_id, pipe_length_dict.get(start_pipe_id, 0.0), [start_pipe_id])]
        visited.add(start_pipe_id)
        cumulative_lengths[start_pipe_id] = pipe_length_dict.get(start_pipe_id, 0.0)
        pipe_paths[start_pipe_id] = [start_pipe_id]

        while queue:
            current_elem, cum_length, path = queue.pop(0)
            for neighbor in adjacency.get(current_elem, []):
                if neighbor in visited:
                    continue
                visited.add(neighbor)
                if neighbor in pipe_ids:
                    new_length = cum_length + pipe_length_dict.get(neighbor, 0.0)
                    cumulative_lengths[neighbor] = new_length
                    pipe_paths[neighbor] = path + [neighbor]
                    queue.append((neighbor, new_length, path + [neighbor]))
                else:
                    queue.append((neighbor, cum_length, path))

    for start_pipe in initial_pipes:
        start_id = start_pipe.Id.IntegerValue
        if start_id in pipe_ids and start_id not in visited:
            bfs(start_id)

    output = {}
    for pipe in pipes:
        output[pipe.Id.IntegerValue] = cumulative_lengths.get(pipe.Id.IntegerValue, 0.0)

    return output, pipe_paths


def compute_cumulative_pressure_drop(pipes, fittings, initial_pipes):
    pipe_ids = [pipe.Id.IntegerValue for pipe in pipes]
    pipe_pressure_drop = {}
    for pipe in pipes:
        pipe_pressure_drop[pipe.Id.IntegerValue] = bar_to_pa(
            get_dynamo_param_value(pipe, PARAM_PCARGA, 0.0)
        )

    initial_pipe_ids = [pipe.Id.IntegerValue for pipe in initial_pipes]
    initial_pipe_set = set(initial_pipe_ids)
    adjacency = build_pipe_adjacency(pipes, fittings)
    cumulative = {}
    pipe_paths = {}

    for start_pipe_id in initial_pipe_ids:
        if start_pipe_id not in pipe_ids:
            continue

        cumulative[start_pipe_id] = pipe_pressure_drop.get(start_pipe_id, 0.0)
        pipe_paths[start_pipe_id] = [str(start_pipe_id)]

        queue = [(start_pipe_id, pipe_pressure_drop.get(start_pipe_id, 0.0), [start_pipe_id])]
        visited_fittings_local = set()
        visited_pipes_local = set([start_pipe_id])

        while queue:
            current_elem, cum_pressure_drop, path = queue.pop(0)

            if current_elem not in pipe_ids:
                if current_elem in visited_fittings_local:
                    continue
                visited_fittings_local.add(current_elem)

            for neighbor in adjacency.get(current_elem, []):
                if neighbor in initial_pipe_set and neighbor != start_pipe_id:
                    continue
                if neighbor in visited_pipes_local:
                    continue

                if neighbor in pipe_ids:
                    neighbor_drop = pipe_pressure_drop.get(neighbor, 0.0)
                    new_cumulative = cum_pressure_drop + neighbor_drop
                    new_path = path + [neighbor]
                    if (
                        neighbor not in cumulative
                        or new_cumulative > cumulative[neighbor]
                    ):
                        cumulative[neighbor] = new_cumulative
                        pipe_paths[neighbor] = [str(pid) for pid in new_path]
                        queue.append((neighbor, new_cumulative, new_path))
                        visited_pipes_local.add(neighbor)
                else:
                    if neighbor not in visited_fittings_local:
                        queue.append((neighbor, cum_pressure_drop, path))

    output = {}
    for pipe in pipes:
        output[pipe.Id.IntegerValue] = cumulative.get(pipe.Id.IntegerValue, 0.0)

    return output, pipe_paths


def filter_mechanical_equipment_with_pot(all_equipment):
    result = []
    for elem in all_equipment:
        param = get_param(elem, PARAM_POT_FRIGO)
        if param is None:
            continue
        value = safe_float(get_dynamo_param_value(elem, PARAM_POT_FRIGO, 0.0), 0.0)
        if value != 0:
            result.append(elem)
    return result


def filter_target_equipment(all_equipment):
    filtered = []
    for elem in all_equipment:
        family_name = get_family_name(elem).lower()
        if not family_name:
            continue
        if not any(keyword in family_name for keyword in FILTER_FAMILY_KEYWORDS):
            continue
        param = get_param(elem, PARAM_POT_FRIGO)
        if param is None:
            continue

        value = safe_float(get_dynamo_param_value(elem, PARAM_POT_FRIGO, 0.0), 0.0)

        if value > 0:
            filtered.append(elem)

    return filtered


def filter_dynamo_pressure_equipment(all_equipment):
    filtered = []
    for elem in all_equipment:
        eq_type = None
        try:
            type_id = elem.GetTypeId()
            if type_id and type_id.IntegerValue > 0:
                eq_type = doc.GetElement(type_id)
        except Exception:
            eq_type = None

        type_name = get_type_name(eq_type).lower()
        if not type_name:
            continue
        if "evap" in type_name or "mueble" in type_name:
            filtered.append(elem)

    return filtered


def find_pipe_through_fittings(element, previous_id, visited_ids):
    for connector in get_connectors(element):
        try:
            refs = list(connector.AllRefs)
        except Exception:
            refs = []
        for ref in refs:
            owner = getattr(ref, "Owner", None)
            if owner is None:
                continue
            owner_id = owner.Id.IntegerValue
            if owner_id == previous_id or owner_id in visited_ids:
                continue
            visited_ids.add(owner_id)
            if is_category(owner, BuiltInCategory.OST_PipeCurves):
                return owner
            if is_category(owner, BuiltInCategory.OST_PipeFitting):
                pipe = find_pipe_through_fittings(owner, element.Id.IntegerValue, visited_ids)
                if pipe is not None:
                    return pipe
    return None


def get_equipment_pipe_lists(equipment):
    result = []
    for connector in get_connectors(equipment):
        pipe_found = None
        try:
            refs = list(connector.AllRefs)
        except Exception:
            refs = []
        visited_ids = set([equipment.Id.IntegerValue])
        for ref in refs:
            owner = getattr(ref, "Owner", None)
            if owner is None or owner.Id == equipment.Id:
                continue
            visited_ids.add(owner.Id.IntegerValue)
            if is_category(owner, BuiltInCategory.OST_PipeCurves):
                pipe_found = owner
            elif is_category(owner, BuiltInCategory.OST_PipeFitting):
                pipe_found = find_pipe_through_fittings(owner, equipment.Id.IntegerValue, visited_ids)
            if pipe_found is not None:
                result.append(pipe_found)
                break
        else:
            result.append(None)
    return result


def get_first_direct_pipe(equipment):
    for connector in get_connectors(equipment):
        try:
            refs = list(connector.AllRefs)
        except Exception:
            refs = []
        for ref in refs:
            owner = getattr(ref, "Owner", None)
            if owner is None or owner.Id == equipment.Id:
                continue
            if is_category(owner, BuiltInCategory.OST_PipeCurves):
                return owner
    return None


def update_equipment_pressures(target_equipment):
    stats = {
        "equipments": len(target_equipment),
        "recursive_asp": 0,
        "recursive_liq": 0,
        "fallback_asp": 0,
        "fallback_liq": 0,
        "missing_clean_lists": 0,
    }

    for equipment in target_equipment:
        pipe_candidates = [pipe for pipe in get_equipment_pipe_lists(equipment) if pipe is not None]
        if not pipe_candidates:
            stats["missing_clean_lists"] += 1
            pipe_candidates = []

        wrote_asp = False
        wrote_liq = False
        asp_pipes = []
        liq_pipes = []
        for pipe in pipe_candidates:
            system_name = get_pipe_system_name(pipe)
            if system_name in ASPIRATION_SYSTEM_NAMES:
                asp_pipes.append(pipe)
            else:
                liq_pipes.append(pipe)

        if asp_pipes:
            asp_value = get_dynamo_param_value(asp_pipes[0], PARAM_PCARGA_ACUM, None)
            converted = safe_float(asp_value, None)
            if converted is not None:
                converted = converted / 0.0689476 * 10000.0 * 2.0
                if set_dynamep_param_value(equipment, PARAM_PCARGA_ASP, converted):
                    stats["recursive_asp"] += 1
                    wrote_asp = True

        if liq_pipes:
            liq_value = get_dynamo_param_value(liq_pipes[0], PARAM_PCARGA_ACUM, None)
            converted = safe_float(liq_value, None)
            if converted is not None:
                converted = converted / 0.0689476 * 10000.0 * 2.0
                if set_dynamep_param_value(equipment, PARAM_PCARGA_LIQ, converted):
                    stats["recursive_liq"] += 1
                    wrote_liq = True

        direct_pipe = get_first_direct_pipe(equipment)
        if direct_pipe is None:
            continue

        direct_value = get_dynamo_param_value(direct_pipe, PARAM_PCARGA_ACUM, None)
        if direct_value is None:
            continue

        system_name = get_pipe_system_name(direct_pipe)
        if system_name in DIRECT_ASP_SYSTEM_NAMES and not wrote_asp:
            converted = safe_float(direct_value, None)
            if converted is not None:
                converted = converted / 0.0689476 * 10000.0 * 2.0
                if set_dynamep_param_value(equipment, PARAM_PCARGA_ASP, converted):
                    stats["fallback_asp"] += 1
        elif system_name in DIRECT_LIQ_SYSTEM_NAMES and not wrote_liq:
            converted = safe_float(direct_value, None)
            if converted is not None:
                converted = converted / 0.0689476 * 10000.0 * 2.0
                if set_dynamep_param_value(equipment, PARAM_PCARGA_LIQ, converted):
                    stats["fallback_liq"] += 1

    return stats


def main():
    print("=" * 70)
    print("PERDIDA CARGA POR TRAMO")
    print("=" * 70)

    print("\n[0] Parametros compartidos...")
    shared_report = ensure_shared_parameters()
    if shared_report["errors"]:
        for item in shared_report["errors"]:
            print("    ERROR: {}".format(item))
    else:
        print(
            "    presentes={} | actualizados={} | agregados={}".format(
                len(shared_report["present"]),
                len(shared_report["updated"]),
                len(shared_report["added"]),
            )
        )
    if shared_report["missing_in_file"]:
        print("    No encontrados en el archivo compartido: {}".format(
            ", ".join(shared_report["missing_in_file"])
        ))

    pipes = list(
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_PipeCurves)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    fittings = list(
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_PipeFitting)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    all_equipment = list(
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    print("\n[1] Colecciones base...")
    print("    Tuberias: {}".format(len(pipes)))
    print("    Uniones : {}".format(len(fittings)))
    print("    Equipos : {}".format(len(all_equipment)))

    initial_pipes, initial_source = get_initial_pipes(pipes, all_equipment)
    print("\n[2] Tuberias iniciales...")
    print("    Fuente: {}".format(initial_source))
    print("    Cantidad: {}".format(len(initial_pipes)))

    if not pipes:
        print("\nNo hay tuberias en el modelo. Fin.")
        return

    if not initial_pipes:
        print("\nERROR: No se pudieron resolver las tuberias iniciales del Dynamo.")
        print("Selecciona manualmente las tuberias de arranque o revisa los UniqueId embebidos en el script.")
        return

    target_equipment = filter_target_equipment(all_equipment)
    all_pot_equipment = filter_mechanical_equipment_with_pot(all_equipment)

    trans = Transaction(doc, "PERDIDA CARGA POR TRAMO")
    trans.Start()
    try:
        print("\n[3] Pot.Frigorifica en tuberias...")
        pot_written = 0
        pot_missing = 0
        for pipe in pipes:
            system_type = get_pipe_system_type(pipe)
            flow = get_dynamo_param_value(pipe, PARAM_FLOW, 0.0)
            density = get_dynamo_param_value(system_type, PARAM_DENSITY, 0.0)
            delta_h = get_dynamo_param_value(system_type, PARAM_DIF_ENTALPIA, 0.0)
            pot_value = calculate_pot_frigorifica(flow, density, delta_h)
            if set_dynamo_param_value(pipe, PARAM_POT_FRIGO, pot_value):
                pot_written += 1
            else:
                pot_missing += 1
        print("    Escritas: {} | sin parametro/solo lectura: {}".format(
            pot_written, pot_missing
        ))

        print("\n[4] Temperatura de fluido desde tipo de sistema...")
        temp_written = 0
        for pipe in pipes:
            system_type = get_pipe_system_type(pipe)
            temp_value = get_dynamo_param_value(system_type, PARAM_TEMP_FLUID, None)
            if temp_value is None:
                continue
            if set_dynamo_param_value(pipe, PARAM_TEMP_FLUID, temp_value):
                temp_written += 1
        print("    Escritas: {}".format(temp_written))

        print("\n[5] Agrupacion por Tramo + Nombre de tipo...")
        pressure_by_key = {}
        length_by_key = {}
        pipe_key_map = {}
        grouped = {}

        for pipe in pipes:
            key = get_dynamo_group_key(pipe)
            pipe_key_map[pipe.Id.IntegerValue] = key
            grouped.setdefault(key, []).append(pipe)

        for key, group_pipes in grouped.items():
            pressure_values = [
                safe_float(get_dynamo_param_value(pipe, PARAM_PCARGA, 0.0), 0.0)
                for pipe in group_pipes
            ]
            pressure_by_key[key] = sum(pressure_values) if pressure_values else 0.0
            length_by_key[key] = sum(
                safe_float(get_dynamo_param_value(pipe, PARAM_LONGITUD, 0.0), 0.0)
                for pipe in group_pipes
            )

        multi_groups = [group_pipes for group_pipes in grouped.values() if len(group_pipes) > 1]
        multi_group_pipe_count = sum(len(group_pipes) for group_pipes in multi_groups)

        tramo_pressure_written = 0
        tramo_length_written = 0
        for pipe in pipes:
            key = pipe_key_map.get(pipe.Id.IntegerValue, u"")
            if set_dynamo_param_value(pipe, PARAM_PCARGA_TRAMO, pressure_by_key.get(key, 0.0)):
                tramo_pressure_written += 1
            if set_dynamo_param_value(pipe, PARAM_LONGTRAMO, length_by_key.get(key, 0.0)):
                tramo_length_written += 1

        print("    Grupos: {}".format(len(grouped)))
        print("    Grupos multi-tuberia: {} | tuberias agrupadas: {}".format(
            len(multi_groups), multi_group_pipe_count
        ))
        print("    P.Carga_Tramo escritas: {}".format(tramo_pressure_written))
        print("    LongTramo escritos    : {}".format(tramo_length_written))

        print("\n[6] Longitud acumulada...")
        cumulative_lengths, length_paths = compute_cumulative_lengths(
            pipes, fittings, initial_pipes
        )
        long_acum_written = 0
        for pipe in pipes:
            value = cumulative_lengths.get(pipe.Id.IntegerValue, 0.0)
            if set_dynamep_param_value(pipe, PARAM_LONG_ACUM, value):
                long_acum_written += 1
        print("    Escritas: {}".format(long_acum_written))
        print("    Rutas calculadas: {}".format(len(length_paths)))

        print("\n[7] Perdida de carga acumulada...")
        cumulative_pressure, pressure_paths = compute_cumulative_pressure_drop(
            pipes, fittings, initial_pipes
        )
        pcarga_acum_written = 0
        for pipe in pipes:
            value = cumulative_pressure.get(pipe.Id.IntegerValue, 0.0)
            if set_dynamep_param_value(pipe, PARAM_PCARGA_ACUM, value):
                pcarga_acum_written += 1
        print("    Escritas: {}".format(pcarga_acum_written))
        print("    Rutas calculadas: {}".format(len(pressure_paths)))

        print("\n[8] Equipos mecanicos...")
        print("    Con Pot.Frigorifica != 0: {}".format(len(all_pot_equipment)))
        print("    Objetivo evap/mueble    : {}".format(len(target_equipment)))
        equipment_stats = update_equipment_pressures(target_equipment)
        print(
            "    Recursivo asp/liquido   : {}/{}".format(
                equipment_stats["recursive_asp"],
                equipment_stats["recursive_liq"],
            )
        )
        print(
            "    Fallback directo asp/liq: {}/{}".format(
                equipment_stats["fallback_asp"],
                equipment_stats["fallback_liq"],
            )
        )
        if equipment_stats["missing_clean_lists"]:
            print("    Equipos sin tuberias trazables: {}".format(
                equipment_stats["missing_clean_lists"]
            ))

        trans.Commit()

        print("\n" + "=" * 70)
        print("RESUMEN")
        print("=" * 70)
        print("Pot.Frigorifica en tuberias     : {}".format(pot_written))
        print("Temperatura de fluido escrita   : {}".format(temp_written))
        print("P.Carga_Tramo escritos          : {}".format(tramo_pressure_written))
        print("LongTramo escritos              : {}".format(tramo_length_written))
        print("longitud_acumulada escrita      : {}".format(long_acum_written))
        print("P.Carga_Acumulada escrita       : {}".format(pcarga_acum_written))
        print("Equipos objetivo procesados     : {}".format(len(target_equipment)))
        print("=" * 70)

    except Exception as exc:
        try:
            trans.RollBack()
        except Exception:
            pass
        print("\nERROR FATAL: {}".format(exc))
        traceback.print_exc()


if __name__ == "__main__":
    main()
