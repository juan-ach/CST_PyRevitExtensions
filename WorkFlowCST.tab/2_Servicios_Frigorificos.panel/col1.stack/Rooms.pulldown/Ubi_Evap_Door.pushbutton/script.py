# -*- coding: utf-8 -*-

__title__ = "Evap/Door Location" 
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
Date    = 10.12.2025
_____________________________________________________________________
Description:

Set the evaporator and door "ubicacion" parameter with the room's name.
_____________________________________________________________________
How-to:

Just run the script to unpin set parametere values

IMPORTANT: IN CHAMBERS SHARED BY TWO WALLS, IT CAN FAIL, SO CHECK MANUALLY IF CORRECT VALUE WAS PLACED !!!
_____________________________________________________________________
Last update: 10.12.2025
_____________________________________________________________________
Author: Juan Manuel Achenbach Anguita & ChatGPT"""


from pyrevit import revit
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import Architecture
from Autodesk.Revit.UI import TaskDialog

doc = revit.doc
view = doc.ActiveView

TARGET_PARAM_NAME = "ubicación"


def set_param(elem, name, value):
    """Escribe value en el parámetro 'name' si existe y no es de solo lectura."""
    if not elem:
        return False
    p = elem.LookupParameter(name)
    if p and not p.IsReadOnly:
        try:
            p.Set(value)
            return True
        except:
            return False
    return False


def get_point(elem):
    """Devuelve un punto representativo del elemento (LocationPoint) o None."""
    loc = elem.Location
    if isinstance(loc, LocationPoint):
        return loc.Point
    return None


def get_room_name(room):
    """Devuelve el nombre de la habitación de forma segura."""
    if not room:
        return u""
    # 1) Intentar atributo Name
    name_attr = getattr(room, "Name", None)
    if name_attr:
        return name_attr
    # 2) Intentar parámetro ROOM_NAME
    try:
        p = room.get_Parameter(BuiltInParameter.ROOM_NAME)
        if p and p.AsString():
            return p.AsString()
    except:
        pass
    return u""


# ==================================================
# VALIDACIONES
# ==================================================
if not isinstance(view, ViewPlan):
    TaskDialog.Show("INFO", "La vista activa debe ser una planta.")
    raise SystemExit

nivel = view.GenLevel
if not nivel:
    TaskDialog.Show("INFO", "La vista no tiene nivel asociado.")
    raise SystemExit

# ==================================================
# COLECTAR ELEMENTOS
# ==================================================
rooms = FilteredElementCollector(doc, view.Id) \
    .OfCategory(BuiltInCategory.OST_Rooms) \
    .WhereElementIsNotElementType() \
    .ToElements()

rooms_validas = [r for r in rooms if r.Location and r.Area > 0]

if not rooms_validas:
    TaskDialog.Show("INFO", "No hay habitaciones válidas en la vista.")
    raise SystemExit

equipos = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_MechanicalEquipment) \
    .WhereElementIsNotElementType() \
    .ToElements()

puertas = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_Doors) \
    .WhereElementIsNotElementType() \
    .ToElements()

# Mapa muro -> puertas
puertas_por_muro = {}
for p in puertas:
    host = getattr(p, "Host", None)
    if host and isinstance(host, Wall):
        puertas_por_muro.setdefault(host.Id, []).append(p)

boundary_options = SpatialElementBoundaryOptions()
boundary_options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish

equipos_modificados = 0
puertas_modificadas = 0

# ==================================================
# TRANSACCIÓN SEGURA
# ==================================================
t = Transaction(doc, "Asignar 'ubicación' desde habitación")

try:
    t.Start()

    # -------------------------------------------
    # 1) Asignar a equipos mecánicos
    # -------------------------------------------
    for eq in equipos:
        pt = get_point(eq)
        if not pt:
            continue

        room_name = u""

        for r in rooms_validas:
            try:
                if r.IsPointInRoom(pt):
                    room_name = get_room_name(r)
                    break
            except:
                pass

        if room_name:
            if set_param(eq, TARGET_PARAM_NAME, room_name):
                equipos_modificados += 1

    # -------------------------------------------
    # 2) Asignar a puertas en muros límite de rooms
    # -------------------------------------------
    for r in rooms_validas:
        room_name = get_room_name(r)
        if not room_name:
            continue

        try:
            blists = r.GetBoundarySegments(boundary_options)
        except:
            blists = None

        if not blists:
            continue

        muros_limite = set()

        for seg_list in blists:
            for seg in seg_list:
                elem_id = seg.ElementId
                if elem_id == ElementId.InvalidElementId:
                    continue

                m = doc.GetElement(elem_id)
                if isinstance(m, Wall):
                    muros_limite.add(m.Id)

        for mid in muros_limite:
            if mid in puertas_por_muro:
                for p in puertas_por_muro[mid]:
                    if set_param(p, TARGET_PARAM_NAME, room_name):
                        puertas_modificadas += 1

    t.Commit()

except Exception as e:
    t.RollBack()
    TaskDialog.Show("ERROR", "Se produjo un error dentro del script:\n\n{}".format(e))
    raise SystemExit

# ==================================================
# MENSAJE FINAL
# ==================================================
TaskDialog.Show(
    "RESULTADO",
    "Habitaciones analizadas: {}\n"
    "Equipos con 'ubicación' asignada: {}\n"
    "Puertas con 'ubicación' asignada: {}\n".format(
        len(rooms_validas), equipos_modificados, puertas_modificadas
    )
)
