# -*- coding: utf-8 -*-

__title__ = "Evap/Door Location" 
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
_____________________________________________________________________
Description:

Set the evaporator and door "ubicacion" parameter with the room's name.
_____________________________________________________________________
How-to:

Just run the script to unpin set parametere values

IMPORTANT: IN CHAMBERS SHARED BY TWO WALLS, IT CAN FAIL, SO CHECK MANUALLY IF CORRECT VALUE WAS PLACED !!!
_____________________________________________________________________
Author: Juan Manuel Achenbach Anguita & ChatGPT"""


from pyrevit import revit, script
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import Architecture
from Autodesk.Revit.UI import TaskDialog

doc = revit.doc

TARGET_PARAM_NAME = "ubicación"
Z_OFFSETS_M = [0.0, 0.05, -0.05, 0.15, -0.15]  # metros
Z_OFFSETS_FT = [m / 0.3048 for m in Z_OFFSETS_M]  # Revit usa pies internamente


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


def get_elem_label(elem):
    """Devuelve una etiqueta legible para logs (nombre + id)."""
    if not elem:
        return "<None>"
    name = ""
    try:
        name = getattr(elem, "Name", "") or ""
    except:
        name = ""
    if not name:
        try:
            p = elem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            if p:
                name = p.AsString() or ""
        except:
            name = ""
    if not name:
        name = "<sin nombre>"
    return "{} (Id:{})".format(name, elem.Id.IntegerValue)


def get_room_label(room):
    """Etiqueta legible de habitacion para logs."""
    if not room:
        return "<None>"
    rname = get_room_name(room)
    if not rname:
        rname = "<sin nombre>"
    return "{} (Id:{})".format(rname, room.Id.IntegerValue)


def find_room_with_z_offsets(pt, rooms):
    """Busca room para un punto probando pequeños offsets en Z."""
    if not pt:
        return None
    for dz in Z_OFFSETS_FT:
        test_pt = pt if dz == 0 else XYZ(pt.X, pt.Y, pt.Z + dz)
        for r in rooms:
            try:
                if r.IsPointInRoom(test_pt):
                    return r
            except:
                pass
    return None


# ==================================================
# COLECTAR ELEMENTOS
# ==================================================
rooms = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_Rooms) \
    .WhereElementIsNotElementType() \
    .ToElements()

rooms_validas = [r for r in rooms if r.Location and r.Area > 0]

if not rooms_validas:
    TaskDialog.Show("INFO", "No hay habitaciones válidas en el proyecto.")
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
equipos_log = {}
puertas_log = {}
room_status = {}
for r in rooms_validas:
    room_status[r.Id.IntegerValue] = {
        "label": get_room_label(r),
        "mech_named": False,
        "door_named": False
    }

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
        room_id_match = None

        room_match = find_room_with_z_offsets(pt, rooms_validas)
        if room_match:
            room_name = get_room_name(room_match)
            room_id_match = room_match.Id.IntegerValue

        if room_name:
            if set_param(eq, TARGET_PARAM_NAME, room_name):
                equipos_modificados += 1
                equipos_log[eq.Id.IntegerValue] = (get_elem_label(eq), room_name)
                if room_id_match in room_status:
                    room_status[room_id_match]["mech_named"] = True

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
                        puertas_log[p.Id.IntegerValue] = (get_elem_label(p), room_name)
                        rid = r.Id.IntegerValue
                        if rid in room_status:
                            room_status[rid]["door_named"] = True

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

output = script.get_output()
output.print_md("## Log de elementos actualizados")

if equipos_log:
    output.print_md("### Mechanical Equipment")
    output.print_md("| Element | Value Set |")
    output.print_md("|---|---|")
    for eid in sorted(equipos_log.keys()):
        label, value = equipos_log[eid]
        output.print_md("| {} | `{}` |".format(label, value))
else:
    output.print_md("### Mechanical Equipment")
    output.print_md("| Element | Value Set |")
    output.print_md("|---|---|")
    output.print_md("| No se actualizaron equipos mecanicos. | - |")

if puertas_log:
    output.print_md("### Doors")
    output.print_md("| Element | Value Set |")
    output.print_md("|---|---|")
    for did in sorted(puertas_log.keys()):
        label, value = puertas_log[did]
        output.print_md("| {} | `{}` |".format(label, value))
else:
    output.print_md("### Doors")
    output.print_md("| Element | Value Set |")
    output.print_md("|---|---|")
    output.print_md("| No se actualizaron puertas. | - |")

rooms_missing = []
for rid in sorted(room_status.keys()):
    st = room_status[rid]
    if (not st["mech_named"]) or (not st["door_named"]):
        rooms_missing.append(st)

output.print_md("### Rooms Missing Assignments")
output.print_md("| Room | Mechanical Equipment Named | Door Named |")
output.print_md("|---|---|---|")
if rooms_missing:
    for st in rooms_missing:
        mech_val = "No" if not st["mech_named"] else "Yes"
        door_val = "No" if not st["door_named"] else "Yes"
        output.print_md("| {} | {} | {} |".format(st["label"], mech_val, door_val))
else:
    output.print_md("| All rooms have at least one mechanical equipment and one door named. | Yes | Yes |")
