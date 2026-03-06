# -*- coding: utf-8 -*-

__title__ = "Create Circuits" 
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
Date    = 15.12.2025
_____________________________________________________________________
Description:

Create Circuits between electrical panels and electrical connectors

_____________________________________________________________________
How-to:

Place the electrical panels, equipment and cable tray and then run the script

IMPORTANT: IF THE SCRIPT DON'T RECOGNIZE THE PANEL CHECK PANEL'S TYPE IN PANEL_NAME_FILTER LIST!
_____________________________________________________________________
Last update: 15.12.2025
_____________________________________________________________________
Author: Juan Manuel Achenbach Anguita & ChatGPT"""

# pyRevit – Crear un circuito por equipo mecánico (Evaporadores / Mueble)
# usando el conector eléctrico y asignarlo a un tablero (por nombre de FAMILIA).

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import ElectricalSystem, ElectricalSystemType
from Autodesk.Revit.UI import TaskDialog

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# ----------------------------------------------------------------------
# CONFIGURACIÓN
# ----------------------------------------------------------------------
# Lista de posibles textos que puede contener el nombre del TABLERO (instancia o tipo)
PANEL_NAME_FILTER = [
    "Cuadro Servicios Frigorificos",
    "Cuadro Servicios Frigoríficos HD",
    "Cuadro Servicios Frigoríficos SD",
    "Quadre Serveis Frigorífics",
    "Cuadro Central frigorífica"
    # Añade aquí los posibles nombres de Cuadros Eléctricos
]


# ----------------------------------------------------------------------
# FUNCIONES AUXILIALES
# ----------------------------------------------------------------------

def _normalize_name_filters(name_filters):
    """Admite string o lista/tupla y devuelve lista de filtros en minúsculas."""
    if name_filters is None:
        return []
    # IronPython2 no siempre tiene "basestring" accesible, pero esto funciona en pyRevit normalmente.
    try:
        is_string = isinstance(name_filters, basestring)  # noqa: F821
    except:
        is_string = isinstance(name_filters, str)

    if is_string:
        name_filters = [name_filters]

    out = []
    for f in name_filters:
        try:
            s = str(f).strip()
        except:
            continue
        if s:
            out.append(s.lower())
    return out


def find_panel_by_name_filters(doc, name_filters):
    """
    Busca tableros (ElectricalEquipment) cuyo nombre de instancia o de tipo contenga
    alguno de los textos configurados en la lista.

    Devuelve:
      - panel: una instancia de tablero si SOLO hay un tipo coincidente en el proyecto; si hay varios tipos -> None
      - matches_by_type: dict con los tipos encontrados y sus datos (para notificar)
    """
    panels = (FilteredElementCollector(doc)
              .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
              .OfClass(FamilyInstance)
              .ToElements())

    filters = _normalize_name_filters(name_filters)
    matches_by_type = {}  # key: typeId.IntegerValue ; value: {'type_id','type_name','filters', 'panels'}

    for p in panels:
        inst_name = (getattr(p, "Name", "") or "")
        inst_low = inst_name.lower()

        type_elem = doc.GetElement(p.GetTypeId())
        if type_elem:
            p_name_param = type_elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            type_name = p_name_param.AsString() if p_name_param else ""
        else:
            type_name = ""
        type_low = (type_name or "").lower()

        matched = []
        for f in filters:
            if f and (f in inst_low or f in type_low):
                matched.append(f)

        if not matched:
            continue

        tid = p.GetTypeId()
        try:
            key = tid.IntegerValue
        except:
            key = str(tid)

        rec = matches_by_type.get(key)
        if not rec:
            rec = {"type_id": tid, "type_name": type_name, "filters": set(), "panels": []}
            matches_by_type[key] = rec

        for mf in matched:
            rec["filters"].add(mf)
        rec["panels"].append(p)

    if not matches_by_type:
        return None, matches_by_type

    # Si hay más de un TIPO coincidente, notificamos y NO elegimos automáticamente.
    if len(matches_by_type) > 1:
        return None, matches_by_type

    only = list(matches_by_type.values())[0]
    if only["panels"]:
        return only["panels"][0], matches_by_type

    return None, matches_by_type


def get_mechanical_equipments_filtered(doc):
    """
    Equipos mecánicos cuya FAMILIA (no el tipo) contenga
    'Evaporadores' o 'Mueble' en el nombre.
    """
    result = []
    mechs = (FilteredElementCollector(doc)
             .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
             .OfClass(FamilyInstance)
             .ToElements())

    for e in mechs:
        type_elem = doc.GetElement(e.GetTypeId())
        if not type_elem:
            continue

        fam = type_elem.Family
        fam_name = fam.Name if fam else ""
        low = (fam_name or "").lower()

        if "evaporadores" in low or "mueble" in low:
            result.append(e)

    return result


def get_electrical_connector(elem):
    """
    Devuelve el primer conector eléctrico del elemento.
    Primero probamos con DomainElectrical; si falla, probamos por ElectricalSystemType.
    """
    mep_model = getattr(elem, "MEPModel", None)
    if mep_model is None or mep_model.ConnectorManager is None:
        return None
    connectors = mep_model.ConnectorManager.Connectors
    if connectors is None:
        return None

    # 1) Intento por dominio
    for c in connectors:
        try:
            if c.Domain == Domain.DomainElectrical:
                return c
        except:
            continue

    # 2) Fallback: cualquier conector que exponga ElectricalSystemType
    for c in connectors:
        try:
            _ = c.ElectricalSystemType
            return c
        except:
            continue

    return None


def element_has_any_electrical_system(elem):
    """True si ya está en algún sistema eléctrico (para no duplicar circuitos)."""
    mep_model = getattr(elem, "MEPModel", None)
    if mep_model is None or mep_model.ConnectorManager is None:
        return False
    connectors = mep_model.ConnectorManager.Connectors
    if connectors is None:
        return False
    for c in connectors:
        try:
            esystems = list(c.ElectricalSystems)
        except:
            esystems = []
        if esystems:
            return True
    return False


# ----------------------------------------------------------------------
# SCRIPT PRINCIPAL
# ----------------------------------------------------------------------

panel, panel_matches = find_panel_by_name_filters(doc, PANEL_NAME_FILTER)

# 1) No hay coincidencias: no se encontró ningún tablero con los filtros configurados.
if not panel_matches:
    TaskDialog.Show(
        "Asignar tableros a equipos mecánicos",
        "No se encontró ningún TABLERO cuyo nombre (instancia o tipo) contenga alguno de estos textos:\n\n- {}".format(
            "\n- ".join([str(x) for x in PANEL_NAME_FILTER])
        )
    )

# 2) Hay más de un TIPO de tablero coincidente: notificamos y no continuamos (evita ambigüedad).
elif panel is None and len(panel_matches) > 1:
    lines = []
    for k, info in panel_matches.items():
        type_name = info.get("type_name", "") or "(sin nombre de tipo)"
        filters = sorted(list(info.get("filters", [])))
        count = len(info.get("panels", []))
        lines.append("- Tipo: '{}' | Instancias: {} | Coincide con: {}".format(
            type_name, count, ", ".join(filters) if filters else "(sin detalle)"
        ))

    TaskDialog.Show(
        "Asignar tableros a equipos mecánicos",
        "Se han encontrado VARIOS TIPOS de tablero que coinciden con la lista PANEL_NAME_FILTER.\n"
        "Para evitar asignaciones ambiguas, el script NO continuará.\n\n"
        "Tipos encontrados:\n{}\n\n"
        "Solución: ajusta PANEL_NAME_FILTER para que solo haya un tipo coincidente en el proyecto, "
        "o elimina/renombra el tipo que no corresponda.".format("\n".join(lines))
    )

# 3) Solo hay un tipo coincidente: continuamos como siempre.
else:
    mech_elems = get_mechanical_equipments_filtered(doc)

    if not mech_elems:
        TaskDialog.Show(
            "Asignar tableros a equipos mecánicos",
            "No se encontraron equipos mecánicos cuya FAMILIA contenga\n"
            "'Evaporadores' o 'Mueble'."
        )
    else:
        total = len(mech_elems)
        processed = 0
        created_count = 0
        skipped_in_system = 0
        skipped_no_conn = 0
        failed_create = 0
        failed_assign = 0
        first_create_error = ""
        first_assign_error = ""

        t = Transaction(doc, "Crear circuitos por equipo mecánico")

        try:
            t.Start()

            for e in mech_elems:
                processed += 1

                # No tocar si ya pertenece a algún sistema eléctrico
                if element_has_any_electrical_system(e):
                    skipped_in_system += 1
                    continue

                conn = get_electrical_connector(e)
                if conn is None:
                    skipped_no_conn += 1
                    continue

                # Tipo de sistema según el propio conector
                ckt_type = conn.ElectricalSystemType

                esys = None
                # Primer intento: tipo del conector
                try:
                    esys = ElectricalSystem.Create(conn, ckt_type)
                except Exception as ex:
                    # Segundo intento: forzar PowerCircuit
                    try:
                        esys = ElectricalSystem.Create(conn, ElectricalSystemType.PowerCircuit)
                    except Exception as ex2:
                        failed_create += 1
                        if not first_create_error:
                            first_create_error = str(ex2)
                        continue

                if not esys:
                    failed_create += 1
                    continue

                try:
                    esys.SelectPanel(panel)
                    created_count += 1
                except Exception as ex:
                    failed_assign += 1
                    if not first_assign_error:
                        first_assign_error = str(ex)
                    continue

            t.Commit()

        except Exception as ex:
            if t.HasStarted() and t.GetStatus() == TransactionStatus.Started:
                t.RollBack()
            TaskDialog.Show(
                "Error en transacción",
                "Se produjo una excepción durante la creación de circuitos:\n\n{}".format(ex)
            )
        else:
            msg = (
                "Equipos mecánicos filtrados: {total}\n"
                "Procesados: {processed}\n\n"
                "Circuitos creados y asignados al tablero: {created}\n"
                "Saltados (ya estaban en algún sistema eléctrico): {skipped_sys}\n"
                "Saltados (sin conector eléctrico): {skipped_no_conn}\n"
                "Fallos al CREAR circuito: {failed_create}\n"
                "Fallos al ASIGNAR al tablero: {failed_assign}"
            ).format(
                total=total,
                processed=processed,
                created=created_count,
                skipped_sys=skipped_in_system,
                skipped_no_conn=skipped_no_conn,
                failed_create=failed_create,
                failed_assign=failed_assign
            )

            extra = ""
            if first_create_error:
                extra += "\n\nPrimer error al CREAR circuito:\n{}".format(first_create_error)
            if first_assign_error:
                extra += "\n\nPrimer error al ASIGNAR a tablero:\n{}".format(first_assign_error)

            TaskDialog.Show("Asignar tableros a equipos mecánicos", msg + extra)
