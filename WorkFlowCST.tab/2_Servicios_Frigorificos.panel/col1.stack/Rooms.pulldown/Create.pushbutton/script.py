# -*- coding: utf-8 -*-

from pyrevit import revit
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import Architecture, SpatialElementTag
from Autodesk.Revit.UI import TaskDialog
import unicodedata

doc = revit.doc
view = doc.ActiveView


# ==================================================
# CONFIG: si quieres SOLO analizar y NO borrar, pon False
# ==================================================
BORRAR_ROOMS_NO_FRIGORIFICO = True

# Config de la etiqueta de habitación
NOMBRE_FAMILIA_TAG_ROOM = "CST_TAG de habitación v4"
TIPO_TAG_ROOM = "Superficie Nombre y Altura"


# ==================================================
# Helpers
# ==================================================
def normalizar(texto):
    """Minúsculas y sin tildes."""
    if not texto:
        return u""
    texto_norm = unicodedata.normalize('NFD', texto)
    return u"".join(c for c in texto_norm if unicodedata.category(c) != 'Mn').lower()


def nombre_tipo_muro(elem, doc):
    """
    Si elem es algo de categoría Muros, devuelve el nombre de su ElementType.
    Si no, devuelve None.
    """
    if not elem or not elem.Category:
        return None

    if elem.Category.Id.IntegerValue != int(BuiltInCategory.OST_Walls):
        return None

    try:
        type_id = elem.GetTypeId()
        type_elem = doc.GetElement(type_id)
        if not type_elem:
            return None

        # Intentar .Name primero
        name_attr = getattr(type_elem, "Name", None)
        if name_attr:
            return name_attr

        # Plan B: parámetro de nombre de tipo
        p = type_elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p and p.AsString():
            return p.AsString()

    except:
        return None

    return None


# ==================================================
# Comprobación de vista
# ==================================================
if not isinstance(view, ViewPlan):
    TaskDialog.Show("Crear/Filtrar Rooms", "La vista activa debe ser una vista de planta.")
else:
    nivel = view.GenLevel
    if not nivel:
        TaskDialog.Show("Crear/Filtrar Rooms", "La vista activa no tiene un nivel asociado.")
    else:
        # ==================================================
        # 1) CREAR HABITACIONES AUTOMÁTICAMENTE EN EL NIVEL
        #    (equivalente a 'Colocar habitaciones automáticamente')
        # ==================================================
        t_crear = Transaction(doc, "Crear rooms automáticamente en nivel")
        t_crear.Start()

        try:
            room_ids_creados = doc.Create.NewRooms2(nivel)
        except Exception as ex:
            t_crear.RollBack()
            TaskDialog.Show("Crear/Filtrar Rooms",
                            "Error creando rooms automáticamente:\n{}".format(ex))
            raise

        t_crear.Commit()

        # ==================================================
        # 2) ANALIZAR Y BORRAR LAS QUE NO CUMPLAN CONDICIÓN
        # ==================================================
        boundary_options = SpatialElementBoundaryOptions()
        boundary_options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish

        # Rooms visibles en la vista actual
        rooms = FilteredElementCollector(doc, view.Id) \
            .OfCategory(BuiltInCategory.OST_Rooms) \
            .WhereElementIsNotElementType() \
            .ToElements()

        total_rooms = len(rooms)
        rooms_ok = 0
        rooms_borradas = 0
        tipos_muro_detectados = set()

        t_filtrar = Transaction(doc, "Filtrar Rooms por muros tipo 'Frigorifico'")
        t_filtrar.Start()

        for room in rooms:
            # Aseguramos mismo nivel
            if hasattr(room, "LevelId") and room.LevelId != nivel.Id:
                continue

            try:
                blists = room.GetBoundarySegments(boundary_options)
            except:
                blists = None

            if not blists:
                # Sin contorno válido -> se considera no válida
                if BORRAR_ROOMS_NO_FRIGORIFICO:
                    try:
                        doc.Delete(room.Id)
                        rooms_borradas += 1
                    except:
                        pass
                continue

            tiene_muros = False
            muro_no_frigo = False

            for seg_list in blists:
                for seg in seg_list:
                    elem_id = seg.ElementId
                    if elem_id == ElementId.InvalidElementId:
                        continue

                    elem = doc.GetElement(elem_id)
                    tipo_nombre = nombre_tipo_muro(elem, doc)

                    if tipo_nombre is None:
                        # No es muro, lo ignoramos
                        continue

                    tiene_muros = True
                    tipos_muro_detectados.add(tipo_nombre)

                    # Condición: el nombre de tipo debe contener 'Frigorifico'
                    if "frigorifico" not in normalizar(tipo_nombre):
                        muro_no_frigo = True
                        break

                if muro_no_frigo:
                    break

            if tiene_muros and not muro_no_frigo:
                rooms_ok += 1
            else:
                if BORRAR_ROOMS_NO_FRIGORIFICO:
                    try:
                        doc.Delete(room.Id)
                        rooms_borradas += 1
                    except:
                        pass

        t_filtrar.Commit()

        # ==================================================
        # 3) ETIQUETAR LAS HABITACIONES RESTANTES
        # ==================================================

        # 3.1 Buscar el tipo de etiqueta de habitación
        tag_room_symbol = None

        for fs in FilteredElementCollector(doc) \
                .OfClass(FamilySymbol) \
                .OfCategory(BuiltInCategory.OST_RoomTags):

            fam_name = fs.FamilyName
            tipo_nombre = fs.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()

            if fam_name == NOMBRE_FAMILIA_TAG_ROOM and tipo_nombre == TIPO_TAG_ROOM:
                tag_room_symbol = fs
                break

        if not tag_room_symbol:
            TaskDialog.Show(
                "Etiquetar Rooms",
                u"No se encontró el tipo de etiqueta '{}' en la familia '{}'.".format(
                    TIPO_TAG_ROOM, NOMBRE_FAMILIA_TAG_ROOM
                )
            )
        else:
            # Activar el símbolo si hace falta
            t_act = Transaction(doc, "Activar tipo de etiqueta de habitación")
            t_act.Start()
            if not tag_room_symbol.IsActive:
                tag_room_symbol.Activate()
            t_act.Commit()

            # 3.2 Obtener rooms y tags existentes en la vista
            rooms_en_vista = FilteredElementCollector(doc, view.Id) \
                .OfCategory(BuiltInCategory.OST_Rooms) \
                .WhereElementIsNotElementType() \
                .ToElements()

            tags_existentes = FilteredElementCollector(doc, view.Id) \
                .OfCategory(BuiltInCategory.OST_RoomTags) \
                .WhereElementIsNotElementType() \
                .ToElements()

            rooms_ya_etiquetadas = set()

            for tag in tags_existentes:
                try:
                    if isinstance(tag, SpatialElementTag):
                        r = tag.Room
                        if r:
                            rooms_ya_etiquetadas.add(r.Id)
                except:
                    pass

            # 3.3 Crear etiquetas para rooms sin tag
            t_tag = Transaction(doc, "Etiquetar habitaciones")
            t_tag.Start()

            etiquetas_creadas = 0

            for room in rooms_en_vista:
                if room.Id in rooms_ya_etiquetadas:
                    continue
                if room.Area <= 0:
                    continue
                if not room.Location:
                    continue

                loc = room.Location
                if not isinstance(loc, LocationPoint):
                    continue

                pt = loc.Point
                uv = UV(pt.X, pt.Y)

                try:
                    new_tag = doc.Create.NewRoomTag(LinkElementId(room.Id), uv, view.Id)
                    new_tag.ChangeTypeId(tag_room_symbol.Id)
                    etiquetas_creadas += 1
                except:
                    pass

            t_tag.Commit()

        # ==================================================
        # 4) Mensaje final
        # ==================================================
        tipos_txt = "\n".join(sorted(tipos_muro_detectados)) if tipos_muro_detectados else u"Ninguno"

        mensaje = (
            u"Rooms creadas automáticamente en la vista/nivel: {tot}\n"
            u"Rooms que se conservan "
            u"(rodeadas SOLO por muros cuyo tipo contiene 'Frigorifico'): {ok}\n"
            u"Rooms borradas por no cumplir condición: {bor}\n\n"
            u"Tipos de muro detectados como límites (ElementType.Name):\n{tipos}"
        ).format(
            tot=total_rooms,
            ok=rooms_ok,
            bor=rooms_borradas if BORRAR_ROOMS_NO_FRIGORIFICO else 0,
            tipos=tipos_txt
        )

        TaskDialog.Show("Crear/Filtrar Rooms", mensaje)
