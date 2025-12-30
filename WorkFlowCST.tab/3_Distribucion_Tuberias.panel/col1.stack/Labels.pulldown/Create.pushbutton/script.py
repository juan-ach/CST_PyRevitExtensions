# -*- coding: utf-8 -*-

__title__ = "Create TUB Labels" 
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
Date    = 15.12.2025
_____________________________________________________________________
Description:

Create Tags for all Services & Pipe Diameter

_____________________________________________________________________
How-to:

1.- Pin al elements with Pin tool
2.- Run this script
3.- Click on "Select pinned elements" to only select created tag and place them correctly
_____________________________________________________________________
Last update: 15.12.2025
_____________________________________________________________________
Author: Juan Manuel Achenbach Anguita & ChatGPT"""

from pyrevit import revit
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Mechanical import MechanicalEquipment
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB import FamilyInstance, BuiltInCategory, Transaction

from System.Windows.Forms import Form, Label, Timer
import System.Drawing

doc = revit.doc
view = doc.ActiveView


# ==================================================
# Popup con autocierre
# ==================================================
class AutoClosePopup(Form):
    def __init__(self, message, duration_ms=2500):
        self.Text = "Info"
        self.Width = 380
        self.Height = 140
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen

        label = Label()
        label.Text = message
        label.Dock = System.Windows.Forms.DockStyle.Fill
        label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        label.Font = System.Drawing.Font("Arial", 11)
        self.Controls.Add(label)

        timer = Timer()
        timer.Interval = duration_ms
        timer.Tick += self.close_popup
        timer.Start()

    def close_popup(self, sender, args):
        self.Close()


# ==================================================
# CONFIGURACIÓN
# ==================================================

# ---------- TAGS DE EQUIPO MECÁNICO ----------
NOMBRE_FAMILIA_TAG_EQ = "CST_TAG Equipos Mecanicos v5"
TIPO_TAG_EQ_MUEBLE = "mueble modulos"
TIPO_TAG_EQ_EVAP = "Evaporador"

# Offsets en pies para la etiqueta DEL EQUIPO (relativos al punto del equipo)
OFFSET_EQ_MUEBLE_X = 0.0 * 3.28084
OFFSET_EQ_MUEBLE_Y = -0.9 * 3.28084
OFFSET_EQ_MUEBLE_Z = 0.0

OFFSET_EQ_EVAP_X   = 0.0 * 3.28084
OFFSET_EQ_EVAP_Y   = -1 * 3.28084
OFFSET_EQ_EVAP_Z   = 0.0

# Condiciones de búsqueda en el nombre (se buscará en Family, Tipo y Nombre)
COND_MUEBLE = "mueble"
COND_UCOND  = "ucond"
COND_EVAP   = "evap"


# ---------- TAGS DE TUBERÍA (UN SOLO TIPO, PERO OFFSETS POR GRUPO) ----------
NOMBRE_FAMILIA_TAG_TUBO = "CST_TAG Diametro Tubería v28"   # <-- ajusta si hace falta
TIPO_TAG_TUBO = "1.5"                                      # <-- ajusta si hace falta

# Offsets en pies para las etiquetas de tubería,
# RELATIVOS a la posición de la etiqueta del EQUIPO,
# diferenciados por tipo de equipo ("mueble", "ucond", "evap").

# --- Equipos tipo "mueble" ---
OFFSET_TUBO_MUEBLE_X = -0.7 * 3.28084
OFFSET_TUBO_MUEBLE_Y = -1.1 * 3.28084
OFFSET_TUBO_MUEBLE_Z = 0.0
STEP_TUBO_MUEBLE_X   = 1 * 3.28084   # separación entre tags en X

# --- Equipos tipo "ucond" ---
OFFSET_TUBO_UCOND_X = -0.7 * 3.28084
OFFSET_TUBO_UCOND_Y = -0.3 * 3.28084
OFFSET_TUBO_UCOND_Z = 0.0
STEP_TUBO_UCOND_X   = 1 * 3.28084   # puedes cambiarlo si quieres otro patrón

# --- Equipos tipo "evap" ---
OFFSET_TUBO_EVAP_X = -0.35 * 3.28084
OFFSET_TUBO_EVAP_Y = -0.3 * 3.28084
OFFSET_TUBO_EVAP_Z = 0.0
STEP_TUBO_EVAP_X   = 0.85 * 3.28084


# ==================================================
# Funciones auxiliares
# ==================================================
def calcular_punto_tag_desde_punto(base_point, dx=0, dy=0, dz=0):
    """Devuelve un XYZ a partir de un XYZ base más el offset."""
    return base_point + XYZ(dx, dy, dz)


def obtener_punto_equipo(eq):
    loc = eq.Location
    if not isinstance(loc, LocationPoint):
        return None
    return loc.Point


def obtener_tuberia_y_tipo_sistema(conector):
    """
    Devuelve (Pipe, nombre_sistema) buscando:
        - Conexión directa a tubería.
        - O tubería detrás de un fitting / accesorio intermedio.
    """
    refs = conector.AllRefs
    for r in refs:
        owner = r.Owner

        # Caso directo
        if isinstance(owner, Pipe):
            sistema = owner.MEPSystem.Name if owner.MEPSystem else ""
            return owner, sistema

        # Caso con fitting / accesorio intermedio
        if isinstance(owner, FamilyInstance):
            mepmodel = getattr(owner, "MEPModel", None)
            if mepmodel and hasattr(mepmodel, "ConnectorManager") and mepmodel.ConnectorManager:
                for cf in mepmodel.ConnectorManager.Connectors:
                    for rr in cf.AllRefs:
                        other_owner = rr.Owner
                        if isinstance(other_owner, Pipe):
                            sistema = other_owner.MEPSystem.Name if other_owner.MEPSystem else ""
                            return other_owner, sistema
    return None, None


# ==================================================
# Buscar tipos de etiquetas
# ==================================================
# --- Tags de equipo mecánico ---
tag_eq_mueble = None
tag_eq_evap = None

for fs in FilteredElementCollector(doc) \
        .OfClass(FamilySymbol) \
        .OfCategory(BuiltInCategory.OST_MechanicalEquipmentTags):

    tipo_nombre = fs.get_Parameter(
        BuiltInParameter.SYMBOL_NAME_PARAM
    ).AsString()

    if fs.FamilyName == NOMBRE_FAMILIA_TAG_EQ:
        if tipo_nombre == TIPO_TAG_EQ_MUEBLE:
            tag_eq_mueble = fs
        elif tipo_nombre == TIPO_TAG_EQ_EVAP:
            tag_eq_evap = fs

if not tag_eq_mueble or not tag_eq_evap:
    raise Exception(
        u"No se encontraron todos los tipos de etiqueta de EQUIPO requeridos en la familia '{}'.".format(
            NOMBRE_FAMILIA_TAG_EQ)
    )

# --- Tag de TUBERÍA (único tipo) ---
tag_tubo = None

for fs in FilteredElementCollector(doc) \
        .OfClass(FamilySymbol) \
        .OfCategory(BuiltInCategory.OST_PipeTags):

    tipo_nombre = fs.get_Parameter(
        BuiltInParameter.SYMBOL_NAME_PARAM
    ).AsString()

    if fs.FamilyName == NOMBRE_FAMILIA_TAG_TUBO and tipo_nombre == TIPO_TAG_TUBO:
        tag_tubo = fs

if not tag_tubo:
    raise Exception(
        u"No se encontró el tipo de etiqueta de TUBERÍA '{}' en la familia '{}'.".format(
            TIPO_TAG_TUBO, NOMBRE_FAMILIA_TAG_TUBO)
    )


# Activar símbolos si es necesario
t = Transaction(doc, "Activar tipos de etiquetas")
t.Start()
for sym in (tag_eq_mueble, tag_eq_evap, tag_tubo):
    if sym and not sym.IsActive:
        sym.Activate()
t.Commit()


# ==================================================
# Equipos mecánicos en la vista activa
# ==================================================
equipos = FilteredElementCollector(doc, view.Id) \
    .OfCategory(BuiltInCategory.OST_MechanicalEquipment) \
    .WhereElementIsNotElementType() \
    .ToElements()


# ==================================================
# Tags existentes en la vista
#   - Guardamos posición de la etiqueta de cada equipo (si ya tenía)
#   - Guardamos qué tuberías ya están etiquetadas
# ==================================================
tags_eq_existentes = FilteredElementCollector(doc, view.Id) \
    .OfCategory(BuiltInCategory.OST_MechanicalEquipmentTags) \
    .WhereElementIsNotElementType() \
    .ToElements()

tags_tubo_existentes = FilteredElementCollector(doc, view.Id) \
    .OfCategory(BuiltInCategory.OST_PipeTags) \
    .WhereElementIsNotElementType() \
    .ToElements()

punto_tag_equipo_por_id = {}   # ElementId de equipo -> XYZ posición de la tag
tuberias_ya_etiquetadas = set()


# --- Tags de equipo ya existentes ---
for tag in tags_eq_existentes:
    try:
        tagged_ids = tag.GetTaggedElementIds()
        for eid in tagged_ids:
            if eid != ElementId.InvalidElementId:
                try:
                    punto_tag_equipo_por_id[eid] = tag.TagHeadPosition
                except:
                    pass
    except:
        pass

# --- Tags de tubería ya existentes ---
for tag in tags_tubo_existentes:
    try:
        for eid in tag.GetTaggedElementIds():
            if eid != ElementId.InvalidElementId:
                tuberias_ya_etiquetadas.add(eid)
    except:
        pass


# ==================================================
# PROCESO PRINCIPAL: crear tags de equipo + tags de tubería
# ==================================================
contador_eq_nuevas = 0
contador_tub_nuevas = 0

t = Transaction(doc, u"Etiquetar equipos y tramos de tubería")
t.Start()

for eq in equipos:
    # ----------------------------------------
    # 1) ETIQUETA DEL EQUIPO
    # ----------------------------------------
    punto_equipo = obtener_punto_equipo(eq)
    if not punto_equipo:
        continue

    family_name = ""
    type_name = ""
    if eq.Symbol:
        family_name = eq.Symbol.FamilyName or ""
        p_type_name = eq.Symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        type_name = p_type_name.AsString() if p_type_name else ""

    elem_name = eq.Name or ""

    # Texto combinado para buscar "mueble", "ucond", "evap"
    search_text = u"{} {} {}".format(family_name, type_name, elem_name).lower()

    # Selección del tipo de etiqueta de EQUIPO + offset
    tag_symbol_eq = None
    dx_eq = dy_eq = dz_eq = 0.0

    # Variable para saber qué grupo de offsets de TUBERÍA usar
    grupo_offsets_tubo = None  # "mueble", "ucond" o "evap"

    # "mueble" -> etiqueta base "mueble modulos"
    if COND_MUEBLE in search_text:
        tag_symbol_eq = tag_eq_mueble
        dx_eq, dy_eq, dz_eq = OFFSET_EQ_MUEBLE_X, OFFSET_EQ_MUEBLE_Y, OFFSET_EQ_MUEBLE_Z
        grupo_offsets_tubo = "mueble"

    # "ucond" -> misma etiqueta de equipo, pero grupo de offsets de tubería distinto
    elif COND_UCOND in search_text:
        tag_symbol_eq = tag_eq_mueble
        dx_eq, dy_eq, dz_eq = OFFSET_EQ_MUEBLE_X, OFFSET_EQ_MUEBLE_Y, OFFSET_EQ_MUEBLE_Z
        grupo_offsets_tubo = "ucond"

    # "evap" -> etiqueta "Evaporador"
    elif COND_EVAP in search_text:
        tag_symbol_eq = tag_eq_evap
        dx_eq, dy_eq, dz_eq = OFFSET_EQ_EVAP_X, OFFSET_EQ_EVAP_Y, OFFSET_EQ_EVAP_Z
        grupo_offsets_tubo = "evap"

    else:
        # Equipo que no cumple ninguna condición: lo ignoramos por completo
        continue

    # Punto de la etiqueta del equipo
    if eq.Id in punto_tag_equipo_por_id:
        # Ya tenía tag: usamos su posición existente como referencia
        punto_tag_eq = punto_tag_equipo_por_id[eq.Id]
    else:
        # No tenía tag: creamos una nueva
        punto_tag_eq = calcular_punto_tag_desde_punto(punto_equipo, dx_eq, dy_eq, dz_eq)

        try:
            eq_tag = IndependentTag.Create(
                doc,
                view.Id,
                Reference(eq),
                False,
                TagMode.TM_ADDBY_CATEGORY,
                TagOrientation.Horizontal,
                punto_tag_eq
            )
            eq_tag.ChangeTypeId(tag_symbol_eq.Id)
            contador_eq_nuevas += 1
            # Guardamos la posición de la nueva tag
            punto_tag_equipo_por_id[eq.Id] = punto_tag_eq
        except:
            # Si falla la creación de la tag, no podremos referenciarla para las tuberías
            continue

    # ----------------------------------------
    # 2) ETIQUETAS EN TRAMOS DE TUBERÍA
    #     RELATIVAS A LA ETIQUETA DEL EQUIPO
    #     Aspiración primero, luego líquido en X
    # ----------------------------------------
    mepmodel = getattr(eq, "MEPModel", None)
    if not mepmodel or not hasattr(mepmodel, "ConnectorManager") or not mepmodel.ConnectorManager:
        continue

    # Seleccionar offsets base y paso según el grupo de equipo
    if grupo_offsets_tubo == "mueble":
        base_dx_tubo = OFFSET_TUBO_MUEBLE_X
        base_dy_tubo = OFFSET_TUBO_MUEBLE_Y
        base_dz_tubo = OFFSET_TUBO_MUEBLE_Z
        step_tubo_x  = STEP_TUBO_MUEBLE_X

    elif grupo_offsets_tubo == "ucond":
        base_dx_tubo = OFFSET_TUBO_UCOND_X
        base_dy_tubo = OFFSET_TUBO_UCOND_Y
        base_dz_tubo = OFFSET_TUBO_UCOND_Z
        step_tubo_x  = STEP_TUBO_UCOND_X

    elif grupo_offsets_tubo == "evap":
        base_dx_tubo = OFFSET_TUBO_EVAP_X
        base_dy_tubo = OFFSET_TUBO_EVAP_Y
        base_dz_tubo = OFFSET_TUBO_EVAP_Z
        step_tubo_x  = STEP_TUBO_EVAP_X

    else:
        # Por seguridad, aunque en principio nunca debería caer aquí
        base_dx_tubo = 0.0
        base_dy_tubo = 0.0
        base_dz_tubo = 0.0
        step_tubo_x  = 0.3 * 3.28084

    # Contadores para separar aspiración / líquido
    idx_A = 0   # aspiración
    idx_L = 0   # líquido

    for c in mepmodel.ConnectorManager.Connectors:
        tubo, sistema = obtener_tuberia_y_tipo_sistema(c)
        if not tubo:
            continue

        # Evitar etiquetas duplicadas sobre la misma tubería
        if tubo.Id in tuberias_ya_etiquetadas:
            continue

        sistema_upper = (sistema or "").upper()

        es_aspiracion = "A" in sistema_upper  # ajusta si tus nombres son distintos
        es_liquido    = "L" in sistema_upper

        # Si no lo identificamos como A o L, lo saltamos
        if not (es_aspiracion or es_liquido):
            continue

        # Cálculo de offset en X, según sea aspiración o líquido
        if es_aspiracion:
            # Aspiración: la primera en X
            dx_tub = base_dx_tubo + idx_A * step_tubo_x
            idx_A += 1
        else:
            # Líquido: más a la derecha que aspiración
            dx_tub = base_dx_tubo + step_tubo_x + idx_L * step_tubo_x
            idx_L += 1

        # Y y Z fijos según grupo
        dy_tub = base_dy_tubo
        dz_tub = base_dz_tubo

        punto_tag_tub = calcular_punto_tag_desde_punto(
            punto_tag_eq,
            dx_tub,
            dy_tub,
            dz_tub
        )

        try:
            tub_tag = IndependentTag.Create(
                doc,
                view.Id,
                Reference(tubo),
                False,
                TagMode.TM_ADDBY_CATEGORY,
                TagOrientation.Horizontal,
                punto_tag_tub
            )
            tub_tag.ChangeTypeId(tag_tubo.Id)
            contador_tub_nuevas += 1
            tuberias_ya_etiquetadas.add(tubo.Id)
        except:
            pass

t.Commit()


# ==================================================
# Popup final
# ==================================================
mensaje = u"Se crearon {} etiquetas de equipos y {} etiquetas de tramos de tubería conectados.".format(
    contador_eq_nuevas, contador_tub_nuevas
)

popup = AutoClosePopup(mensaje)
popup.ShowDialog()
