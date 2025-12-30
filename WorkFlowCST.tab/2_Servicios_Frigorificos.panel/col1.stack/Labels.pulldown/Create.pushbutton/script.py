# -*- coding: utf-8 -*-

__title__ = "Create SERV Labels" 
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
Date    = 10.12.2025
_____________________________________________________________________
Description:

Create Tags for all Services

_____________________________________________________________________
How-to:

1.- Pin al elements with Pin tool
2.- Run this script
3.- Click on "Select pinned elements" to only select created tag and place them correctly
_____________________________________________________________________
Last update: 10.12.2025
_____________________________________________________________________
Author: Juan Manuel Achenbach Anguita & ChatGPT"""

from pyrevit import revit
from Autodesk.Revit.DB import *
from System.Windows.Forms import Form, Label, Timer
import System.Drawing

doc = revit.doc
view = doc.ActiveView

# ==================================================
# Popup con autocierre
# ==================================================
class AutoClosePopup(Form):
    def __init__(self, message, duration_ms=2000):
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
NOMBRE_FAMILIA_TAG = "CST_TAG Equipos Mecanicos v5"

TIPO_TAG_BASE = "mueble modulos"
TIPO_TAG_EVAP = "Evaporador"

# Offsets en pies (SEPARADOS POR TIPO DE TAG)
# --- Etiquetas "mueble modulos" ---
OFFSET_MUEBLE_X = 0.0 * 3.28084
OFFSET_MUEBLE_Y = -0.9 * 3.28084
OFFSET_MUEBLE_Z = 0.0

# --- Etiquetas "Evaporador" ---
OFFSET_EVAP_X = 0.0 * 3.28084
OFFSET_EVAP_Y = -0.5 * 3.28084
OFFSET_EVAP_Z = 0.0

# ==================================================
# Función: calcular punto del tag
# ==================================================
def calcular_punto_tag(equipo, dx=0, dy=0, dz=0):
    loc = equipo.Location
    if not isinstance(loc, LocationPoint):
        return None
    return loc.Point + XYZ(dx, dy, dz)

# ==================================================
# Buscar tipos de etiquetas (IronPython-safe)
# ==================================================
tag_symbol_base = None
tag_symbol_evap = None

for fs in FilteredElementCollector(doc) \
        .OfClass(FamilySymbol) \
        .OfCategory(BuiltInCategory.OST_MechanicalEquipmentTags):

    tipo_nombre = fs.get_Parameter(
        BuiltInParameter.SYMBOL_NAME_PARAM
    ).AsString()

    if fs.FamilyName == NOMBRE_FAMILIA_TAG:
        if tipo_nombre == TIPO_TAG_BASE:
            tag_symbol_base = fs
        elif tipo_nombre == TIPO_TAG_EVAP:
            tag_symbol_evap = fs

if not tag_symbol_base or not tag_symbol_evap:
    raise Exception(
        u"No se encontraron todos los tipos de etiqueta requeridos en la familia '{}'.".format(
            NOMBRE_FAMILIA_TAG)
    )

# Activar símbolos si es necesario
t = Transaction(doc, "Activar tipos de etiquetas")
t.Start()
if not tag_symbol_base.IsActive:
    tag_symbol_base.Activate()
if not tag_symbol_evap.IsActive:
    tag_symbol_evap.Activate()
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
# ==================================================
tags_existentes = FilteredElementCollector(doc, view.Id) \
    .OfCategory(BuiltInCategory.OST_MechanicalEquipmentTags) \
    .WhereElementIsNotElementType() \
    .ToElements()

equipos_ya_etiquetados = set()

for tag in tags_existentes:
    try:
        for eid in tag.GetTaggedElementIds():
            if eid != ElementId.InvalidElementId:
                equipos_ya_etiquetados.add(eid)
    except:
        pass

# ==================================================
# CREACIÓN DE TAGS CON LÓGICA CONDICIONAL
# ==================================================
contador = 0

t = Transaction(doc, u"Etiquetar equipos mecánicos según nombre")
t.Start()

for eq in equipos:
    if eq.Id in equipos_ya_etiquetados:
        continue

    # Nombre de la FAMILIA del equipo (IronPython-safe)
    familia_eq = eq.Symbol.FamilyName
    if not familia_eq:
        continue

    familia_eq_lower = familia_eq.lower()

    # Selección del tipo de etiqueta según nombre de FAMILIA
    tag_symbol = None
    dx = dy = dz = 0.0

    # "mueble" o "UCond" -> etiqueta base "mueble modulos"
    if "mueble" in familia_eq_lower or "ucond" in familia_eq_lower:
        tag_symbol = tag_symbol_base
        dx, dy, dz = OFFSET_MUEBLE_X, OFFSET_MUEBLE_Y, OFFSET_MUEBLE_Z

    # "evap" -> etiqueta "Evaporador"
    elif "evap" in familia_eq_lower:
        tag_symbol = tag_symbol_evap
        dx, dy, dz = OFFSET_EVAP_X, OFFSET_EVAP_Y, OFFSET_EVAP_Z

    else:
        continue  # No cumple ninguna condición

    punto_tag = calcular_punto_tag(eq, dx, dy, dz)
    if not punto_tag:
        continue

    try:
        tag = IndependentTag.Create(
            doc,
            view.Id,
            Reference(eq),
            False,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            punto_tag
        )

        tag.ChangeTypeId(tag_symbol.Id)
        contador += 1

    except:
        pass

t.Commit()

# ==================================================
# Popup final
# ==================================================
popup = AutoClosePopup(
    u"Se etiquetaron {} equipos mecánicos según su nombre.".format(contador)
)
popup.ShowDialog()
