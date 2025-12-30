# -*- coding: utf-8 -*-

__title__ = "Create ELE Labels" 
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
Date    = 15.12.2025
_____________________________________________________________________
Description:

Create Tags for all Electrical Circuits on Mechabnical Equipement

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
        self.Width = 420
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
NOMBRE_TIPO_TAG = "ubicación circuito"

OFFSET_X = 0.0 * 3.28084
OFFSET_Y = -0.7 * 3.28084
OFFSET_Z = 0.0

# ==================================================
# Utilidades
# ==================================================
def calcular_punto_tag(equipo, dx, dy, dz):
    loc = equipo.Location
    if not isinstance(loc, LocationPoint):
        return None
    return loc.Point + XYZ(dx, dy, dz)

# ==================================================
# Buscar el tipo de etiqueta (SIN depender de categoría)
# ==================================================
tag_symbol = None

for fs in FilteredElementCollector(doc).OfClass(FamilySymbol):
    if fs.FamilyName != NOMBRE_FAMILIA_TAG:
        continue

    param_name = fs.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    if not param_name:
        continue

    if param_name.AsString() == NOMBRE_TIPO_TAG:
        tag_symbol = fs
        break

if not tag_symbol:
    raise Exception(
        u"No se encontró el tipo de etiqueta '{}' en la familia '{}'.".format(
            NOMBRE_TIPO_TAG, NOMBRE_FAMILIA_TAG
        )
    )

# ==================================================
# Activar el tipo de etiqueta
# ==================================================
t = Transaction(doc, "Activar tipo de etiqueta")
t.Start()
if not tag_symbol.IsActive:
    tag_symbol.Activate()
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
    .OfClass(IndependentTag) \
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
# CREACIÓN DE TAGS
# ==================================================
contador = 0

t = Transaction(doc, u"Etiquetar equipos mecánicos - ubicación circuito")
t.Start()

for eq in equipos:
    if eq.Id in equipos_ya_etiquetados:
        continue

    punto_tag = calcular_punto_tag(eq, OFFSET_X, OFFSET_Y, OFFSET_Z)
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
    u"Se crearon {} etiquetas de 'ubicación circuito'.".format(contador)
)
popup.ShowDialog()
