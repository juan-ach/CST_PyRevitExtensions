# -*- coding: utf-8 -*-

__title__ = "Create Evaporadores Labels"
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.1'
__doc__ = """Version: 1.1
_____________________________________________________________________
Description:

Create the same Mechanical Equipment tag for ALL mechanical equipment
in the active view (skips already tagged elements).
_____________________________________________________________________
How-to:

1.- Pin all elements with Pin tool (optional)
2.- Run this script
_____________________________________________________________________
Author: Juan Manuel Achenbach Anguita & ChatGPT"""

from pyrevit import revit
from Autodesk.Revit.DB import *
from System.Windows.Forms import Form, Label, Timer, FormStartPosition, DockStyle
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
        self.Height = 150
        self.StartPosition = FormStartPosition.CenterScreen

        label = Label()
        label.Text = message
        label.Dock = DockStyle.Fill
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
NOMBRE_FAMILIA_TAG = "CST_rectangulo informativo_v14_catalan"
TIPO_TAG = "evaporadores"

# Offsets en metros -> convertidos a pies (Revit trabaja en pies)
# Ajusta a tu gusto:
OFFSET_X = 0.0 * 3.28084
OFFSET_Y = -0.5 * 3.28084
OFFSET_Z = 0.0

# ==================================================
# Función: calcular punto del tag (más robusto)
# ==================================================
def calcular_punto_tag(elem, dx=0, dy=0, dz=0):
    loc = elem.Location

    # Caso típico: LocationPoint
    if isinstance(loc, LocationPoint):
        return loc.Point + XYZ(dx, dy, dz)

    # Caso: LocationCurve (p.ej. algunos equipos "lineales")
    if isinstance(loc, LocationCurve) and loc.Curve:
        try:
            mid = loc.Curve.Evaluate(0.5, True)
            return mid + XYZ(dx, dy, dz)
        except:
            pass

    # Fallback: centro del bounding box (en la vista si existe)
    try:
        bb = elem.get_BoundingBox(view) or elem.get_BoundingBox(None)
        if bb:
            center = (bb.Min + bb.Max) * 0.5
            return center + XYZ(dx, dy, dz)
    except:
        pass

    return None

# ==================================================
# Buscar el tipo de etiqueta (IronPython-safe)
# ==================================================
tag_symbol = None

for fs in FilteredElementCollector(doc) \
        .OfClass(FamilySymbol) \
        .OfCategory(BuiltInCategory.OST_MechanicalEquipmentTags):

    tipo_nombre = fs.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()

    if fs.FamilyName == NOMBRE_FAMILIA_TAG and tipo_nombre == TIPO_TAG:
        tag_symbol = fs
        break

if not tag_symbol:
    raise Exception(
        u"No se encontró el tipo de etiqueta '{}' dentro de la familia '{}'.".format(
            TIPO_TAG, NOMBRE_FAMILIA_TAG
        )
    )

# Activar símbolo si es necesario
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
# Tags existentes en la vista (para no duplicar)
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
# CREACIÓN DE TAGS (UN SOLO TIPO PARA TODO)
# ==================================================
contador = 0

t = Transaction(doc, u"Etiquetar todos los equipos mecánicos (evaporadores)")
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
    u"Se etiquetaron {} equipos mecánicos con:\nFamilia: '{}'\nTipo: '{}'".format(
        contador, NOMBRE_FAMILIA_TAG, TIPO_TAG
    )
)
popup.ShowDialog()
