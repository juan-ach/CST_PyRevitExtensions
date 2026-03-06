# -*- coding: utf-8 -*-

__title__ = "Delete Evap. Tags"
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
_____________________________________________________________________
Description:

Delete ONLY the Mechanical Equipment tags created with:
Family pattern: CST_rectangulo informativo_v1x_(catalan|castellano), x=10..25
Types:  all types found in the detected family

Works on ACTIVE VIEW only.
_____________________________________________________________________
How-to:

Just run the script to erase those tags in the active view.
_____________________________________________________________________

Author: Juan Manuel Achenbach Anguita & ChatGPT
"""

from pyrevit import revit
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
import re

doc = revit.doc
view = doc.ActiveView

# ==================================================
# CONFIG
# ==================================================
PATRON_FAMILIA_TAG = re.compile(
    r"^CST_rectangulo informativo_v(\d+)_(catalan|castellano)$"
)
VERSION_MIN = 10
VERSION_MAX = 25
PRIORIDAD_VARIANTES = ["catalan", "castellano"]
NOMBRE_FAMILIA_TAG = None

# ==================================================
# Detectar familia de tags disponible (mismo criterio que Create)
# ==================================================
mejor_familia_por_variante = {
    "catalan": {"name": None, "version": -1},
    "castellano": {"name": None, "version": -1}
}

for fs in FilteredElementCollector(doc) \
        .OfClass(FamilySymbol) \
        .OfCategory(BuiltInCategory.OST_MechanicalEquipmentTags):

    match = PATRON_FAMILIA_TAG.match(fs.FamilyName)
    if not match:
        continue

    version = int(match.group(1))
    variante = match.group(2)

    if version < VERSION_MIN or version > VERSION_MAX:
        continue

    if version > mejor_familia_por_variante[variante]["version"]:
        mejor_familia_por_variante[variante]["version"] = version
        mejor_familia_por_variante[variante]["name"] = fs.FamilyName

for variante in PRIORIDAD_VARIANTES:
    candidata = mejor_familia_por_variante[variante]["name"]
    if candidata:
        NOMBRE_FAMILIA_TAG = candidata
        break

if not NOMBRE_FAMILIA_TAG:
    raise Exception(
        u"No se encontro ninguna familia de etiquetas con patron "
        u"'CST_rectangulo informativo_vXX_(catalan|castellano)' "
        u"en el rango v{}..v{}.".format(
            VERSION_MIN, VERSION_MAX
        )
    )

# ==================================================
# Recolectar tags de equipos mecánicos en la vista activa
# ==================================================
tags = FilteredElementCollector(doc, view.Id) \
    .OfCategory(BuiltInCategory.OST_MechanicalEquipmentTags) \
    .WhereElementIsNotElementType() \
    .ToElements()

tags_para_borrar = []

for tag in tags:
    try:
        type_id = tag.GetTypeId()
        if type_id == ElementId.InvalidElementId:
            continue

        tag_type = doc.GetElement(type_id)  # FamilySymbol normalmente
        if not tag_type:
            continue

        # Coincidencia por familia (todos los tipos)
        if tag_type.FamilyName == NOMBRE_FAMILIA_TAG:
            tags_para_borrar.append(tag)

    except:
        pass

# ==================================================
# Borrado
# ==================================================
t = Transaction(doc, "Borrar tags evaporadores (vista activa)")
t.Start()

contador = 0
for tag in tags_para_borrar:
    try:
        doc.Delete(tag.Id)
        contador += 1
    except:
        pass

t.Commit()

# ==================================================
# Mensaje final
# ==================================================
TaskDialog.Show(
    "Borrar Etiquetas",
    "Se borraron {} etiquetas en la vista activa.\n\nFamilia: {}\nTipos: todos".format(
        contador, NOMBRE_FAMILIA_TAG
    )
)
