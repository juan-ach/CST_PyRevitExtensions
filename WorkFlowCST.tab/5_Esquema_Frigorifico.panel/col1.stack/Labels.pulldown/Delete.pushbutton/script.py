# -*- coding: utf-8 -*-

__title__ = "Delete Evaporadores Tags"
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
_____________________________________________________________________
Description:

Delete ONLY the Mechanical Equipment tags created with:
Family: CST_rectangulo informativo_v14_catalan
Type:   evaporadores

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

doc = revit.doc
view = doc.ActiveView

# ==================================================
# CONFIG
# ==================================================
NOMBRE_FAMILIA_TAG = "CST_rectangulo informativo_v14_catalan"
TIPO_TAG = "evaporadores"

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

        # Nombre de tipo (IronPython-safe)
        tipo_nombre = tag_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()

        # Coincidencia por familia + tipo
        if tag_type.FamilyName == NOMBRE_FAMILIA_TAG and tipo_nombre == TIPO_TAG:
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
    "Se borraron {} etiquetas en la vista activa.\n\nFamilia: {}\nTipo: {}".format(
        contador, NOMBRE_FAMILIA_TAG, TIPO_TAG
    )
)
