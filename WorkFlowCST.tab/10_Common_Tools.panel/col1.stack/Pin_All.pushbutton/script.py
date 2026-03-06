# -*- coding: utf-8 -*-

__title__ = "Pin All" 
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
_____________________________________________________________________
Description:

Pin all model elements
_____________________________________________________________________
How-to:

Just run the script to pin all
_____________________________________________________________________
Author: Juan Manuel Achenbach Anguita & ChatGPT"""


from pyrevit import revit
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog

doc = revit.doc
view = doc.ActiveView

# ==================================================
# Obtener todos los elementos de la vista activa
# (excluyendo tipos)
# ==================================================
elementos = FilteredElementCollector(doc, view.Id) \
    .WhereElementIsNotElementType() \
    .ToElements()

# ==================================================
# Pinnear elementos
# ==================================================
t = Transaction(doc, "Pinnear elementos en vista activa")
t.Start()

contador = 0

for el in elementos:
    try:
        # Solo los que se pueden pinnear y no estén ya pinneados
        if hasattr(el, "Pinned"):
            if not el.Pinned:
                el.Pinned = True
                contador += 1
    except:
        # Por si algún elemento raro da error al pinnear
        pass

t.Commit()

# ==================================================
# Mensaje final
# ==================================================
TaskDialog.Show(
    "Pinnear elementos",
    "Se pinnearon {} elementos en la vista activa.".format(contador)
)
