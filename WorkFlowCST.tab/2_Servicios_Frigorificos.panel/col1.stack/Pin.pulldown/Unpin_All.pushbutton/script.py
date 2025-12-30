# -*- coding: utf-8 -*-

__title__ = "Unpin All" 
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
Date    = 10.12.2025
_____________________________________________________________________
Description:

Unpin all model elements
_____________________________________________________________________
How-to:

Just run the script to unpin all
_____________________________________________________________________
Last update: 10.12.2025
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
# Unpinnear elementos
# ==================================================
t = Transaction(doc, "Unpinnear elementos en vista activa")
t.Start()

contador = 0

for el in elementos:
    try:
        # Solo los que se pueden unpinnear y estén pinneados
        if hasattr(el, "Pinned"):
            if el.Pinned:
                el.Pinned = False
                contador += 1
    except:
        # Por si algún elemento raro da error
        pass

t.Commit()

# ==================================================
# Mensaje final
# ==================================================
TaskDialog.Show(
    "Unpinnear elementos",
    "Se despinnearon {} elementos en la vista activa.".format(contador)
)
