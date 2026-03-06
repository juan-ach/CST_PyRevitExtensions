# -*- coding: utf-8 -*-

__title__ = "Delete Rooms" 
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
_____________________________________________________________________
Description:

Delete rooms when bounded by isolation panels.
_____________________________________________________________________
How-to:

Just run the script to delete rooms, then check in a control table they were effectively deleted.

_____________________________________________________________________
Author: Juan Manuel Achenbach Anguita & ChatGPT"""

from pyrevit import revit
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog

doc = revit.doc

# Colectar TODAS las habitaciones del proyecto
rooms = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_Rooms) \
    .WhereElementIsNotElementType() \
    .ToElements()

num_rooms = len(rooms)

t = Transaction(doc, "Eliminar todas las habitaciones")
t.Start()

for r in rooms:
    doc.Delete(r.Id)

t.Commit()

TaskDialog.Show("Eliminar Rooms",
                u"Se eliminaron {} habitaciones del proyecto.".format(num_rooms))
