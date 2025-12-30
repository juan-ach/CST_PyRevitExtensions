# -*- coding: utf-8 -*-

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
