# -*- coding: utf-8 -*-

__title__ = "Delete Circuits" 
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
_____________________________________________________________________
Description:

Delete Circuits between electrical panels and electrical connectors

_____________________________________________________________________
How-to:

Place the electrical panels, equipment and cable tray and then run the script

IMPORTANT: IF THE SCRIPT DON'T RECOGNIZE THE PANEL CHECK PANEL'S TYPE IN PANEL_NAME_FILTER VALUE!
_____________________________________________________________________

Author: Juan Manuel Achenbach Anguita & ChatGPT"""

# pyRevit – Borrar todos los circuitos eléctricos de un proyecto (SIN CONFIRMACIÓN)

from Autodesk.Revit.DB import *

# Documento activo (pyRevit)
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# Obtener todos los circuitos eléctricos
circuits = (FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_ElectricalCircuit)
            .ToElementIds())

count = len(circuits)

t = Transaction(doc, "Borrar todos los circuitos eléctricos")
t.Start()

# Borrado masivo
if count > 0:
    doc.Delete(circuits)

t.Commit()

# Si quieres eliminar también este mensaje final, puedes borrar esta línea.
print("Circuitos eliminados: {}".format(count))
