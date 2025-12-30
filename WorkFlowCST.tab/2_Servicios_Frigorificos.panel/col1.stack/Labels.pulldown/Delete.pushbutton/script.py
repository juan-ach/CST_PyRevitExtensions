# -*- coding: utf-8 -*-

__title__ = "Delete SERV Labels" 
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
Date    = 10.12.2025
_____________________________________________________________________
Description:

Delete Tags for all Services
_____________________________________________________________________
How-to:

Just run the script to erase all Mechanical Equipement tags
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
# Categorías que representan ETIQUETAS en Revit
# (puedes agregar más si las necesitas)
# ==================================================
categorias_tags = [
    BuiltInCategory.OST_Tags,                     # Categoría general de tags
    BuiltInCategory.OST_MechanicalEquipmentTags,
    BuiltInCategory.OST_PipeTags,
    BuiltInCategory.OST_DuctTags,
    BuiltInCategory.OST_ElectricalEquipmentTags,
    BuiltInCategory.OST_ElectricalCircuitTags,
    BuiltInCategory.OST_AreaTags,
]

# Convertir a un set para evitar duplicados
categorias_tags = set(categorias_tags)

# ==================================================
# Recolectar todas las etiquetas de la vista activa
# ==================================================
coleccion = FilteredElementCollector(doc, view.Id) \
    .WhereElementIsNotElementType()

tags_para_borrar = []

for el in coleccion:
    try:
        if el.Category and el.Category.Id.IntegerValue in [int(cat) for cat in categorias_tags]:
            tags_para_borrar.append(el)
    except:
        pass

# ==================================================
# Borrado
# ==================================================
t = Transaction(doc, "Borrar etiquetas de vista")
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
    "Se borraron {} etiquetas en la vista activa.".format(contador)
)
