# -*- coding: utf-8 -*-

__title__ = "Diameter Checker" 
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
_____________________________________________________________________
Description:

Check that mechanical equipment diameter matchs with pipe diameter defined by enginnering team.

_____________________________________________________________________
How-to:

After pipe diameter correction, run the script to set mechanical equipment diameters values equal to pipe diameters

_____________________________________________________________________

Author: Juan Manuel Achenbach Anguita & ChatGPT"""

from pyrevit import revit, script
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Mechanical import MechanicalEquipment
from Autodesk.Revit.DB import FamilyInstance, BuiltInCategory, Transaction

import System
from System.Windows.Forms import Form, Label, Timer
import System.Drawing

doc = revit.doc


# ---------------------------
# Popup con autocierre
# ---------------------------
class AutoClosePopup(Form):
    def __init__(self, message, duration_ms=3000):
        self.Text = "Info"
        self.Width = 360
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


# ----------------------------------------
# Nombres de los parametros de los conectores
# ----------------------------------------
PARAM_A = u"Aspiración"
PARAM_L = u"Líquido"
# ----------------------------------------

def obtener_diametro_conexion_real(conector):
    refs = conector.AllRefs
    for r in refs:
        owner = r.Owner

        # Caso directo
        if isinstance(owner, Pipe):
            tipo_sistema = owner.MEPSystem.Name if owner.MEPSystem else ""
            return owner.Diameter, tipo_sistema

        # Caso con fitting intermedio
        if isinstance(owner, FamilyInstance):
            if owner.MEPModel and hasattr(owner.MEPModel, "ConnectorManager") and owner.MEPModel.ConnectorManager:
                for cf in owner.MEPModel.ConnectorManager.Connectors:
                    for rr in cf.AllRefs:
                        other_owner = rr.Owner
                        if isinstance(other_owner, Pipe):
                            tipo_sistema = other_owner.MEPSystem.Name if other_owner.MEPSystem else ""
                            return other_owner.Diameter, tipo_sistema
    return None, None


def actualizar_equipo(eq):
    if not eq.MEPModel:
        return False
    if not hasattr(eq.MEPModel, "ConnectorManager") or not eq.MEPModel.ConnectorManager:
        return False

    modifico = False

    for c in eq.MEPModel.ConnectorManager.Connectors:
        diam, tipo_sistema = obtener_diametro_conexion_real(c)
        if diam and tipo_sistema:
            if "A" in tipo_sistema:
                p = eq.LookupParameter(PARAM_A)
                if p and p.StorageType == StorageType.Double:
                    p.Set(diam)
                    modifico = True
            elif "L" in tipo_sistema:
                p = eq.LookupParameter(PARAM_L)
                if p and p.StorageType == StorageType.Double:
                    p.Set(diam)
                    modifico = True

    return modifico


# ----------------------------------------
# PROCESO PRINCIPAL
# ----------------------------------------
equipos = FilteredElementCollector(doc)\
    .OfClass(FamilyInstance)\
    .OfCategory(BuiltInCategory.OST_MechanicalEquipment)\
    .WhereElementIsNotElementType()\
    .ToElements()

equipos_validos = [
    eq for eq in equipos
    if eq.MEPModel and hasattr(eq.MEPModel, "ConnectorManager") and eq.MEPModel.ConnectorManager
]

contador = 0

t = Transaction(doc, "Actualizar diametros Aspiracion y Liquido")
t.Start()

for eq in equipos_validos:
    if actualizar_equipo(eq):
        contador += 1

t.Commit()

# Mostrar popup con autocierre
popup = AutoClosePopup(
    u"Es van revisar {} equips mecànics.".format(contador),
    duration_ms=2000
)
popup.ShowDialog()

