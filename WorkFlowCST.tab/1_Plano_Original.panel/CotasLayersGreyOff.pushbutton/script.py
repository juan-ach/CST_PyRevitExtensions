# -*- coding: utf-8 -*-

__title__ = "Turn Off/Grey Layers"
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
Date    = 10.12.2025
_____________________________________________________________________
Description:

Turns off measure and text layers from original plan, also turn into soft gray remaining layers

_____________________________________________________________________
How-to:

Just press the button and see how your CAD import changes. If not, look for the script to add new layer names.
_____________________________________________________________________
Last update:
_____________________________________________________________________
Author: Juan Manuel Achenbach Anguita & ChatGPT"""

import clr

# Revit API
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

# pyRevit helpers
from pyrevit import revit

# ---------------------------
# Popup con autocierre
# ---------------------------
import System
from System.Windows.Forms import Form, Label, Timer
import System.Drawing

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
        label.Font = System.Drawing.Font("Arial", 12)
        self.Controls.Add(label)

        timer = Timer()
        timer.Interval = duration_ms
        timer.Tick += self.close_popup
        timer.Start()

    def close_popup(self, sender, args):
        self.Close()


# -------------------------------------------------------------------
# CONFIGURACION
# -------------------------------------------------------------------
viewTemplateNames = [
"CST_FLO_ELE_Trazado",
"CST_FLO_HVAC",
"CST_FLO_SEG v1",
"CST_FLO_Servicios",
"CST_FLO_TUB_ESQ v2",
"CST_FLO_CO2 liq com v2",
"CST_FLO_Servicios BP",
"CST_FLO_SERV",
"CST_FLO_GLI",
"CST_FLO_CO2 liq com v3",
"CST_FLO_TUB LIQ COM2",
"CST_FLO_GLI_448",
"CST_FLO_TUB"
]

searchStrings = ["tex", "txt", "cota", "implan", "seccions", "cotes"]

# -------------------------------------------------------------------
# SCRIPT
# -------------------------------------------------------------------
doc = revit.doc

grayColor = Color(128, 128, 128)

# 1. View templates
allViews = list(FilteredElementCollector(doc).OfClass(View))
viewTemplates = [v for v in allViews if v.IsTemplate and v.Name in viewTemplateNames]

# 2. Enlaces DWG
allDWGLinks = list(FilteredElementCollector(doc).OfClass(ImportInstance))

# 3. OverrideGraphicSettings
overrideSettings = OverrideGraphicSettings()
overrideSettings = overrideSettings.SetProjectionLineColor(grayColor)
overrideSettings = overrideSettings.SetProjectionLineWeight(1)

# 4. Transaccion y logica principal
with revit.Transaction("Depurar DWG en View Templates"):
    for viewTemplate in viewTemplates:
        for dwgLink in allDWGLinks:
            linkCategory = dwgLink.Category
            if linkCategory is None:
                continue

            subCats = list(linkCategory.SubCategories) if linkCategory.SubCategories else []

            for subCat in subCats:
                catNameLower = subCat.Name.lower()
                hideLayer = any(s.lower() in catNameLower for s in searchStrings)

                if hideLayer:
                    viewTemplate.SetCategoryHidden(subCat.Id, True)
                else:
                    viewTemplate.SetCategoryOverrides(subCat.Id, overrideSettings)

# ---------------------------
# Popup final autocerrable
# ---------------------------
popup = AutoClosePopup("Cotes esborrades i capes en gris", duration_ms=2000)
popup.ShowDialog()
