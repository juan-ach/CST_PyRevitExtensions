# -*- coding: utf-8 -*-

from __future__ import print_function, division

__title__ = "448A" 
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
_____________________________________________________________________
Description:

Add insulation following 448A insulation criteria

_____________________________________________________________________
How-to:

After pipe diameter correction, run the script to add insulation

_____________________________________________________________________

Author: Juan Manuel Achenbach Anguita & ChatGPT"""

import clr
clr.AddReference("RevitAPI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from pyrevit import revit, DB
from Autodesk.Revit.DB.Plumbing import PipeInsulation

import System
from System.Windows.Forms import Form, Label, Timer
import System.Windows.Forms
import System.Drawing

doc = revit.doc

# -------------------------------------------------------------
# CLASE POPUP AUTOCIERRE
# -------------------------------------------------------------

class AutoClosePopup(Form):
    def __init__(self, message, duration_ms=3000):
        self.Text = "Info"
        self.Width = 420
        self.Height = 180
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen

        label = Label()
        label.Text = message
        label.Dock = System.Windows.Forms.DockStyle.Fill
        label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        label.Font = System.Drawing.Font("Arial", 10)
        self.Controls.Add(label)

        timer = Timer()
        timer.Interval = duration_ms
        timer.Tick += self.close_popup
        timer.Start()

    def close_popup(self, sender, args):
        self.Close()

# -------------------------------------------------------------
# CONFIGURACIÓN
# -------------------------------------------------------------

SYSTEMS_POSITIVOS = ["A1+", "A2+", "A3+", "L1+_ASPIRACIÓN", "L2+_ASPIRACIÓN"]
SYSTEMS_NEGATIVOS = ["A1-", "A2-", "A3-", "L1-_ASPIRACIÓN", "L2-_ASPIRACIÓN"]

SIZE_INCH_TO_MM = {
    "2 5/8": 66.675,
    "2 1/8": 53.975,
    "1 5/8": 41.275,
    "1 3/8": 34.925,
    "1 1/8": 28.575,
    "7/8":   22.225,
    "3/4":   19.050,
    "5/8":   15.875,
    "1/2":   12.700,
    "3/8":   9.525,
    "1/4":   6.350,
}

INSULATION_32_BY_SIZE = {
    "2 1/8": u"AISLAMIENTO INSTAL. TUBERÍA COBRE 2 1/8 - 32mm",
    "2 5/8": u"AISLAMIENTO INSTAL. TUBERÍA COBRE 2 5/8 - 32mm",
    "1 5/8": u"AISLAMIENTO INSTAL. TUBERÍA COBRE 1 5/8 - 32mm",
    "1 3/8": u"AISLAMIENTO INSTAL. TUBERÍA COBRE 1 3/8 - 32mm",
    "1 1/8": u"AISLAMIENTO INSTAL. TUBERÍA COBRE 1 1/8 - 32mm",
    "7/8":   u"AISLAMIENTO INSTAL. TUBERÍA COBRE 7/8 - 32mm",
    "3/4":   u"AISLAMIENTO INSTAL. TUBERÍA COBRE 3/4 - 32mm",
    "5/8":   u"AISLAMIENTO INSTAL. TUBERÍA COBRE 5/8 - 32mm",
    "1/2":   u"AISLAMIENTO INSTAL. TUBERÍA COBRE 1/2 - 32mm",
    "3/8":   u"AISLAMIENTO INSTAL. TUBERÍA COBRE 3/8 - 32mm",
    "1/4":   u"AISLAMIENTO INSTAL. TUBERÍA COBRE 1/4 - 32mm",
}

INSULATION_19_BY_SIZE = {
    "3/8":   u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 3/8 - 19mm",
    "1/2":   u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 1/2 - 19mm",
    "5/8":   u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 5/8 - 19mm",
    "3/4":   u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 3/4 - 19mm",
    "1 1/8": u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 1 1/8 - 19mm",
    "7/8":   u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 7/8 - 19mm",
    "1 3/8": u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 1 3/8 - 19mm",
    "1 5/8": u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 1 5/8 - 19mm",
    "2 5/8": u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 2 5/8 - 19mm",
    "2 1/8": u"_AISLAMIENTO INSTAL. TUBERÍA COBRE 2 1/8 - 19mm",
}

SPECIAL_POS_1_4_NAME_CONTAINS = u"AF-6-006 -   1/4 - 32.0mm"
SPECIAL_POS_1_4_THICKNESS_MM = 15.0

THICKNESS_19_MM = 19.0
THICKNESS_32_MM = 32.0

PARAM_DIAMETER = u"Diámetro"
PARAM_SYSTEM_TYPE = u"Tipo de sistema"
PARAM_TYPE_NAME = u"Nombre de tipo"

# -------------------------------------------------------------
# FUNCIONES AUXILIARES
# -------------------------------------------------------------

def get_param_str(elem, param_name):
    p = elem.LookupParameter(param_name)
    if not p:
        return u""
    val = p.AsString()
    if not val:
        val = p.AsValueString()
    return val or u""


def get_diameter_mm(elem):
    p = elem.LookupParameter(PARAM_DIAMETER)
    if not p:
        return None
    val = p.AsDouble()
    if val is None:
        return None
    try:
        mm = DB.UnitUtils.ConvertFromInternalUnits(val, DB.UnitTypeId.Millimeters)
    except AttributeError:
        mm = DB.UnitUtils.ConvertFromInternalUnits(val, DB.DisplayUnitType.DUT_MILLIMETERS)
    return mm


def find_closest_size_by_mm(d_mm, tol_mm=0.5):
    if d_mm is None:
        return None
    best_size = None
    best_diff = None
    for size, mm in SIZE_INCH_TO_MM.items():
        diff = abs(d_mm - mm)
        if best_diff is None or diff < best_diff:
            best_diff = diff
            best_size = size
    if best_diff is not None and best_diff <= tol_mm:
        return best_size
    return None


def build_insulation_type_map():
    """
    nombre de tipo -> ElementType de categoría OST_PipeInsulations
    """
    coll = (DB.FilteredElementCollector(doc)
            .OfCategory(DB.BuiltInCategory.OST_PipeInsulations)
            .WhereElementIsElementType())
    result = {}
    for t in coll:
        name_param = t.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if not name_param:
            continue
        name = name_param.AsString()
        if name:
            result[name] = t
    return result


def find_insulation_type_by_contains(all_types, substring):
    substring_lower = substring.lower()
    for name, t in all_types.items():
        if substring_lower in name.lower():
            return t
    return None


def convert_mm_to_internal(mm):
    try:
        return DB.UnitUtils.ConvertToInternalUnits(mm, DB.UnitTypeId.Millimeters)
    except AttributeError:
        return DB.UnitUtils.ConvertToInternalUnits(mm, DB.DisplayUnitType.DUT_MILLIMETERS)


def get_pipes_already_insulated():
    coll_ins = (DB.FilteredElementCollector(doc)
                .OfCategory(DB.BuiltInCategory.OST_PipeInsulations)
                .WhereElementIsNotElementType())
    host_ids = set()
    for ins in coll_ins:
        host_id = getattr(ins, "HostElementId", None)
        if host_id and host_id.IntegerValue not in host_ids:
            host_ids.add(host_id.IntegerValue)
    return host_ids


# -------------------------------------------------------------
# LÓGICA PRINCIPAL
# -------------------------------------------------------------

def main():
    # Mapa de tipos
    all_types = build_insulation_type_map()
    if not all_types:
        popup = AutoClosePopup(
            u"No se han encontrado tipos de aislamiento de tubería en el modelo.",
            duration_ms=3000
        )
        popup.ShowDialog()
        return

    already_insulated = get_pipes_already_insulated()

    special_1_4_type = find_insulation_type_by_contains(all_types, SPECIAL_POS_1_4_NAME_CONTAINS)

    pipe_coll = (DB.FilteredElementCollector(doc)
                 .OfCategory(DB.BuiltInCategory.OST_PipeCurves)
                 .WhereElementIsNotElementType())
    pipes_to_process = list(pipe_coll)

    count_created = 0
    count_skipped_have_ins = 0
    count_skipped_no_type = 0
    count_skipped_no_match = 0

    t = DB.Transaction(doc, "Aislamiento tuberías transcrítico (pyRevit)")
    t.Start()

    try:
        for pipe in pipes_to_process:
            # ya tiene aislamiento
            if pipe.Id.IntegerValue in already_insulated:
                count_skipped_have_ins += 1
                continue

            system_type = get_param_str(pipe, PARAM_SYSTEM_TYPE)

            is_pos = system_type in SYSTEMS_POSITIVOS
            is_neg = system_type in SYSTEMS_NEGATIVOS

            if not (is_pos or is_neg):
                count_skipped_no_match += 1
                continue

            diam_text = get_param_str(pipe, PARAM_DIAMETER)
            d_mm = get_diameter_mm(pipe)
            size_inch = find_closest_size_by_mm(d_mm)

            insulation_type = None
            thickness_mm = None

            # Grupo POSITIVO
            if is_pos:
                # caso especial 1/4
                if diam_text.strip() == "1/4" and special_1_4_type:
                    insulation_type = special_1_4_type
                    thickness_mm = SPECIAL_POS_1_4_THICKNESS_MM
                else:
                    if not size_inch or size_inch not in INSULATION_19_BY_SIZE:
                        count_skipped_no_match += 1
                        continue
                    ins_name = INSULATION_19_BY_SIZE[size_inch]
                    insulation_type = all_types.get(ins_name, None)
                    thickness_mm = THICKNESS_19_MM

            # Grupo NEGATIVO
            elif is_neg:
                if not size_inch or size_inch not in INSULATION_32_BY_SIZE:
                    count_skipped_no_match += 1
                    continue
                ins_name = INSULATION_32_BY_SIZE[size_inch]
                insulation_type = all_types.get(ins_name, None)
                thickness_mm = THICKNESS_32_MM

            if insulation_type is None:
                count_skipped_no_type += 1
                continue

            thickness_internal = convert_mm_to_internal(thickness_mm)

            try:
                # IMPORTANTE: clase en Autodesk.Revit.DB.Plumbing
                PipeInsulation.Create(doc, pipe.Id, insulation_type.Id, thickness_internal)
                count_created += 1
            except Exception as e:
                print("Error creando aislamiento en tubería {}: {}".format(pipe.Id.IntegerValue, e))
                count_skipped_no_type += 1
                continue

        t.Commit()

    except Exception:
        t.RollBack()
        raise

    # ---------------------------------------------------------
    # POPUP CON RESUMEN (AUTO-CIERRE 3 s)
    # ---------------------------------------------------------
    mensaje = u"Aislamiento tuberías transcrítico\n\n" \
              u"Tuberías procesadas:   {}\n" \
              u"Aislamientos creados:  {}\n" \
              u"Saltadas (ya tenían aislamiento): {}\n" \
              u"Saltadas (sin tipo de aislamiento): {}\n" \
              u"Saltadas (sin coincidencia tamaño/sistema): {}".format(
                  len(pipes_to_process),
                  count_created,
                  count_skipped_have_ins,
                  count_skipped_no_type,
                  count_skipped_no_match
              )

    popup = AutoClosePopup(mensaje, duration_ms=3000)
    popup.ShowDialog()


if __name__ == "__main__":
    main()
