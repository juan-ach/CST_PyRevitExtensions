# -*- coding: utf-8 -*-
# pyRevit script: Migración base de "antes de mandar a presto_v2.dyn"
# Revit 2025

from __future__ import print_function, division

import clr
clr.AddReference("RevitAPI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from pyrevit import revit, DB
from Autodesk.Revit.DB.Plumbing import PipeInsulation  # por si lo necesitas más adelante

import System
from System.Windows.Forms import Form, Label, Timer
import System.Windows.Forms
import System.Drawing

doc = revit.doc

# -------------------------------------------------------------
# POPUP AUTOCIERRE
# -------------------------------------------------------------

class AutoClosePopup(Form):
    def __init__(self, message, title=u"Info", duration_ms=10000):
        self.Text = title
        self.Width = 460
        self.Height = 240
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
# UTILIDADES GENERALES
# -------------------------------------------------------------

PARAM_PARTIDAS = u"Partidas_PRESTO"
PARAM_CODIGO   = u"Codigo_Presto"

def set_param_safe(elem, param_name, value):
    if elem is None:
        return False
    p = elem.LookupParameter(param_name)
    if not p or p.IsReadOnly:
        return False
    try:
        p.Set(value)
        return True
    except:
        return False

def get_param_str(elem, param_name):
    if not elem:
        return u""
    p = elem.LookupParameter(param_name)
    if not p:
        return u""
    val = p.AsString()
    if not val:
        val = p.AsValueString()
    return val or u""


# -------------------------------------------------------------
# 1) ZÓCALOS (WALL SWEEP) + CÓDIGOS PRESTO
#    (parte que ya migramos de tu otro Dynamo)
# -------------------------------------------------------------

ZOCALO_TYPE_NAME   = u"CST_Zocalo"
SEARCH_ZOCALO_2    = u"zocalo 2 lados"
SEARCH_ZOCALO_1    = u"zocalo 1 lado"

WALLSWEEP_PARTIDA = u"09"
WALLSWEEP_CODIGO  = u"ZOC.PP500.300"

PARAM_TIPO          = u"Tipo"
PARAM_NOMBRE_TIPO   = u"Nombre de tipo"

def get_wall_type(elem):
    if not isinstance(elem, DB.Wall):
        return None
    try:
        return elem.WallType
    except:
        try:
            return doc.GetElement(elem.GetTypeId())
        except:
            return None

def wall_matches_keywords(wall, keywords):
    if not isinstance(keywords, (list, tuple)):
        keywords = [keywords]

    strings = []
    strings.append(get_param_str(wall, PARAM_TIPO))
    strings.append(get_param_str(wall, PARAM_NOMBRE_TIPO))

    wtype = get_wall_type(wall)
    if wtype:
        try:
            strings.append(wtype.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
        except:
            try:
                strings.append(wtype.Name)
            except:
                pass

    strings = [s.lower() for s in strings if s]

    for kw in keywords:
        kw_l = kw.lower()
        for s in strings:
            if kw_l in s:
                return True
    return False

def find_wall_sweep_type_by_name(name):
    coll = DB.FilteredElementCollector(doc).WhereElementIsElementType()
    for et in coll:
        try:
            if et.Name == name:
                return et
        except:
            continue
    return None

def create_wall_sweeps_for_walls(walls, sweep_type, both_sides=False):
    created = []

    wstype_enum = DB.WallSweepType.Sweep
    wsi = DB.WallSweepInfo(wstype_enum, False)  # horizontal

    # altura 0 (ya está en unidades internas)
    wsi.Distance = 0.0

    for w in walls:
        if not isinstance(w, DB.Wall):
            continue

        if both_sides:
            wsi.WallSide = DB.WallSide.Exterior
            ws_ext = DB.WallSweep.Create(w, sweep_type.Id, wsi)
            if ws_ext:
                created.append(ws_ext)

            wsi.WallSide = DB.WallSide.Interior
            ws_int = DB.WallSweep.Create(w, sweep_type.Id, wsi)
            if ws_int:
                created.append(ws_int)
        else:
            ws = DB.WallSweep.Create(w, sweep_type.Id, wsi)
            if ws:
                created.append(ws)

    return created

def set_presto_on_wallsweeps(elems):
    count_ok = 0
    for e in elems:
        ok1 = set_param_safe(e, PARAM_PARTIDAS, WALLSWEEP_PARTIDA)
        ok2 = set_param_safe(e, PARAM_CODIGO, WALLSWEEP_CODIGO)
        if ok1 or ok2:
            count_ok += 1
    return count_ok


# -------------------------------------------------------------
# 2) ASIGNACIÓN BÁSICA DE PARTIDAS POR CATEGORÍA
#    (parte inferida claramente del Dynamo)
# -------------------------------------------------------------

PARTIDAS_BY_CATEGORY = {
    DB.BuiltInCategory.OST_PipeCurves:      u"04",    # Tuberías
    DB.BuiltInCategory.OST_DuctCurves:      u"01.09", # Conductos
    DB.BuiltInCategory.OST_DuctFitting:     u"01.09", # Uniones de conducto
    DB.BuiltInCategory.OST_PipeInsulations: u"04",    # Aislamientos de tubería
    DB.BuiltInCategory.OST_Walls:           u"09",    # Muros
    DB.BuiltInCategory.OST_Floors:          u"09",    # Suelos
}

def assign_partidas_for_category(bic, partida_value):
    coll = (DB.FilteredElementCollector(doc)
            .OfCategory(bic)
            .WhereElementIsNotElementType())
    elems = list(coll)
    count_ok = 0
    for e in elems:
        if set_param_safe(e, PARAM_PARTIDAS, partida_value):
            count_ok += 1
    return len(elems), count_ok


# -------------------------------------------------------------
# 3) HUECOS PARA LÓGICA COMPLEJA DEL DYNAMO
#    (aquí es donde pegarías el contenido exacto de tus Python nodes)
# -------------------------------------------------------------

def etiquetar_conductos_codigo_presto():
    """
    Aquí va la lógica del Python node del Dynamo que genera las ETIQUETAS
    de conductos (OUT = etiquetas), que luego se usaban como Codigo_Presto
    o similar.

    No puedo ver el script completo desde aquí (se recorta en medio), así que:

    1. Abre Dynamo.
    2. Localiza el Python node que:
       - Usa FilteredElementCollector(OST_DuctCurves)
       - Recorre los ducts
       - Crea una lista 'etiquetas'
    3. Copia TODO su código.
    4. Pégalo aquí dentro adaptando:
       - En vez de IN/OUT, usa directamente FilteredElementCollector como ya hace el script.
       - En vez de OUT = etiquetas, recorre los conductos y haz:
            set_param_safe(duct, PARAM_CODIGO, etiqueta_correspondiente)
    5. Devuelve el número de conductos tratados, por ejemplo.
    """
    # EJEMPLO VACÍO:
    return 0, 0   # (ducts_totales, ducts_con_codigo)


def cambiar_tipos_de_tuberia_segun_reglas():
    """
    Aquí va la lógica de los Python nodes que cambian el tipo de tubería
    (los que usan BuiltInParameter.ELEM_TYPE_PARAM para asignar pipe_type.Id).

    Igual que arriba:
    - Copias el contenido de esos Python nodes.
    - En lugar de IN/OUT, usas colecciones de tuberías desde FilteredElementCollector.
    - Devuelves cuántas tuberías cambiaste.
    """
    # EJEMPLO VACÍO:
    return 0


# -------------------------------------------------------------
# 4) MAIN
# -------------------------------------------------------------

def main():
    t = DB.Transaction(doc, "Antes de mandar a PRESTO (migración base)")
    t.Start()

    try:
        resumen_lineas = []

        # 4.1 ZÓCALOS (barridos de muro + parámetros PRESTO)
        wstype = find_wall_sweep_type_by_name(ZOCALO_TYPE_NAME)
        walls_coll = (DB.FilteredElementCollector(doc)
                      .OfCategory(DB.BuiltInCategory.OST_Walls)
                      .WhereElementIsNotElementType())
        walls_all = list(walls_coll)

        walls_zoc_2 = [w for w in walls_all if wall_matches_keywords(w, SEARCH_ZOCALO_2)]
        walls_zoc_1 = [w for w in walls_all if wall_matches_keywords(w, SEARCH_ZOCALO_1)]

        sweeps_zoc_2 = []
        sweeps_zoc_1 = []
        sweeps_all   = []

        if wstype:
            sweeps_zoc_2 = create_wall_sweeps_for_walls(walls_zoc_2, wstype, both_sides=True)
            sweeps_zoc_1 = create_wall_sweeps_for_walls(walls_zoc_1, wstype, both_sides=False)
            sweeps_all   = sweeps_zoc_2 + sweeps_zoc_1

            ws_with_presto = set_presto_on_wallsweeps(sweeps_all)

            resumen_lineas.append(
                u"Zócalos: {} muros zócalo 2 lados, {} muros zócalo 1 lado.".format(
                    len(walls_zoc_2), len(walls_zoc_1)
                )
            )
            resumen_lineas.append(
                u"WallSweeps creados: {} (2 lados), {} (1 lado); {} con PRESTO.".format(
                    len(sweeps_zoc_2), len(sweeps_zoc_1), ws_with_presto
                )
            )
        else:
            resumen_lineas.append(
                u"[AVISO] No se encontró el tipo de zócalo '{}'; no se crearon WallSweeps.".format(
                    ZOCALO_TYPE_NAME
                )
            )

        # 4.2 PARTIDAS PRESTO POR CATEGORÍA
        total_cat_elems = 0
        total_cat_asig  = 0
        for bic, partida in PARTIDAS_BY_CATEGORY.items():
            n_cat, n_ok = assign_partidas_for_category(bic, partida)
            total_cat_elems += n_cat
            total_cat_asig  += n_ok
            resumen_lineas.append(
                u"{}: {} elementos, {} con Partidas_PRESTO='{}'".format(
                    bic.ToString(), n_cat, n_ok, partida
                )
            )

        # 4.3 LÓGICA COMPLEJA (A COMPLETAR CON LOS PYTHON NODES DE DYNAMO)
        ducts_tot, ducts_cod = etiquetar_conductos_codigo_presto()
        if ducts_tot or ducts_cod:
            resumen_lineas.append(
                u"Ductos: {} elementos, {} con Codigo_Presto asignado vía script etiquetas.".format(
                    ducts_tot, ducts_cod
                )
            )

        pipes_changed = cambiar_tipos_de_tuberia_segun_reglas()
        if pipes_changed:
            resumen_lineas.append(
                u"Tuberías: {} elementos cambiaron de tipo según reglas de Dynamo.".format(
                    pipes_changed
                )
            )

        t.Commit()

    except Exception as e:
        t.RollBack()
        msg_err = u"Error en 'Antes de mandar a PRESTO':\n{}".format(e)
        popup = AutoClosePopup(msg_err, title=u"Antes de mandar a PRESTO", duration_ms=5000)
        popup.ShowDialog()
        raise

    # ---------------------------------------------------------
    # POPUP RESUMEN (AUTO-CIERRE 3 s)
    # ---------------------------------------------------------
    msg = u"Antes de mandar a PRESTO\n\n" + u"\n".join(resumen_lineas)
    popup = AutoClosePopup(msg, title=u"Antes de mandar a PRESTO", duration_ms=3000)
    popup.ShowDialog()


if __name__ == "__main__":
    main()
