# -*- coding: utf-8 -*-

__title__ = "Create Evaporadores Labels"
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.1'
__doc__ = """Version: 1.1
_____________________________________________________________________
Description:

Create the same Mechanical Equipment tag for ALL mechanical equipment
in the active view (skips already tagged elements).
_____________________________________________________________________
Author: Juan Manuel Achenbach Anguita & ChatGPT"""

from pyrevit import revit
from Autodesk.Revit.DB import *
from System.Windows.Forms import Form, Label, Timer, FormStartPosition, DockStyle
import System.Drawing

doc = revit.doc
view = doc.ActiveView

# ==================================================
# Popup con autocierre
# ==================================================
class AutoClosePopup(Form):
    def __init__(self, message, duration_ms=2000):
        self.Text = "Info"
        self.Width = 420
        self.Height = 150
        self.StartPosition = FormStartPosition.CenterScreen

        label = Label()
        label.Text = message
        label.Dock = DockStyle.Fill
        label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        label.Font = System.Drawing.Font("Arial", 11)
        self.Controls.Add(label)

        timer = Timer()
        timer.Interval = duration_ms
        timer.Tick += self.close_popup
        timer.Start()

    def close_popup(self, sender, args):
        self.Close()

# ==================================================
# CONFIGURACIÓN
# ==================================================
POSIBLES_FAMILIAS = [
    "CST_rectangulo informativo_v14_catalan",
    "CST_rectangulo informativo_v11_castellano"
]
TIPOS_REQUERIDOS = ["1 modulo", "2 modulos", "3 modulos"]
NOMBRE_FAMILIA_TAG = None # Se definirá al encontrar la familia

# Offsets en metros -> convertidos a pies (Revit trabaja en pies)
# Ajusta a tu gusto:
OFFSET_X = 0.0 * 3.28084
OFFSET_Y = -0.5 * 3.28084
OFFSET_Z = 0.0

# ==================================================
# Función: calcular punto del tag (más robusto)
# ==================================================
def calcular_punto_tag(elem, dx=0, dy=0, dz=0):
    loc = elem.Location

    # Caso típico: LocationPoint
    if isinstance(loc, LocationPoint):
        return loc.Point + XYZ(dx, dy, dz)

    # Caso: LocationCurve (p.ej. algunos equipos "lineales")
    if isinstance(loc, LocationCurve) and loc.Curve:
        try:
            mid = loc.Curve.Evaluate(0.5, True)
            return mid + XYZ(dx, dy, dz)
        except:
            pass

    # Fallback: centro del bounding box (en la vista si existe)
    try:
        bb = elem.get_BoundingBox(view) or elem.get_BoundingBox(None)
        if bb:
            center = (bb.Min + bb.Max) * 0.5
            return center + XYZ(dx, dy, dz)
    except:
        pass

    return None

# ==================================================
# Helper: Check numeric value (float/int/string)
# ==================================================
def get_param_float(elem, param_name):
    p = elem.LookupParameter(param_name)
    if not p:
        return 0.0
    if p.StorageType == StorageType.Double:
        return p.AsDouble()
    if p.StorageType == StorageType.Integer:
        return float(p.AsInteger())
    if p.StorageType == StorageType.String:
        try:
            val_str = p.AsString()
            if val_str:
                return float(val_str)
        except:
            pass
    return 0.0

# ==================================================
# Buscar el tipo de etiqueta (IronPython-safe)
# ==================================================
tag_symbols = {} # Map: "tipo" -> FamilySymbolElement

# 1. Buscar familia disponible
for fs in FilteredElementCollector(doc) \
        .OfClass(FamilySymbol) \
        .OfCategory(BuiltInCategory.OST_MechanicalEquipmentTags):

    if fs.FamilyName in POSIBLES_FAMILIAS:
        NOMBRE_FAMILIA_TAG = fs.FamilyName
        break

if not NOMBRE_FAMILIA_TAG:
    raise Exception(
        u"No se encontró ninguna de las familias de etiquetas esperadas: {}.".format(
            ", ".join(POSIBLES_FAMILIAS)
        )
    )

# 2. Cargar todos los tipos requeridos de esa familia
for fs in FilteredElementCollector(doc) \
        .OfClass(FamilySymbol) \
        .OfCategory(BuiltInCategory.OST_MechanicalEquipmentTags):

    if fs.FamilyName == NOMBRE_FAMILIA_TAG:
        tn = fs.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        if tn in TIPOS_REQUERIDOS:
            tag_symbols[tn] = fs

# 3. Verificar que existen todos
faltantes = [t for t in TIPOS_REQUERIDOS if t not in tag_symbols]
if faltantes:
    raise Exception(
        u"La familia '{}' no tiene los tipos requeridos: {}.".format(
            NOMBRE_FAMILIA_TAG, ", ".join(faltantes)
        )
    )

# 4. Activar símbolos si es necesario
t = Transaction(doc, "Activar tipos de etiqueta")
t.Start()
for fs in tag_symbols.values():
    if not fs.IsActive:
        fs.Activate()
t.Commit()

# ==================================================
# Equipos mecánicos en la vista activa
# ==================================================
equipos = FilteredElementCollector(doc, view.Id) \
    .OfCategory(BuiltInCategory.OST_MechanicalEquipment) \
    .WhereElementIsNotElementType() \
    .ToElements()

# ==================================================
# Tags existentes en la vista (para no duplicar)
# ==================================================
tags_existentes = FilteredElementCollector(doc, view.Id) \
    .OfCategory(BuiltInCategory.OST_MechanicalEquipmentTags) \
    .WhereElementIsNotElementType() \
    .ToElements()

equipos_ya_etiquetados = set()

for tag in tags_existentes:
    try:
        for eid in tag.GetTaggedElementIds():
            if eid != ElementId.InvalidElementId:
                equipos_ya_etiquetados.add(eid)
    except:
        pass

# ==================================================
# CREACIÓN DE TAGS (UN SOLO TIPO PARA TODO)
# ==================================================
contador = 0

t = Transaction(doc, u"Etiquetar todos los equipos mecánicos (evaporadores)")
t.Start()

for eq in equipos:
    if eq.Id in equipos_ya_etiquetados:
        continue

    skip = False
    # Parametros a comprobar: si alguno es 0, no etiquetar
    for p_name in ["Aspiración", "Líquido", "W_Metro"]:
        param = eq.LookupParameter(p_name)
        if param:
            # Check por tipo de almacenamiento
            if param.StorageType == StorageType.Double:
                # 0.001 es una tolerancia razonable para float
                if abs(param.AsDouble()) < 0.001:
                    skip = True
                    break
            elif param.StorageType == StorageType.Integer:
                if param.AsInteger() == 0:
                    skip = True
                    break
            elif param.StorageType == StorageType.String:
                # Intentar convertir string a float ("0", "0.0")
                try:
                    val_str = param.AsString()
                    if val_str and abs(float(val_str)) < 0.001:
                        skip = True
                        break
                except:
                    pass
    
    if skip:
        continue

    # Determinar tipo de tag
    tipo_a_usar = "1 modulo" # Default

    # Lógica especial para familias "Mueble"
    elem_fam_name = eq.Symbol.FamilyName
    if "Mueble" in elem_fam_name:
        try:
            m2 = get_param_float(eq, "Modulo_2")
            m3 = get_param_float(eq, "Modulo_3")
            
            # Tolerancia float
            is_m2_nonzero = abs(m2) > 0.001
            is_m3_nonzero = abs(m3) > 0.001

            if is_m2_nonzero and not is_m3_nonzero:
                tipo_a_usar = "2 modulos"
            elif is_m2_nonzero and is_m3_nonzero:
                tipo_a_usar = "3 modulos"
            # Si ninguno es != 0, o solo m3 lo es (raro), se queda en "1 modulo"
        except:
            pass
            
    # Obtener el símbolo correspondiente
    simbolo_tag = tag_symbols.get(tipo_a_usar)
    if not simbolo_tag:
        continue # No debería pasar si validamos al inicio

    punto_tag = calcular_punto_tag(eq, OFFSET_X, OFFSET_Y, OFFSET_Z)
    if not punto_tag:
        continue

    try:
        tag = IndependentTag.Create(
            doc,
            view.Id,
            Reference(eq),
            False,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            punto_tag
        )

        tag.ChangeTypeId(simbolo_tag.Id)
        contador += 1

    except:
        pass

t.Commit()

# ==================================================
# Popup final
# ==================================================
popup = AutoClosePopup(
    u"Se etiquetaron {} equipos mecánicos con:\nFamilia: '{}'\n(Tipos dinámicos usados)".format(
        contador, NOMBRE_FAMILIA_TAG
    )
)
popup.ShowDialog()
