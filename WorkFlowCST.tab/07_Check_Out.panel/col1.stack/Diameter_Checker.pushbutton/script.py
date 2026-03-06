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
import unicodedata

doc = revit.doc
output = script.get_output()


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

def normalizar_nombre(txt):
    if not txt:
        return ""
    return " ".join(texto_celda(txt).lower().split())


def normalizar_texto_sistema(txt):
    if not txt:
        return ""
    try:
        t = unicode(txt)
    except:
        t = str(txt)
    t = t.lower().strip()
    try:
        t = unicodedata.normalize("NFKD", t)
        t = u"".join(ch for ch in t if not unicodedata.combining(ch))
    except:
        pass
    return t


def clasificar_sistema(tipo_sistema):
    t = normalizar_texto_sistema(tipo_sistema)
    if not t:
        return None

    # Regla especial AUTONOMO:
    # no usar "A"/"L" sueltos; usar ASPIRACION / LIQUIDO.
    if "autonomo" in t:
        if "aspiracion" in t:
            return "A"
        if "liquido" in t:
            return "L"
        return None

    # Regla general (legacy): contiene A o L.
    if "a" in t:
        return "A"
    if "l" in t:
        return "L"
    return None


def es_equipo_excluido(eq):
    objetivos_exactos = set([
        "central frigorifica"
    ])
    objetivos_contiene_base = [
        "gascooler",
        "desrecalentador"
    ]
    objetivos_contiene_tipo_ubic_extra = [
        "cond.",
        "central"
    ]

    def coincide_exclusion_base(valor):
        n = normalizar_nombre(valor)
        if not n:
            return False
        if n in objetivos_exactos:
            return True
        for token in objetivos_contiene_base:
            if token in n:
                return True
        return False

    def coincide_exclusion_tipo_ubic(valor):
        n = normalizar_nombre(valor)
        if not n:
            return False
        if coincide_exclusion_base(n):
            return True
        for token in objetivos_contiene_tipo_ubic_extra:
            if token in n:
                return True
        return False

    # Regla solicitada: excluir por NOMBRE DE TIPO
    # (incluye contiene gascooler/desrecalentador/cond./central).
    try:
        if eq.Symbol and coincide_exclusion_tipo_ubic(eq.Symbol.Name):
            return True
    except:
        pass

    # Regla solicitada: excluir por valor del parametro "ubicacion/ubicación"
    # (incluye contiene gascooler/desrecalentador/cond./central).
    try:
        ubic = obtener_parametro_texto(eq, [u"ubicación", u"Ubicación", u"ubicacion", u"Ubicacion"])
        if coincide_exclusion_tipo_ubic(ubic):
            return True
    except:
        pass

    # Mantener exclusión por otros textos de identificación del equipo.
    candidatos = []

    try:
        candidatos.append(eq.Name)
    except:
        pass

    try:
        if eq.Symbol:
            candidatos.append(eq.Symbol.Name)
            if eq.Symbol.Family:
                candidatos.append(eq.Symbol.Family.Name)
    except:
        pass

    try:
        candidatos.append(nombre_evaporador(eq))
    except:
        pass

    for c in candidatos:
        if coincide_exclusion_tipo_ubic(c):
            return True
    return False


def es_conector_tuberia(conector):
    try:
        return conector.Domain == Domain.DomainPiping
    except:
        try:
            return "piping" in str(conector.Domain).lower()
        except:
            return False


def clave_conector(conector):
    owner_id = -1
    ox = 0.0
    oy = 0.0
    oz = 0.0
    dom = 0
    ctype = 0

    try:
        owner_id = conector.Owner.Id.IntegerValue
    except:
        pass
    try:
        o = conector.Origin
        ox = round(o.X, 6)
        oy = round(o.Y, 6)
        oz = round(o.Z, 6)
    except:
        pass
    try:
        dom = int(conector.Domain)
    except:
        pass
    try:
        ctype = int(conector.ConnectorType)
    except:
        pass

    return (owner_id, ox, oy, oz, dom, ctype)


def nombre_owner(owner):
    try:
        return owner.Name
    except:
        try:
            return owner.GetType().Name
        except:
            return "Owner"


def descripcion_owner(owner):
    try:
        return "{}({})".format(texto_celda(nombre_owner(owner)), owner.Id.IntegerValue)
    except:
        return texto_celda(nombre_owner(owner))


def obtener_conectores_owner(owner):
    conectores = []
    try:
        if hasattr(owner, "ConnectorManager") and owner.ConnectorManager:
            conectores = list(owner.ConnectorManager.Connectors)
    except:
        conectores = []

    if not conectores:
        try:
            if owner.MEPModel and hasattr(owner.MEPModel, "ConnectorManager") and owner.MEPModel.ConnectorManager:
                conectores = list(owner.MEPModel.ConnectorManager.Connectors)
        except:
            conectores = []

    return [c for c in conectores if es_conector_tuberia(c)]


def obtener_conectores_tuberia_equipo(eq):
    if not eq or not eq.MEPModel:
        return []
    if not hasattr(eq.MEPModel, "ConnectorManager") or not eq.MEPModel.ConnectorManager:
        return []
    try:
        return [c for c in eq.MEPModel.ConnectorManager.Connectors if es_conector_tuberia(c)]
    except:
        return []


def conector_sin_conexion(conector):
    try:
        return not bool(conector.IsConnected)
    except:
        pass

    try:
        refs = list(conector.AllRefs)
    except:
        refs = []

    for r in refs:
        try:
            owner = r.Owner
            cowner = conector.Owner
            if owner and cowner and owner.Id.IntegerValue != cowner.Id.IntegerValue:
                return False
        except:
            continue

    return True


def buscar_pipe_desde_conector(conector_inicial):
    if not conector_inicial or not es_conector_tuberia(conector_inicial):
        return {"diam": None, "sistema": None, "sistema_param": None, "ruta": "Conector no es de tuberia"}

    cola = [(conector_inicial, [])]
    visitados = set()

    while cola:
        conector_actual, ruta_actual = cola.pop(0)
        key = clave_conector(conector_actual)
        if key in visitados:
            continue
        visitados.add(key)

        try:
            refs = list(conector_actual.AllRefs)
        except:
            refs = []

        for ref in refs:
            if not es_conector_tuberia(ref):
                continue

            owner = ref.Owner
            if not owner:
                continue

            ruta_owner = ruta_actual + [descripcion_owner(owner)]

            if isinstance(owner, Pipe):
                tipo_sistema = owner.MEPSystem.Name if owner.MEPSystem else ""
                tipo_sistema_param = ""
                try:
                    p_ts = owner.LookupParameter(u"Tipo de sistema")
                    if p_ts:
                        tipo_sistema_param = p_ts.AsString() or p_ts.AsValueString() or ""
                except:
                    tipo_sistema_param = ""
                return {
                    "diam": owner.Diameter,
                    "sistema": tipo_sistema,
                    "sistema_param": tipo_sistema_param,
                    "ruta": " -> ".join(ruta_owner)
                }

            for siguiente in obtener_conectores_owner(owner):
                skey = clave_conector(siguiente)
                if skey not in visitados:
                    cola.append((siguiente, ruta_owner))

    return {"diam": None, "sistema": None, "sistema_param": None, "ruta": "No se detecta Pipe conectado"}


def obtener_diametro_conexion_real(conector):
    info = buscar_pipe_desde_conector(conector)
    tipo_sistema = info.get("sistema_param") or info.get("sistema")
    return info["diam"], tipo_sistema


def obtener_parametro_double(eq, param_name):
    p = eq.LookupParameter(param_name)
    if not p or p.StorageType != StorageType.Double:
        return None
    return p.AsDouble()


def obtener_parametro_texto(eq, nombres_parametro):
    for nombre in nombres_parametro:
        p = eq.LookupParameter(nombre)
        if not p:
            continue
        try:
            val = p.AsString()
        except:
            val = None
        if not val:
            try:
                val = p.AsValueString()
            except:
                val = None
        if val:
            return val
    return "-"


def convertir_a_mm(valor_interno):
    if valor_interno is None:
        return None
    try:
        return UnitUtils.ConvertFromInternalUnits(valor_interno, UnitTypeId.Millimeters)
    except:
        return UnitUtils.ConvertFromInternalUnits(valor_interno, DisplayUnitType.DUT_MILLIMETERS)


def formatear_mm(valor_interno):
    mm = convertir_a_mm(valor_interno)
    if mm is None:
        return "-"
    return "{:.3f}".format(mm)


def obtener_archivo_modelo():
    try:
        path = doc.PathName
    except:
        path = ""
    if path and path.strip():
        return path
    try:
        return "Modelo sin guardar ({})".format(doc.Title)
    except:
        return "Modelo sin guardar"


def ha_cambiado(valor_antes, valor_despues, tolerancia=1e-09):
    if valor_antes is None and valor_despues is None:
        return False
    if valor_antes is None or valor_despues is None:
        return True
    return abs(valor_antes - valor_despues) > tolerancia


def texto_celda(valor):
    if valor is None:
        return "-"
    try:
        txt = u"{}".format(valor)
    except:
        txt = str(valor)
    return txt.replace("\r", " ").replace("\n", " ").strip()


def acotar_texto(texto, max_len):
    if texto is None:
        return "-"
    if len(texto) <= max_len:
        return texto
    if max_len <= 3:
        return texto[:max_len]
    return texto[:max_len - 3] + "..."


def nombre_evaporador(eq):
    try:
        if eq.Symbol and eq.Symbol.Family:
            return u"{} : {}".format(eq.Symbol.Family.Name, eq.Symbol.Name)
    except:
        pass
    return eq.Name if eq.Name else u"Element {}".format(eq.Id.IntegerValue)


def recopilar_diametros_por_equipo(equipos):
    datos = {}
    for eq in equipos:
        eq_id = eq.Id.IntegerValue
        datos[eq_id] = {
            "evaporador": nombre_evaporador(eq),
            "ubicacion": obtener_parametro_texto(eq, [u"ubicación", u"Ubicación", u"ubicacion", u"Ubicacion"]),
            "asp": obtener_parametro_double(eq, PARAM_A),
            "liq": obtener_parametro_double(eq, PARAM_L)
        }
    return datos


def analizar_conector_para_log(conector):
    return buscar_pipe_desde_conector(conector)


def recopilar_diagnostico_conectores_por_equipo(equipos):
    data = {}
    for eq in equipos:
        eq_id = eq.Id.IntegerValue
        filas = []
        for con in obtener_conectores_tuberia_equipo(eq):
            diag = analizar_conector_para_log(con)
            filas.append({
                "ruta": diag.get("ruta"),
                "sistema": diag.get("sistema"),
                "sistema_param": diag.get("sistema_param"),
                "diam": diag.get("diam")
            })
        data[eq_id] = filas
    return data


def imprimir_log_diametros(registros):
    output.print_md("## Log de cambios de diametro (solo equipos con al menos un 'No')")

    if not registros:
        output.print_md("No hay equipos para mostrar.")
        return

    filtrados = [r for r in registros if (not r["asp_cambio"]) or (not r["liq_cambio"])]

    def clave_orden(r):
        # Prioridad visual: No/No, No/Si, Si/No
        if (not r["asp_cambio"]) and (not r["liq_cambio"]):
            prioridad = 0
        elif (not r["asp_cambio"]) and r["liq_cambio"]:
            prioridad = 1
        else:
            prioridad = 2
        return (prioridad, texto_celda(r["evaporador"]).lower(), r["id"])

    ordenados = sorted(filtrados, key=clave_orden)

    if not ordenados:
        output.print_md("Todos los equipos cambiaron en aspiracion y liquido.")
        return

    col1 = u"Equipo"
    col2 = u"Ubicacion"
    col3 = u"Cambio en aspiracion"
    col4 = u"Cambio en liquido"
    rows = []
    max_equipo = 45
    max_ubicacion = 25

    for r in ordenados:
        rows.append([
            acotar_texto(texto_celda(r["evaporador"]), max_equipo),
            acotar_texto(texto_celda(r["ubicacion"]), max_ubicacion),
            u"Si" if r["asp_cambio"] else u"No",
            u"Si" if r["liq_cambio"] else u"No"
        ])

    # Tabla nativa de pyRevit (más estable y legible que bloque markdown con monospace).
    if hasattr(output, "print_table"):
        output.print_table(
            table_data=rows,
            columns=[col1, col2, col3, col4]
        )
    else:
        output.print_md("| {} | {} | {} | {} |".format(col1, col2, col3, col4))
        output.print_md("|---|---|---|---|")
        for row in rows:
            output.print_md("| {} | {} | {} | {} |".format(
                row[0].replace("|", "\\|"),
                row[1].replace("|", "\\|"),
                row[2],
                row[3]
            ))


def imprimir_diagnostico_conectores(equipos_por_id, registros, snapshot_antes, snapshot_despues, snapshot_conectores_antes):
    output.print_md("### Equipos donde al menos un conector no fue modificado")

    objetivo = [r for r in registros if (not r["asp_cambio"]) or (not r["liq_cambio"])]
    if not objetivo:
        output.print_md("No hay equipos a diagnosticar.")
        return

    objetivo_ordenado = sorted(objetivo, key=lambda r: (texto_celda(r["evaporador"]).lower(), r["id"]))
    rows = []
    rows_sin_pipe = []
    ids_sin_pipe = set()

    for reg in objetivo_ordenado:
        eq_id = reg["id"]
        eq = equipos_por_id.get(eq_id)
        nombre_eq = acotar_texto(texto_celda(reg.get("evaporador", "-")), 45)
        ubic = acotar_texto(texto_celda(reg.get("ubicacion", "-")), 25)

        if not eq or not eq.MEPModel or not hasattr(eq.MEPModel, "ConnectorManager") or not eq.MEPModel.ConnectorManager:
            continue

        conectores_actuales = obtener_conectores_tuberia_equipo(eq)
        if not conectores_actuales:
            continue

        # Regla: si al menos un conector de tuberia esta suelto,
        # el equipo sale de esta tabla y va a la tabla de "autonomo/no conectado".
        if any(conector_sin_conexion(c) for c in conectores_actuales):
            if eq_id not in ids_sin_pipe:
                rows_sin_pipe.append([nombre_eq, ubic])
                ids_sin_pipe.add(eq_id)
            continue

        conectores_antes = snapshot_conectores_antes.get(eq_id, [])
        if not conectores_antes:
            continue

        asp_cambio = ha_cambiado(snapshot_antes.get(eq_id, {}).get("asp"), snapshot_despues.get(eq_id, {}).get("asp"))
        liq_cambio = ha_cambiado(snapshot_antes.get(eq_id, {}).get("liq"), snapshot_despues.get(eq_id, {}).get("liq"))

        for diag_antes in conectores_antes:
            diam_original = "-"
            sistema = texto_celda(diag_antes.get("sistema_param")) if diag_antes.get("sistema_param") else "-"
            diam_corregido = "-"

            param_ok = "-"
            cambio_param = "-"
            obs = ""

            tipo_sistema_clasif = diag_antes.get("sistema_param") or diag_antes.get("sistema")
            if diag_antes.get("diam") is not None and tipo_sistema_clasif:
                clasificacion = clasificar_sistema(tipo_sistema_clasif)
                if clasificacion == "A":
                    p = eq.LookupParameter(PARAM_A)
                    param_ok = "Si" if (p and p.StorageType == StorageType.Double) else "No"
                    cambio_param = "Si" if asp_cambio else "No"
                    diam_original = formatear_mm(snapshot_antes.get(eq_id, {}).get("asp"))
                    diam_corregido = formatear_mm(snapshot_despues.get(eq_id, {}).get("asp"))
                elif clasificacion == "L":
                    p = eq.LookupParameter(PARAM_L)
                    param_ok = "Si" if (p and p.StorageType == StorageType.Double) else "No"
                    cambio_param = "Si" if liq_cambio else "No"
                    diam_original = formatear_mm(snapshot_antes.get(eq_id, {}).get("liq"))
                    diam_corregido = formatear_mm(snapshot_despues.get(eq_id, {}).get("liq"))
                else:
                    obs = "Sistema no clasificado por regla"
            else:
                obs = "No se detecta Pipe conectado"

            if not obs:
                if param_ok == "No":
                    obs = "Parametro destino no valido"
                elif cambio_param == "No":
                    obs = "Valor final no cambia"
                elif cambio_param == "Si":
                    obs = "Cambio detectado"
                else:
                    obs = "Sin clasificacion"

            rows.append([
                nombre_eq,
                ubic,
                sistema,
                diam_original,
                diam_corregido,
                obs
            ])

    if rows:
        if hasattr(output, "print_table"):
            output.print_table(
                table_data=rows,
                columns=[
                    "Equipo",
                    "Ubicacion",
                    "Sistema",
                    "Diámetro Original",
                    "Diámetro Corregido",
                    "Observacion"
                ]
            )
        else:
            output.print_md("| Equipo | Ubicacion | Sistema | Diámetro Original | Diámetro Corregido | Observacion |")
            output.print_md("|---|---|---|---|---|---|")
            for row in rows:
                output.print_md("| {} | {} | {} | {} | {} | {} |".format(
                    row[0].replace("|", "\\|"),
                    row[1].replace("|", "\\|"),
                    row[2].replace("|", "\\|"),
                    row[3],
                    row[4],
                    row[5].replace("|", "\\|")
                ))
    else:
        output.print_md("No hay filas de diagnostico por conector con pipe conectado.")

    output.print_md("### Chequear si es autónomo o no se encuentra conectado correctamente")
    rows_sin_pipe_ordenadas = sorted(rows_sin_pipe, key=lambda r: (texto_celda(r[0]).lower(), texto_celda(r[1]).lower()))
    if rows_sin_pipe_ordenadas:
        if hasattr(output, "print_table"):
            output.print_table(
                table_data=rows_sin_pipe_ordenadas,
                columns=["Equipo", "Ubicacion"]
            )
        else:
            output.print_md("| Equipo | Ubicacion |")
            output.print_md("|---|---|")
            for row in rows_sin_pipe_ordenadas:
                output.print_md("| {} | {} |".format(
                    row[0].replace("|", "\\|"),
                    row[1].replace("|", "\\|")
                ))
    else:
        output.print_md("No hay equipos en este grupo.")


def actualizar_equipo(eq):
    if not eq.MEPModel:
        return False
    if not hasattr(eq.MEPModel, "ConnectorManager") or not eq.MEPModel.ConnectorManager:
        return False

    modifico = False

    for c in obtener_conectores_tuberia_equipo(eq):
        diam, tipo_sistema = obtener_diametro_conexion_real(c)
        if diam is not None and tipo_sistema:
            clasificacion = clasificar_sistema(tipo_sistema)
            if clasificacion == "A":
                p = eq.LookupParameter(PARAM_A)
                if p and p.StorageType == StorageType.Double:
                    p.Set(diam)
                    modifico = True
            elif clasificacion == "L":
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
    if (not es_equipo_excluido(eq))
    and eq.MEPModel
    and hasattr(eq.MEPModel, "ConnectorManager")
    and eq.MEPModel.ConnectorManager
]
equipos_por_id = {eq.Id.IntegerValue: eq for eq in equipos_validos}

contador = 0
registros_log = []

# 1) Snapshot antes de modificar
snapshot_antes = recopilar_diametros_por_equipo(equipos_validos)
snapshot_conectores_antes = recopilar_diagnostico_conectores_por_equipo(equipos_validos)

t = Transaction(doc, "Actualizar diametros Aspiracion y Liquido")
t.Start()

for eq in equipos_validos:
    modificado = actualizar_equipo(eq)
    if modificado:
        contador += 1

t.Commit()

# 2) Snapshot despues de modificar
snapshot_despues = recopilar_diametros_por_equipo(equipos_validos)

# 3) Comparacion para el log
for eq in equipos_validos:
    eq_id = eq.Id.IntegerValue
    datos_antes = snapshot_antes.get(eq_id, {})
    datos_despues = snapshot_despues.get(eq_id, {})

    asp_cambio = ha_cambiado(datos_antes.get("asp"), datos_despues.get("asp"))
    liq_cambio = ha_cambiado(datos_antes.get("liq"), datos_despues.get("liq"))

    registros_log.append({
        "evaporador": datos_antes.get("evaporador") or datos_despues.get("evaporador") or nombre_evaporador(eq),
        "ubicacion": datos_antes.get("ubicacion") or datos_despues.get("ubicacion") or "-",
        "id": eq_id,
        "asp_cambio": asp_cambio,
        "liq_cambio": liq_cambio
    })

output.print_md("**Archivo Revit:** `{}`".format(obtener_archivo_modelo().replace("`", "'")))

imprimir_diagnostico_conectores(
    equipos_por_id,
    registros_log,
    snapshot_antes,
    snapshot_despues,
    snapshot_conectores_antes
)

# Mostrar popup con autocierre
popup = AutoClosePopup(
    u"Es van revisar {} equips mecànics.".format(contador),
    duration_ms=2000
)
popup.ShowDialog()
