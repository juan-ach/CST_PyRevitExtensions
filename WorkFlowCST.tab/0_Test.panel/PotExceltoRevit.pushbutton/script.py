# -*- coding: utf-8 -*-
"""
Script: autoubipot - Asignación automática de Pot.Frigorifica a Evaporadores
Flujo:
  1. Recoge todos los equipos mecánicos cuya familia contenga "Evaporadores"
     excluyendo los que tengan "Hielo" en su nombre de tipo.
  2. Lee el parámetro de ejemplar "ubicación" de cada equipo.
  3. Abre un diálogo para que el usuario seleccione el fichero Excel.
  4. Busca coincidencias entre los valores de "ubicación" y la columna C
     de la pestaña "3. ELEM. PRINCIPALES I.F." del Excel.
  5. Obtiene el valor calculado de la columna G en la fila coincidente.
  6. Escribe ese valor en el parámetro "Pot.Frigorifica" del equipo.

Nota: la lectura del Excel se realiza con openpyxl (data_only=True), que lee
los valores cacheados por Excel (último resultado calculado guardado en el
fichero). El fichero debe estar guardado con los cálculos actualizados.
"""

from pyrevit import revit, DB, forms, script
import clr
import os
import sys

logger = script.get_logger()
output = script.get_output()

# ===========================================================================
# PASO 1: Recopilar equipos mecánicos que contengan "Evaporadores" en la
#         familia, excluyendo los que tengan "Hielo" en el nombre de tipo.
# ===========================================================================

doc = revit.doc

collector = (
    DB.FilteredElementCollector(doc)
    .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)
    .WhereElementIsNotElementType()
)

evaporadores = []
for equipo in collector:
    familia = equipo.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
    if familia is None:
        continue
    nombre_familia = familia.AsValueString() or ""
    nombre_tipo = equipo.Name or ""
    if "Evaporadores" in nombre_familia and "Hielo" not in nombre_tipo:
        evaporadores.append(equipo)

if not evaporadores:
    forms.alert(
        "No se encontraron equipos mecánicos con 'Evaporadores' en el nombre de familia.",
        exitscript=True
    )

output.print_md("### Equipos encontrados: {}".format(len(evaporadores)))

# ===========================================================================
# PASO 2: Leer el parámetro de ejemplar "ubicación" de cada equipo.
# ===========================================================================

def get_param_value(element, param_name):
    """Devuelve el valor de un parámetro de ejemplar como cadena, o None."""
    param = element.LookupParameter(param_name)
    if param is None:
        return None
    storage = param.StorageType
    if storage == DB.StorageType.String:
        return param.AsString()
    elif storage == DB.StorageType.Double:
        return str(param.AsDouble())
    elif storage == DB.StorageType.Integer:
        return str(param.AsInteger())
    elif storage == DB.StorageType.ElementId:
        return str(param.AsElementId().IntegerValue)
    return None

# Diccionario {elemento: valor_ubicacion}
equipos_ubicacion = {}
for equipo in evaporadores:
    ub = get_param_value(equipo, "ubicación")
    if ub is None:
        ub = get_param_value(equipo, "Ubicación")  # prueba con mayúscula
    if ub is None:
        ub = get_param_value(equipo, "UBICACION")
    if ub:
        ub = ub.strip()
    equipos_ubicacion[equipo] = ub

# Mostrar resumen rápido
output.print_md("**Valores de 'ubicación' recogidos:**")
for eq, ub in equipos_ubicacion.items():
    nombre_tipo = eq.Name or "(sin tipo)"
    output.print_md("- {} → `{}`".format(nombre_tipo, ub))

# ===========================================================================
# PASO 3: Selección del fichero Excel por parte del usuario.
# ===========================================================================

excel_path = forms.pick_file(
    file_ext="xlsm",
    title="Selecciona el fichero Excel del proyecto"
)

if not excel_path:
    forms.alert("No se seleccionó ningún fichero. Se cancela la operación.", exitscript=True)

output.print_md("**Fichero seleccionado:** `{}`".format(excel_path))

# ===========================================================================
# PASO 4 y 5: Leer el Excel mediante un script auxiliar externo que se
#             ejecuta con el Python del sistema (CPython), evitando las
#             limitaciones de IronPython (entorno de pyRevit).
#
#  - El helper está en la carpeta de Helpers compartidos del proyecto.
#  - Se llama con "py.exe" (Python Launcher para Windows).
#  - El helper devuelve un JSON {clave_colC: valor_colG} por stdout.
#  - IMPORTANTE: el fichero debe estar guardado en Excel con los cálculos
#                actualizados (openpyxl lee valores cacheados).
# ===========================================================================

import subprocess
import json
import os

HOJA_NOMBRE = "3. ELEM. PRINCIPALES I.F."

# Ruta fija al helper (fuera de la carpeta del botón)
_helper = r"C:\Users\USUARIO-1\Desktop\JUAN\PyRevit Extensions\WorkFlowCST.extension\Helpers\Lectura de excel para seteo de potencia de evaporadores\read_excel_helper.py"

if not os.path.isfile(_helper):
    forms.alert(
        "No se encontró el fichero auxiliar:\n{}\n\n"
        "Asegúrate de que 'read_excel_helper.py' está en la misma carpeta "
        "que 'script.py'.".format(_helper),
        exitscript=True
    )

# Llamar al helper con el Python del sistema usando el Launcher "py.exe"
try:
    proc = subprocess.Popen(
        ["py", _helper, excel_path, HOJA_NOMBRE],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        creationflags=0x08000000  # CREATE_NO_WINDOW
    )
    stdout, stderr = proc.communicate()
except OSError:
    # "py" no disponible: intentar con "python"
    try:
        proc = subprocess.Popen(
            ["python", _helper, excel_path, HOJA_NOMBRE],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=0x08000000
        )
        stdout, stderr = proc.communicate()
    except OSError as e:
        forms.alert(
            "No se encontró ningún Python del sistema (ni 'py' ni 'python').\n"
            "Instala Python 3 desde python.org y asegúrate de que está en el PATH.\n\n"
            "Error: {}".format(str(e)),
            exitscript=True
        )

# Decodificar la salida del helper
try:
    raw = stdout.decode("utf-8").strip()
    if not raw:
        err_detail = stderr.decode("utf-8", errors="replace")
        forms.alert(
            "El script auxiliar no devolvió ningún resultado.\n\n"
            "Stderr:\n{}".format(err_detail),
            exitscript=True
        )
    helper_result = json.loads(raw)
except Exception as e:
    forms.alert(
        "Error al interpretar la respuesta del script auxiliar:\n{}\n\n"
        "Salida bruta:\n{}".format(str(e), stdout[:500]),
        exitscript=True
    )

# Comprobar si el helper reportó un error
if "error" in helper_result:
    forms.alert(
        "Error en la lectura del Excel:\n{}".format(helper_result["error"]),
        exitscript=True
    )

excel_dict = helper_result.get("data", {})
filas_con_none_g = helper_result.get("none_g", [])
max_row = helper_result.get("max_row", 200)

output.print_md("**Se analizó el documento hasta la fila:** {}".format(max_row))

# Construir el diccionario final: {ubicacion_coincidente: pot_frigorifica}
resultado = {}
sin_coincidencia = []

for equipo, ubicacion in equipos_ubicacion.items():
    if not ubicacion:
        output.print_md(
            "- ⚠️ Equipo `{}` no tiene valor en el parámetro 'ubicación'. Se omite.".format(equipo.Name)
        )
        continue

    if ubicacion in excel_dict:
        pot = excel_dict[ubicacion]
        resultado[equipo] = (ubicacion, pot)
    else:
        sin_coincidencia.append((equipo, ubicacion))

# ===========================================================================
# Comparativa: Equipos en el modelo vs Excel
# ===========================================================================
output.print_md("### 📊 Comparativa de Equipos (Modelo vs Excel)")
if sin_coincidencia:
    output.print_md("⚠️ **Los siguientes equipos están en el modelo de Revit pero NO se encontraron en el Excel:**")
    for eq, ub in sin_coincidencia:
        output.print_md("- `{}` — ubicación: `{}`".format(eq.Name, ub))
else:
    output.print_md("✅ **Todos los equipos ('Evaporadores') presentes en el modelo fueron correctamente hallados en el Excel.**")

# Mostrar diccionario de coincidencias
if resultado:
    output.print_md("### ✅ Coincidencias encontradas:")
    for eq, (ub, pot) in resultado.items():
        output.print_md("- `{}` → clave: `{}` | Pot.Frigorifica: `{}`".format(eq.Name, ub, pot))

# ===========================================================================
# PASO 6: Escribir el valor de la columna G en el parámetro "Pot.Frigorifica"
#         del equipo mecánico correspondiente.
# ===========================================================================

if not resultado:
    forms.alert(
        "No se encontraron coincidencias entre los valores de 'ubicación' "
        "y la columna C del Excel. No se realizaron cambios.",
        exitscript=True
    )

errores = []
escritos = 0

with revit.Transaction("Asignar Pot.Frigorifica a Evaporadores"):
    for equipo, (ubicacion, pot_valor) in resultado.items():
        if pot_valor is None:
            output.print_md(
                "- ⚠️ Valor nulo en columna G para ubicación `{}`. Se omite.".format(ubicacion)
            )
            continue

        param = equipo.LookupParameter("Pot.Frigorifica")
        if param is None:
            errores.append(
                "Equipo `{}` no tiene el parámetro 'Pot.Frigorifica'.".format(equipo.Name)
            )
            continue

        if param.IsReadOnly:
            errores.append(
                "El parámetro 'Pot.Frigorifica' del equipo `{}` es de solo lectura.".format(equipo.Name)
            )
            continue

        try:
            storage = param.StorageType
            if storage == DB.StorageType.String:
                param.Set(str(pot_valor))
            elif storage == DB.StorageType.Double:
                # Revit almacena los valores de potencia (Carga de refrigeración) en unidades
                # internas (sistema Imperial). Si pasamos el número directamente sin convertir, 
                # Revit lo toma como unidad interna y en la interfaz (W) se ve dividido (ej: 242 W).
                # Para replicar el comportamiento de "escribir a mano en UI", convertimos explícitamente
                # el valor de Vatios (W) a la unidad interna antes de inyectarlo en la base de datos.
                raw = float(pot_valor)
                
                # Fallback matemático exacto (1 W = ~10.7639 Unidades Internas de Revit)
                internal_val = raw * ((1.0 / 0.3048) ** 2)
                
                try:
                    if hasattr(DB, 'UnitTypeId') and hasattr(DB.UnitTypeId, 'Watts'):
                        internal_val = DB.UnitUtils.ConvertToInternalUnits(raw, DB.UnitTypeId.Watts)
                    elif hasattr(DB, 'DisplayUnitType') and hasattr(DB.DisplayUnitType, 'DUT_WATTS'):
                        internal_val = DB.UnitUtils.ConvertToInternalUnits(raw, DB.DisplayUnitType.DUT_WATTS)
                except Exception:
                    pass
                
                param.Set(internal_val)
            elif storage == DB.StorageType.Integer:
                param.Set(int(pot_valor))
            else:
                errores.append(
                    "Tipo de dato no soportado para 'Pot.Frigorifica' en equipo `{}`.".format(equipo.Name)
                )
                continue
            escritos += 1
        except Exception as e:
            errores.append(
                "Error al escribir en equipo `{}`: {}".format(equipo.Name, str(e))
            )

# ===========================================================================
# Resumen final
# ===========================================================================

output.print_md("---")
output.print_md("## Resumen")
output.print_md("- Equipos actualizados correctamente: **{}**".format(escritos))
output.print_md("- Equipos sin coincidencia en Excel: **{}**".format(len(sin_coincidencia)))

if errores:
    output.print_md("### ❌ Errores:")
    for err in errores:
        output.print_md("- " + err)

if escritos > 0:
    forms.alert(
        "{} equipo(s) actualizados con 'Pot.Frigorifica' correctamente.".format(escritos)
    )
else:
    forms.alert("No se actualizó ningún equipo. Revisa el log para más detalles.")
