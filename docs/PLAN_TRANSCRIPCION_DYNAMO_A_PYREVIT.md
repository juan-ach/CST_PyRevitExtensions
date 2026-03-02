# FAST VERSION V1 (90 min)
## Dynamo (`antes de mandar a presto_v2.txt`) -> pyRevit (`test_space.pushbutton/script.py`)

Fecha: 2026-03-02  
Objetivo de esta version: tener un diagnostico rapido, accionable y con prioridades claras.

## 1) Hechos confirmados del Dynamo
- Archivo analizado: `antes de mandar a presto_v2.txt` (JSON de Dynamo).
- Tamano del grafo: **422 nodos**.
- `PythonScriptNode`: **15** nodos (**7 logicas unicas**).
- `Element.SetParameterByName`: **22** nodos.
- Parametros escritos confirmados por esos 22 nodos:
  - `Partidas_PRESTO`
  - `Codigo_Presto`
  - `Comentarios`
  - `Vol.Camara` (aparece con codificacion de acento en el JSON)
  - `sup.bruta.panel`
  - `long.bruta.tub`
  - `Duct Fitting Area` (parametro por `StringInputNode`)

## 2) Logicas Python unicas encontradas en Dynamo
1. Creacion de wall sweep en un lado.
2. Creacion de wall sweep en ambos lados (Interior + Exterior).
3. Etiquetado de conductos por nombre de tipo:
   - `condensador` -> `COND.EXT.COND`
   - `turbina` -> `COND.EXT.TURB`
   - `gascooler` -> `COND.EXT.GASCOOLER`
4. Cambio de tipo de tuberia por coincidencia exacta de nombre de tipo (`ELEM_TYPE_PARAM`).
5. Cambio de tipo de tuberia por coincidencia parcial (contains).
6. Helper string contains (lista booleana).
7. Helper string equals (lista booleana).

## 3) Estado actual de `test_space.pushbutton/script.py`
Archivo evaluado:
- [script.py](c:/Users/USUARIO-1/Desktop/JUAN/PyRevit%20Extensions/WorkFlowCST.extension/WorkFlowCST.tab/0_Test.panel/test_space.pushbutton/script.py)

Cobertura actual:
- Si implementa:
  - utilidades de escritura/lectura de parametros.
  - creacion de zocalos (1 lado y 2 lados) y escritura PRESTO en wall sweeps.
  - asignacion base de `Partidas_PRESTO` por categorias.
  - etiquetado de `Codigo_Presto` en conductos por keywords.
- No implementa:
  - reglas reales de cambio de tipo de tuberia (funcion placeholder con `return 0`).
  - escrituras de `Comentarios`, `Vol.Camara`, `sup.bruta.panel`, `long.bruta.tub`, `Duct Fitting Area`.
  - pipeline de copia de parametros de tipo (`Codigo de montaje` -> `Codigo_Presto`).
  - bloques de limpieza (borrados), reglas de refrigerante y varios subflujos de tuberias/aislamientos del Dynamo.

## 4) Veredicto rapido de paridad
Estado de equivalencia respecto Dynamo:
- **NO 100% equivalente**.
- Estimacion de cobertura funcional actual: **30%-45%** (aprox, por bloques principales detectados).

Riesgo principal:
- El script actual parece "base de migracion", pero no replica los bloques de mayor impacto en medicion/codificacion de tuberias.

## 5) Matriz de gap (alta prioridad)
| Bloque | Dynamo | `script.py` | Gap |
|---|---|---|---|
| Zocalos | Si | Si | Ajustar solo detalles de paridad fina (distancia/reveal/vertical) |
| Partidas por categoria | Si | Parcial | Faltan categorias y ramas especiales |
| Codigo conductos por keywords | Si | Si | Validar igualdades exactas de cadenas |
| Cambio tipo tuberia (exact/contains) | Si | No | Critico |
| `long.bruta.tub` | Si | No | Critico |
| `sup.bruta.panel` | Si | No | Alto |
| `Vol.Camara` | Si | No | Alto |
| `Comentarios` (copias/formulas) | Si | No | Alto |
| `Duct Fitting Area` | Si | No | Alto |
| `Codigo de montaje` -> `Codigo_Presto` | Si | No | Alto |
| Limpieza (delete) | Si | No | Medio/Alto (segun uso) |

## 6) Plan tecnico FAST (ejecucion inmediata)
### Fase A (primero)
1. Implementar `cambiar_tipos_de_tuberia_segun_reglas()` con dos modos:
   - match exacto
   - match contains
2. Centralizar busqueda de `PipeType` por `SYMBOL_NAME_PARAM`.
3. Reportar en resumen:
   - tuberias evaluadas
   - cambiadas
   - no encontradas por regla

### Fase B
1. Implementar escrituras faltantes:
   - `long.bruta.tub`
   - `sup.bruta.panel`
   - `Vol.Camara`
   - `Comentarios`
   - `Duct Fitting Area`
2. Portar formulas Dynamo (misma unidad y conversion).
3. Agregar control de tipo numerico/string en `set_param_safe`.

### Fase C
1. Implementar rama `Codigo de montaje` (parametro de tipo) -> `Codigo_Presto`.
2. Completar categorias/ramas de `Partidas_PRESTO`.
3. Agregar modo seguro para limpieza:
   - `ENABLE_DELETE = False` por defecto.

## 7) Definicion de done para esta version fast
Se considera cerrada la migracion fast cuando:
1. Todas las funciones placeholder desaparecen.
2. Se escriben todos los parametros de la lista del punto 1.
3. El resumen final muestra metricas por bloque.
4. No hay excepciones no controladas en ejecucion completa.

## 8) Siguiente entrega recomendada
Proxima iteracion (V2) debe incluir:
1. Tabla bloque-a-bloque Dynamo -> funcion Python concreta.
2. Orden exacto de ejecucion para mantener dependencias.
3. Prueba de paridad en un modelo real con diff de parametros antes/despues.
