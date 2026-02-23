# CST PyRevit → C# Tooling (Nativo)

Este paquete contiene los comandos en C# para la biblioteca de automatizaciones del repo, listos para usarse como `IExternalCommand`.

## Qué incluye

- `NativeAutomationRunner.cs`:
  - Implementación nativa C# de la lógica operativa para:
    - Pin / Unpin
    - Borrado masivo de tags
    - Borrado de rooms, circuits e insulations
    - Limpieza DWG por templates (`Turn Off/Grey Layers`)
    - Creación de tags mecánicos
    - Room tags
    - Asignación de ubicación por puertas
    - Checker de diámetro
    - Creación de circuitos por relación cuadro/carga
    - Insulación por perfiles (`Transcrítico`, `Glicol`, `448A`, `134-448EVI`)
    - Asignación de `Partidas_PRESTO`
- `Commands/Cmd_*.cs`: 30 comandos (uno por cada botón/automatización pyRevit original).
- `Commands/CommandBase.cs`: base común de ejecución.

## Compilación

1. Compila contra .NET Framework 4.8.
2. Referencia `RevitAPI.dll` y `RevitAPIUI.dll` de tu versión de Revit.
3. Registra el ensamblado en tu `.addin` de Revit.

## Nota de migración

Las automatizaciones se han pasado a una arquitectura nativa C# unificada, manteniendo los mismos comandos funcionales del set pyRevit para ejecución desde Revit Add-in.
