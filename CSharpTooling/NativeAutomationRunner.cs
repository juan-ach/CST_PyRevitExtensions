using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;

namespace CST.PyRevitCSharpTooling
{
    internal static class NativeAutomationRunner
    {
        public static Result Execute(string commandKey, ExternalCommandData commandData, ref string message)
        {
            try
            {
                var uiDoc = commandData.Application.ActiveUIDocument;
                var doc = uiDoc.Document;

                if (commandKey.Contains("Pin_All")) return PinElements(doc, true);
                if (commandKey.Contains("Unpin_All")) return PinElements(doc, false);

                if (commandKey.Contains("Labels_pulldown_Delete") || commandKey.Contains("Delete_Evaporadores") ) return DeleteTags(doc, doc.ActiveView);
                if (commandKey.Contains("Rooms_pulldown_Delete")) return DeleteByCategory(doc, BuiltInCategory.OST_Rooms, "Delete Rooms");
                if (commandKey.Contains("Circuits_pulldown_Delete")) return DeleteByCategory(doc, BuiltInCategory.OST_ElectricalCircuit, "Delete Circuits");
                if (commandKey.Contains("Insulator_pulldown_Delete")) return DeleteByCategory(doc, BuiltInCategory.OST_PipeInsulations, "Delete Insulations");

                if (commandKey.Contains("CotasLayersGreyOff")) return ApplyDwgViewTemplateCleanup(doc);

                if (commandKey.Contains("Labels_pulldown_Create")) return CreateMechanicalTags(doc, doc.ActiveView, commandKey);
                if (commandKey.Contains("seguridades_pushbutton")) return CreateMechanicalTags(doc, doc.ActiveView, commandKey, "seguridad");
                if (commandKey.Contains("Diameter_Checker")) return DiameterChecker(doc);

                if (commandKey.Contains("Insulator_pulldown_Transcritico")) return ApplyInsulationProfile(doc, InsulationProfiles.Transcritico());
                if (commandKey.Contains("Insulator_pulldown_Glicol")) return ApplyInsulationProfile(doc, InsulationProfiles.Glicol());
                if (commandKey.Contains("Insulator_pulldown_448A")) return ApplyInsulationProfile(doc, InsulationProfiles.A448());
                if (commandKey.Contains("Insulator_pulldown_134_448EVI")) return ApplyInsulationProfile(doc, InsulationProfiles.A134448Evi());

                if (commandKey.Contains("Rooms_pulldown_Create")) return EnsureRoomTags(doc);
                if (commandKey.Contains("Rooms_pulldown_Ubi_Evap_Door")) return SetRoomLocationParams(doc);
                if (commandKey.Contains("Circuits_pulldown_Create")) return BuildCircuitsByPanelAndLoad(doc);
                if (commandKey.Contains("Antes_Presto")) return PrestoPreparation(doc);

                TaskDialog.Show("CST Tooling", "Comando no mapeado: " + commandKey);
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.ToString();
                TaskDialog.Show("CST Tooling", "Error en comando nativo C#: \n\n" + ex.Message);
                return Result.Failed;
            }
        }

        private static Result PinElements(Document doc, bool pin)
        {
            int count = 0;
            var ids = new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType().ToElementIds();
            using (var tx = new Transaction(doc, pin ? "Pin All" : "Unpin All"))
            {
                tx.Start();
                foreach (var id in ids)
                {
                    var e = doc.GetElement(id);
                    if (e == null) continue;
                    try
                    {
                        if (e.Pinned != pin)
                        {
                            e.Pinned = pin;
                            count++;
                        }
                    }
                    catch { }
                }
                tx.Commit();
            }
            TaskDialog.Show("CST Tooling", (pin ? "Pinned" : "Unpinned") + $" {count} elementos.");
            return Result.Succeeded;
        }

        private static Result DeleteTags(Document doc, View view)
        {
            var bics = new[] { BuiltInCategory.OST_Tags, BuiltInCategory.OST_MechanicalEquipmentTags, BuiltInCategory.OST_PipeTags, BuiltInCategory.OST_DuctTags, BuiltInCategory.OST_ElectricalEquipmentTags, BuiltInCategory.OST_ElectricalCircuitTags, BuiltInCategory.OST_AreaTags, BuiltInCategory.OST_RoomTags };
            int deleted = 0;
            using (var tx = new Transaction(doc, "Delete Tags"))
            {
                tx.Start();
                foreach (var bic in bics)
                {
                    var ids = new FilteredElementCollector(doc, view.Id).OfCategory(bic).WhereElementIsNotElementType().ToElementIds();
                    if (ids.Count > 0) deleted += doc.Delete(ids).Count;
                }
                tx.Commit();
            }
            TaskDialog.Show("CST Tooling", $"Tags eliminados: {deleted}");
            return Result.Succeeded;
        }

        private static Result DeleteByCategory(Document doc, BuiltInCategory bic, string name)
        {
            var ids = new FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElementIds();
            using (var tx = new Transaction(doc, name))
            {
                tx.Start();
                if (ids.Count > 0) doc.Delete(ids);
                tx.Commit();
            }
            TaskDialog.Show("CST Tooling", $"{name}: {ids.Count} elementos.");
            return Result.Succeeded;
        }

        private static Result ApplyDwgViewTemplateCleanup(Document doc)
        {
            var templates = new HashSet<string>{"CST_FLO_ELE_Trazado","CST_FLO_HVAC","CST_FLO_SEG v1","CST_FLO_Servicios","CST_FLO_TUB_ESQ v2","CST_FLO_CO2 liq com v2","CST_FLO_Servicios BP","CST_FLO_SERV","CST_FLO_GLI","CST_FLO_CO2 liq com v3","CST_FLO_TUB LIQ COM2","CST_FLO_GLI_448","CST_FLO_TUB"};
            var hideTokens = new[] { "tex", "txt", "cota", "implan", "seccions", "cotes" };
            var gray = new Color(128, 128, 128);
            var overrideSettings = new OverrideGraphicSettings().SetProjectionLineColor(gray).SetProjectionLineWeight(1);

            using (var tx = new Transaction(doc, "Turn Off/Grey Layers"))
            {
                tx.Start();
                var views = new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>().Where(v=>v.IsTemplate && templates.Contains(v.Name));
                var imports = new FilteredElementCollector(doc).OfClass(typeof(ImportInstance)).Cast<ImportInstance>().ToList();
                foreach (var vt in views)
                foreach (var imp in imports)
                {
                    var cat = imp.Category;
                    if (cat?.SubCategories == null) continue;
                    foreach (Category sub in cat.SubCategories)
                    {
                        var lname = (sub.Name ?? string.Empty).ToLowerInvariant();
                        bool hide = hideTokens.Any(t => lname.Contains(t));
                        if (hide) vt.SetCategoryHidden(sub.Id, true);
                        else vt.SetCategoryOverrides(sub.Id, overrideSettings);
                    }
                }
                tx.Commit();
            }
            TaskDialog.Show("CST Tooling", "Capes DWG depurades.");
            return Result.Succeeded;
        }

        private static Result CreateMechanicalTags(Document doc, View view, string commandKey, string forceKeyword = null)
        {
            int created = 0;
            FamilySymbol tagType = FindTagType(doc, "CST_TAG Equipos Mecanicos v5", commandKey.Contains("Esquema_Frigorifico") ? "evapor" : null)
                                  ?? FindAnyTagType(doc);
            if (tagType == null) { TaskDialog.Show("CST Tooling", "No se encontró tipo de etiqueta."); return Result.Failed; }

            var equipments = new FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements();
            using (var tx = new Transaction(doc, "Create Labels"))
            {
                tx.Start();
                foreach (var e in equipments)
                {
                    var me = e as FamilyInstance;
                    if (me == null) continue;
                    var name = (me.Name + " " + me.Symbol?.FamilyName).ToLowerInvariant();
                    if (!string.IsNullOrEmpty(forceKeyword) && !name.Contains(forceKeyword)) continue;

                    var bb = e.get_BoundingBox(view);
                    XYZ p = bb != null ? (bb.Min + bb.Max) / 2.0 : ((LocationPoint)e.Location)?.Point;
                    if (p == null) continue;
                    var tag = IndependentTag.Create(doc, tagType.Id, view.Id, new Reference(e), false, TagOrientation.Horizontal, p + new XYZ(0.2,0.2,0));
                    if (tag != null) created++;
                }
                tx.Commit();
            }
            TaskDialog.Show("CST Tooling", $"Etiquetas creadas: {created}");
            return Result.Succeeded;
        }

        private static FamilySymbol FindAnyTagType(Document doc)
        {
            return new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().FirstOrDefault(x => x.Category != null && x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_MechanicalEquipmentTags);
        }

        private static FamilySymbol FindTagType(Document doc, string familyName, string typeNameContains)
        {
            return new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .FirstOrDefault(s => s.Category != null
                                     && s.Category.Id.IntegerValue == (int)BuiltInCategory.OST_MechanicalEquipmentTags
                                     && (s.FamilyName ?? string.Empty).Contains(familyName)
                                     && (string.IsNullOrEmpty(typeNameContains) || (s.Name ?? string.Empty).ToLowerInvariant().Contains(typeNameContains.ToLowerInvariant())));
        }

        private static Result EnsureRoomTags(Document doc)
        {
            var rooms = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().Cast<SpatialElement>().ToList();
            int count = 0;
            using (var tx = new Transaction(doc, "Create Rooms"))
            {
                tx.Start();
                foreach (Room room in rooms)
                {
                    if (room.Location is LocationPoint lp)
                    {
                        var uv = new UV(lp.Point.X, lp.Point.Y);
                        var tag = doc.Create.NewRoomTag(new LinkElementId(room.Id), uv, doc.ActiveView.Id);
                        if (tag != null) count++;
                    }
                }
                tx.Commit();
            }
            TaskDialog.Show("CST Tooling", $"Rooms tageadas: {count}");
            return Result.Succeeded;
        }

        private static Result SetRoomLocationParams(Document doc)
        {
            var rooms = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().Cast<SpatialElement>().ToList();
            var doors = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements();
            int updated = 0;
            using (var tx = new Transaction(doc, "Evap/Door Location"))
            {
                tx.Start();
                foreach (Room room in rooms)
                {
                    var roomObj = room as Room;
                    if (roomObj == null) continue;
                    var p = roomObj.LookupParameter("Ubicación puerta") ?? roomObj.LookupParameter("Ubicacion puerta");
                    if (p == null || p.IsReadOnly) continue;
                    bool hasDoor = doors.Any(d => roomObj.IsPointInRoom(((d.Location as LocationPoint)?.Point) ?? XYZ.Zero));
                    p.Set(hasDoor ? "Con puerta" : "Sin puerta");
                    updated++;
                }
                tx.Commit();
            }
            TaskDialog.Show("CST Tooling", $"Parámetro ubicación actualizado en {updated} salas");
            return Result.Succeeded;
        }

        private static Result DiameterChecker(Document doc)
        {
            var eq = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements();
            int missing = 0;
            foreach (var e in eq)
            {
                var p = e.LookupParameter("Diámetro") ?? e.LookupParameter("Diametro");
                if (p == null || string.IsNullOrWhiteSpace(p.AsValueString())) missing++;
            }
            TaskDialog.Show("CST Tooling", $"Equipos revisados: {eq.Count}. Sin diámetro: {missing}");
            return Result.Succeeded;
        }

        private static Result BuildCircuitsByPanelAndLoad(Document doc)
        {
            var loads = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();
            var panels = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();
            int created = 0;
            using (var tx = new Transaction(doc, "Create Circuits"))
            {
                tx.Start();
                foreach (var load in loads)
                {
                    var panelName = (load.LookupParameter("Cuadro")?.AsString() ?? string.Empty).Trim();
                    if (string.IsNullOrEmpty(panelName)) continue;
                    var panel = panels.FirstOrDefault(p => (p.Name ?? string.Empty).Contains(panelName));
                    if (panel == null) continue;

                    var conn = load.MEPModel?.ConnectorManager?.Connectors?.Cast<Connector>().FirstOrDefault(c => c.Domain == Domain.DomainElectrical);
                    if (conn == null) continue;
                    var system = ElectricalSystem.Create(conn, ElectricalSystemType.PowerCircuit);
                    if (system == null) continue;
                    system.SelectPanel(panel);
                    created++;
                }
                tx.Commit();
            }
            TaskDialog.Show("CST Tooling", $"Circuitos creados: {created}");
            return Result.Succeeded;
        }

        private static Result ApplyInsulationProfile(Document doc, InsulationProfile profile)
        {
            var insTypes = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeInsulations).WhereElementIsElementType().Cast<ElementType>()
                .ToDictionary(t => t.Name, t => t, StringComparer.OrdinalIgnoreCase);
            var existingHosts = new HashSet<int>(new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeInsulations).WhereElementIsNotElementType().Cast<PipeInsulation>().Select(i => i.HostElementId.IntegerValue));
            var pipes = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements();
            int created=0, skipped=0;

            using (var tx = new Transaction(doc, profile.Name))
            {
                tx.Start();
                foreach (var p in pipes)
                {
                    if (existingHosts.Contains(p.Id.IntegerValue)) { skipped++; continue; }
                    var sys = p.LookupParameter("Tipo de sistema")?.AsString() ?? p.LookupParameter("System Type")?.AsValueString() ?? string.Empty;
                    bool positive = profile.PositiveSystems.Contains(sys);
                    bool negative = profile.NegativeSystems.Contains(sys);
                    if (!positive && !negative) { skipped++; continue; }

                    var dia = (p.LookupParameter("Diámetro")?.AsValueString() ?? p.LookupParameter("Diameter")?.AsValueString() ?? string.Empty).Trim();
                    var map = negative ? profile.NegativeByDiameter : profile.PositiveByDiameter;
                    if (!map.TryGetValue(dia, out var targetTypeName)) { skipped++; continue; }
                    if (!insTypes.TryGetValue(targetTypeName, out var type)) { skipped++; continue; }

                    double thicknessMm = negative ? profile.NegativeThicknessMm : profile.PositiveThicknessMm;
                    double internalUnits = UnitUtils.ConvertToInternalUnits(thicknessMm, UnitTypeId.Millimeters);
                    PipeInsulation.Create(doc, p.Id, type.Id, internalUnits);
                    created++;
                }
                tx.Commit();
            }
            TaskDialog.Show("CST Tooling", $"{profile.Name}: creados {created}, saltados {skipped}");
            return Result.Succeeded;
        }

        private static Result PrestoPreparation(Document doc)
        {
            var map = new Dictionary<BuiltInCategory, string>
            {
                {BuiltInCategory.OST_PipeCurves, "01.08"},
                {BuiltInCategory.OST_DuctCurves, "01.09"},
                {BuiltInCategory.OST_DuctFitting, "01.09"},
                {BuiltInCategory.OST_PipeInsulations, "04"},
                {BuiltInCategory.OST_Walls, "09"},
                {BuiltInCategory.OST_Floors, "09"}
            };
            int changed=0;
            using (var tx = new Transaction(doc, "Antes de mandar a PRESTO"))
            {
                tx.Start();
                foreach (var kv in map)
                {
                    foreach (var e in new FilteredElementCollector(doc).OfCategory(kv.Key).WhereElementIsNotElementType())
                    {
                        var p = e.LookupParameter("Partidas_PRESTO");
                        if (p != null && !p.IsReadOnly) { p.Set(kv.Value); changed++; }
                    }
                }
                tx.Commit();
            }
            TaskDialog.Show("CST Tooling", $"Antes PRESTO completado. Parámetros actualizados: {changed}");
            return Result.Succeeded;
        }
    }

    internal class InsulationProfile
    {
        public string Name { get; set; }
        public HashSet<string> PositiveSystems { get; set; } = new HashSet<string>();
        public HashSet<string> NegativeSystems { get; set; } = new HashSet<string>();
        public Dictionary<string, string> PositiveByDiameter { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> NegativeByDiameter { get; set; } = new Dictionary<string, string>();
        public double PositiveThicknessMm { get; set; } = 19;
        public double NegativeThicknessMm { get; set; } = 32;
    }

    internal static class InsulationProfiles
    {
        public static InsulationProfile Transcritico() => new InsulationProfile
        {
            Name = "Transcrítico",
            PositiveSystems = new HashSet<string> { "A1+", "A2+", "A3+", "L1", "L2", "L3", "L4", "L5" },
            NegativeSystems = new HashSet<string> { "A1-", "A2-", "A3-" },
            PositiveByDiameter = Common19(),
            NegativeByDiameter = Common32()
        };

        public static InsulationProfile Glicol() => new InsulationProfile
        {
            Name = "Glicol",
            PositiveSystems = new HashSet<string> { "GLICOL IMPULSION", "GLICOL RETORNO", "G1", "G2" },
            NegativeSystems = new HashSet<string>(),
            PositiveByDiameter = Common32(),
            PositiveThicknessMm = 32,
            NegativeThicknessMm = 32
        };

        public static InsulationProfile A448() => new InsulationProfile
        {
            Name = "448A",
            PositiveSystems = new HashSet<string> { "MT", "LT", "IMPULSION", "RETORNO" },
            NegativeSystems = new HashSet<string>(),
            PositiveByDiameter = Common19(),
            PositiveThicknessMm = 19
        };

        public static InsulationProfile A134448Evi() => new InsulationProfile
        {
            Name = "134-448EVI",
            PositiveSystems = new HashSet<string> { "EVI", "ALTA", "BAJA" },
            NegativeSystems = new HashSet<string>(),
            PositiveByDiameter = Common19(),
            PositiveThicknessMm = 19
        };

        private static Dictionary<string, string> Common19() => new Dictionary<string, string>
        {
            {"1/4", "_AISLAMIENTO INSTAL. TUBERÍA COBRE 1/4 - 19mm"},
            {"3/8", "_AISLAMIENTO INSTAL. TUBERÍA COBRE 3/8 - 19mm"},
            {"1/2", "_AISLAMIENTO INSTAL. TUBERÍA COBRE 1/2 - 19mm"},
            {"5/8", "_AISLAMIENTO INSTAL. TUBERÍA COBRE 5/8 - 19mm"},
            {"3/4", "_AISLAMIENTO INSTAL. TUBERÍA COBRE 3/4 - 19mm"},
            {"7/8", "_AISLAMIENTO INSTAL. TUBERÍA COBRE 7/8 - 19mm"},
            {"1 1/8", "_AISLAMIENTO INSTAL. TUBERÍA COBRE 1 1/8 - 19mm"},
            {"1 3/8", "_AISLAMIENTO INSTAL. TUBERÍA COBRE 1 3/8 - 19mm"},
            {"1 5/8", "_AISLAMIENTO INSTAL. TUBERÍA COBRE 1 5/8 - 19mm"},
            {"2 1/8", "_AISLAMIENTO INSTAL. TUBERÍA COBRE 2 1/8 - 19mm"},
            {"2 5/8", "_AISLAMIENTO INSTAL. TUBERÍA COBRE 2 5/8 - 19mm"}
        };

        private static Dictionary<string, string> Common32() => new Dictionary<string, string>
        {
            {"1/4", "AISLAMIENTO INSTAL. TUBERÍA COBRE 1/4 - 32mm"},
            {"3/8", "AISLAMIENTO INSTAL. TUBERÍA COBRE 3/8 - 32mm"},
            {"1/2", "AISLAMIENTO INSTAL. TUBERÍA COBRE 1/2 - 32mm"},
            {"5/8", "AISLAMIENTO INSTAL. TUBERÍA COBRE 5/8 - 32mm"},
            {"3/4", "AISLAMIENTO INSTAL. TUBERÍA COBRE 3/4 - 32mm"},
            {"7/8", "AISLAMIENTO INSTAL. TUBERÍA COBRE 7/8 - 32mm"},
            {"1 1/8", "AISLAMIENTO INSTAL. TUBERÍA COBRE 1 1/8 - 32mm"},
            {"1 3/8", "AISLAMIENTO INSTAL. TUBERÍA COBRE 1 3/8 - 32mm"},
            {"1 5/8", "AISLAMIENTO INSTAL. TUBERÍA COBRE 1 5/8 - 32mm"},
            {"2 1/8", "AISLAMIENTO INSTAL. TUBERÍA COBRE 2 1/8 - 32mm"},
            {"2 5/8", "AISLAMIENTO INSTAL. TUBERÍA COBRE 2 5/8 - 32mm"}
        };
    }
}
