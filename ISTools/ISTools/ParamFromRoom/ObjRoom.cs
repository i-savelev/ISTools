using System;
using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.Architecture;
using Plugin.MyObjects;




namespace Plugin
{
    /// <summary>
    /// a universal class for any element in revit.
    /// </summary>
    internal class ObjRoom : ObjRvt
    {
        public Transform transform;
        public Room room;
        public Document roomDoc;
        public Document doc;
        public GraphicsStyle linestyle;


        /// <summary>
        ///  СОЗДАНИЕ ГЕОМЕТРИИ ПОМЕЩЕНИЙ
        /// </summary>
        public Solid GetRoomSolid(double offset, double thickness)
        {
            var bSegments = room.GetBoundarySegments(new SpatialElementBoundaryOptions());
            var curveLoop = new CurveLoop();
            var translation = Transform.CreateTranslation(new XYZ(0, 0, -offset));
            foreach (var segment in bSegments.First())
            {
                curveLoop.Append(segment.GetCurve().CreateTransformed(transform).CreateTransformed(translation));
            }
            var offsetCurveLoop = CurveLoop.CreateViaOffset(curveLoop, offset, XYZ.BasisZ);
            var list = new List<CurveLoop>() { offsetCurveLoop };
            var roomSolid = GeometryCreationUtilities.CreateExtrusionGeometry(list, XYZ.BasisZ, room.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsDouble() - thickness + offset);
             return roomSolid;
        }
        public void SetRoomSolid(double offset, double thickness)
        {
            var ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
            var listBox = new List<GeometryObject>() { GetRoomSolid(offset, thickness) };
            ds.SetShape(listBox);
            FillPatternElement solid_pattern = null;
            var all_patterns = new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement)).ToElements();
            foreach (FillPatternElement pattern in all_patterns)
            {
                if (pattern.GetFillPattern().IsSolidFill)
                {
                    solid_pattern = pattern;
                    break;
                }
            }
            var View = doc.ActiveView;
            var color = new Color(255, 0, 0);
            var override_settings = new OverrideGraphicSettings();
            override_settings.SetSurfaceForegroundPatternColor(color);
            override_settings.SetCutForegroundPatternId(solid_pattern.Id);
            override_settings.SetCutForegroundPatternColor(color);  
            override_settings.SetSurfaceTransparency(50);
            override_settings.SetSurfaceForegroundPatternId(solid_pattern.Id);
            ds.LookupParameter("Марка").Set($"##room_{room.Name}-{room.Number}");
            View.SetElementOverrides(ds.Id, override_settings);
        }

        public void SetRoomCurve53(double offset, Random rnd, int lineWidth, List<string> checklist, List<Obj_Finishing> finishing_list, List<Obj_Finishing_Group> finishing_group_list)
        {
            SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
            options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;
            options.StoreFreeBoundaryFaces = true;
            IList<IList<BoundarySegment>> bSegments = room.GetBoundarySegments(options);
            List<string> debug_list = new List<string>();

            Dictionary<string, string> rough_finishingDict = new Dictionary<string, string>
                {
                    {"ГБ", "CPI_Черновая-ГБ_Описание" },
                    {"Каркас", "CPI_Черновая-Каркас_Описание" },
                    {"КР", "CPI_Черновая-КР_Описание" },
                    {"ЖБ", "CPI_Черновая-ЖБ_Описание" },
                };
            foreach (var segments in bSegments)
            {
                foreach (var segment in segments)
                {
                    var elId = segment.ElementId;
                    if (elId.ToString() != "-1")
                    {
                        var el = roomDoc.GetElement(elId);
                        try
                        {
                            ObjRvt obj = new ObjRvt();
                            obj.elem = el;
                            var rough_finishing = obj.GetParam("CPI_Основа черновой отделки");
                            

                            debug_list.Add($"{(string)rough_finishing} --- {obj.elem.Name}");

                            if (rough_finishing != null)
                            {
                                if (checklist.Contains($"Черновая: {rough_finishing}") & (string)rough_finishing != "-")
                                {
                                    var rough_finishing_Description = room.LookupParameter(rough_finishingDict[rough_finishing.ToString()]).AsString().Replace("\n", " ").Replace("\r", " ").Replace("\t", " ");
                                    var discriptionKey = $"Черновая: {rough_finishing}&{rough_finishing_Description}";
                                    Color color = null;
                                    if (!finishing_list.Any(p => p.Discription_key == discriptionKey))
                                    {
                                        color = new Color((byte)rnd.Next(255), (byte)rnd.Next(255), (byte)rnd.Next(255));
                                        int R = color.Red;
                                        int G = color.Green;
                                        int B = color.Blue;
                                        System.Drawing.Color system_color = System.Drawing.Color.FromArgb(R, G, B);
                                        Color rvt_color = new Color((byte)R, (byte)G, (byte)B);
                                        Obj_Finishing objFinishing = new Obj_Finishing(discriptionKey, system_color);
                                        finishing_list.Add(objFinishing);
                                    }
                                    else
                                    {
                                        var finishing = finishing_list.FirstOrDefault(p => p.Discription_key == discriptionKey);
                                        System.Drawing.Color system_color = finishing.Group.Curve_color();
                                        int R = system_color.R;
                                        int G = system_color.G;
                                        int B = system_color.B;
                                        Color rvt_color = new Color((byte)R, (byte)G, (byte)B);
                                        color = rvt_color;
                                    }
                                    if (finishing_list.FirstOrDefault(p => p.Discription_key == discriptionKey).Curve_create)
                                    {
                                        Curve curve = segment.GetCurve().CreateTransformed(transform);
                                        var line = room.Document.Create.NewDetailCurve(room.Document.ActiveView, curve);
                                        line.LineStyle = linestyle;
                                        var line_id = line.Id;
                                        var View = room.Document.ActiveView;
                                        var override_settings = new OverrideGraphicSettings();
                                        override_settings.SetProjectionLineColor(color);
                                        override_settings.SetProjectionLineWeight(lineWidth);
                                        View.SetElementOverrides(line_id, override_settings);
                                    }
                                }
                            }

                            if (checklist.Contains($"Чистовая: {rough_finishing}"))
                            {
                                var fine_finishing_Description = "";
                                if (room.LookupParameter("CPI_Чистовая_Описание отделки") != null)
                                {
                                    fine_finishing_Description = room.LookupParameter("CPI_Чистовая_Описание отделки").AsString().Replace("\n", " ").Replace("\r", " ").Replace("\t", " ");
                                }
                                else if (room.LookupParameter("CPI_Чистовая_Описание") != null)
                                {
                                    fine_finishing_Description = room.LookupParameter("CPI_Чистовая_Описание").AsString().Replace("\n", " ").Replace("\r", " ").Replace("\t", " ");
                                }
                                if (fine_finishing_Description != null)
                                {
                                    var discriptionKey = $"Чистовая: {rough_finishing}&{fine_finishing_Description}";
                                    Color color = null;
                                    if (!finishing_list.Any(p => p.Discription_key == discriptionKey))
                                    {
                                        color = new Color((byte)rnd.Next(255), (byte)rnd.Next(255), (byte)rnd.Next(255));

                                        int R = color.Red;
                                        int G = color.Green;
                                        int B = color.Blue;
                                        System.Drawing.Color system_color = System.Drawing.Color.FromArgb(R, G, B);
                                        Color rvt_color = new Color((byte)R, (byte)G, (byte)B);
                                        Obj_Finishing objFinishing = new Obj_Finishing(discriptionKey, system_color);
                                        finishing_list.Add(objFinishing);
                                    }
                                    else
                                    {
                                        var finishing = finishing_list.FirstOrDefault(p => p.Discription_key == discriptionKey);
                                        System.Drawing.Color system_color = finishing.Group.Curve_color();
                                        int R = system_color.R;
                                        int G = system_color.G;
                                        int B = system_color.B;
                                        Color rvt_color = new Color((byte)R, (byte)G, (byte)B);
                                        color = rvt_color;
                                    }
                                    if (finishing_list.FirstOrDefault(p => p.Discription_key == discriptionKey).Curve_create)
                                    {
                                        Curve curve2 = segment.GetCurve().CreateTransformed(transform).CreateOffset(-offset, XYZ.BasisZ);
                                        var line2 = room.Document.Create.NewDetailCurve(room.Document.ActiveView, curve2);
                                        line2.LineStyle = linestyle;
                                        var b = line2.Id;
                                        var View = room.Document.ActiveView;
                                        var override_settings = new OverrideGraphicSettings();
                                        override_settings.SetProjectionLineColor(color);
                                        override_settings.SetProjectionLineWeight(lineWidth);
                                        View.SetElementOverrides(b, override_settings);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            debug_list.Add($"#Ошибка {el.Name}: {ex}");
                            string debug_result = string.Join(Environment.NewLine, debug_list);
                            //TaskDialog.Show("debug", debug_result); 
                        }
                    }
                }
            }
        }
    }
}
