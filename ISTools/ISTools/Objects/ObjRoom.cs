using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.Architecture;


namespace ISTools
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
    }
}
