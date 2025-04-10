using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.Structure;


namespace ISTools
{
    internal class ObjPlateCollector
    {
        public Document doc;

        public List<ObjPlateInJoint> GetPlates()
        {
            List <ObjPlateInJoint> jointPlates = new List <ObjPlateInJoint>();
            List<StructuralConnectionHandler> joints = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(StructuralConnectionHandler))
                .Cast<StructuralConnectionHandler>()
                .ToList();

            foreach (StructuralConnectionHandler joint in joints)
            {
                List<Subelement> subelems = joint.GetSubelements().ToList();
                foreach (Subelement subelem in subelems)
                {
                    if (subelem.Category.Id == new ElementId(BuiltInCategory.OST_StructConnectionPlates))
                    {
                        ObjPlateInJoint pij = new ObjPlateInJoint();
                        pij.subelem = subelem;
                        jointPlates.Add(pij);
                    }
                }
            }
            return jointPlates;
        }
    
    }
}
