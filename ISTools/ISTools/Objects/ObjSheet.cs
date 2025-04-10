using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace ISTools
{
    public class ObjSheet
    {
        public string Name { get; set; } = null;

        public string Number { get; set; } = null;

        public List<string> GroupList = new List<string>();

        public ViewSheet Elem { get; set; }
        
        public ObjSheet(string name, string number)
        {
            Name = name;
            Number = number;
        }

        public ObjSheet()
        {

        }
    }
}
