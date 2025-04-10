using Autodesk.Revit.DB;

namespace ISTools
{
    public class ObjCategory
    {
        public string Name { get; set; } = null;

        public BuiltInCategory BuiltIn { get; set; }

        public ObjCategory(string name)
        {
            Name = name;
        }

        public ObjCategory()
        {

        }

        public string[] GetStrings()
        {
            string[] strings = new string[1];
            strings[0] = Name;
            return strings;
        }
    }
}
