using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ISTools.MyObjects
{
    public class ObjWorkset
    {
        public string Name {  get; set; }

        public List<ObjParamCondition> Conditions { get; set; }

        public List<ObjCategory> Categories { get; set; }

        public string CategoriesAndParams { get; set; } = "и";

        public ObjWorkset(string name) 
        {
            Name = name;
        }

        public ObjWorkset()
        {
        }

        public ObjWorkset Clone(ObjWorkset other)
        {
            ObjWorkset obj = new ObjWorkset();
            obj.Conditions = new List<ObjParamCondition>(other.Conditions);
            obj.Name = other.Name+"(1)";
            obj.Categories = new List<ObjCategory>(other.Categories);
            obj.CategoriesAndParams = other.CategoriesAndParams;
            return obj;
        }
    }
}
