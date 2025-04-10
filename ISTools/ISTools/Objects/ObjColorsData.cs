using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ISTools.MyObjects
{
    public class ObjColorsData
    {
        public List<ObjCategory> CategoryList = new List<ObjCategory>();

        public List<ObjColorByKey> ColorList = new List<ObjColorByKey>();

        public string ParameterName {  get; set; }

        public ObjColorsData() { }
    }
}
