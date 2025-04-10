using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ISTools.MyObjects
{
    public class ObjColorByKey
    {
        public string DiscriptionKey { get; set; } = null;
        public List<int> Color_int_list { get; set; }

        public ObjColorByKey()
        {

        }

        public System.Drawing.Color ColorFromList()
        {
            if (Color_int_list != null)
            {
                return System.Drawing.Color.FromArgb(Color_int_list[0], Color_int_list[1], Color_int_list[2]);
            }
            else
            {
                return System.Drawing.Color.White;
            }
        }
    }
}
