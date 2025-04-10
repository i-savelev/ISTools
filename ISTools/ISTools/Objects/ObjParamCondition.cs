using Autodesk.Revit.DB;

namespace ISTools
{
    public class ObjParamCondition
    {
        public string BoolCondition { get; set; } = null;
        public string ParamName { get; set; } = null;
        public string Condition { get; set; } = null;
        public string ParamValue { get; set; } = null;


        public ObjParamCondition(string boolCondition, string paramName, string condition, string paramValue)
        {
            BoolCondition = boolCondition;
            ParamName = paramName;
            Condition = condition;
            ParamValue = paramValue;
        }
        public ObjParamCondition()
        {
        }

        public string[] GetStrings()
        {
            string[] strings = new string[4];
            strings[0] = BoolCondition;
            strings[1] = ParamName;
            strings[2] = Condition;
            strings[3] = ParamValue;
            return strings;
        }
    }
}