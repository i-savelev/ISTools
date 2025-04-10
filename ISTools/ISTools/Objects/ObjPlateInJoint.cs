using System;
using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AdvanceSteel.Modelling;
using Autodesk.AdvanceSteel.CADAccess;

namespace ISTools
{
    /// <summary>
    /// the specific class for plates in joints that creating using "Steel" tab in revit
    /// </summary>
    internal class ObjPlateInJoint : ObjKm
    {
        public Subelement subelem;
        public double mass;
        public double volume;
        public double thickness;
        public double width;
        public double length;
        public string mark;
        public int constructionGroup;
        public const string profileName = "Профили стальные горячекатанные. ГОСТ 19903-2015";

        /// <summary>
        /// a method that return the thickness of the steel plates in joints using advance steel api
        /// </summary>
        public double GetThickness()
        {
            FilerObject filerObj = GetFilerObj(subelem.Document, subelem.GetReference());
            Plate pl = filerObj as Plate;
            return pl.Thickness / 304.8;
        }

        /// <summary>
        /// method that return the length of the steel plates in joints using advance steel api
        /// </summary>
        public double GetLength()
        {
            FilerObject filerObj = GetFilerObj(subelem.Document, subelem.GetReference());
            Plate pl = filerObj as Plate;
            return pl.Length / 304.8;
        }

        /// <summary>
        /// method that return the width of the steel plates in joints using advance steel api
        /// </summary>
        public double GetWidth()
        {
            FilerObject filerObj = GetFilerObj(subelem.Document, subelem.GetReference());
            Plate pl = filerObj as Plate;
            return pl.Width / 304.8;
        }

        /// <summary>
        /// method that return the mass of the steel plates in joints using advance steel api
        /// </summary>
        public override double GetMass()
        {
            FilerObject filerObj = GetFilerObj(subelem.Document, subelem.GetReference());
            Plate pl = filerObj as Plate;
            ParameterValue pv = subelem.GetParameterValue(new ElementId(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM));
            ElementIdParameterValue idpv = pv as ElementIdParameterValue;
            ElementId mid = idpv.Value;
            double density = GetMaterialDensity(mid, subelem.Document);
            return pl.Width / 304.8 * pl.Length / 304.8 * pl.Thickness / 304.8 * density; ;
        }

        /// <summary>
        /// method that return a advance steel uid of the steel plates in joints
        /// </summary>
        public override Guid GetUid()
        {
            FilerObject filerObj = GetFilerObj(subelem.Document, subelem.GetReference());
            Plate pl = filerObj as Plate;
            return pl.GetUniqueId();
        }

        /// <summary>
        /// a method that set parameter value by parameter name for steel plates in joints
        /// </summary>
        public override void SetParam(string parName, object parValue)
        {
            if (subelem != null)
            {
                List<ElementId> paramIds = subelem.GetAllParameters().ToList();
                foreach (ElementId paramId in paramIds)
                {   
                    try
                    {
                        Element param = subelem.Document.GetElement(paramId);
                        if (param == null) continue;
                        if (param.Name == parName)
                        {
                            if (parValue.GetType().ToString() == "System.Double")
                            {
                                DoubleParameterValue dpv = new DoubleParameterValue((double)parValue);
                                subelem.SetParameterValue(paramId, dpv);
                            }
                            if (parValue.GetType().ToString() == "System.String")
                            {
                                StringParameterValue dpv = new StringParameterValue((string)parValue);
                                subelem.SetParameterValue(paramId, dpv);
                            }
                            if (parValue.GetType().ToString() == "System.Int32")
                            {
                                IntegerParameterValue dpv = new IntegerParameterValue((int)parValue);
                                subelem.SetParameterValue(paramId, dpv);
                            }
                        }
                    }
                    catch 
                    {
 
                    }
                }
            }  
        }

        /// <summary>
        /// a method that get parameter value by parameter name for steel plates in joints
        /// </summary>
        public override object GetParam(string parName)
        {
            object val = null;

            List<ElementId> paramIds = subelem.GetAllParameters().ToList();
            foreach (ElementId paramId in paramIds)
            {
                Element param = subelem.Document.GetElement(paramId);
                if (param == null) continue;
                if (param.Name == parName)
                {
                    ParameterValue pv = subelem.GetParameterValue(paramId);
                    if (pv.GetType().ToString() ==  "Autodesk.Revit.DB.StringParameterValue")
                    {
                        StringParameterValue spv = pv as StringParameterValue;
                        if (spv.Value == null)
                        {
                            val = "";
                        }
                        else val = spv.Value;
                    }
                    if (pv.GetType().ToString() == "Autodesk.Revit.DB.DoubleParameterValue")
                    {
                        DoubleParameterValue spv = pv as DoubleParameterValue;
                        val = spv.Value;
                    }
                    if (pv.GetType().ToString() == "Autodesk.Revit.DB.IntegerParameterValue")
                    {
                        IntegerParameterValue spv = pv as IntegerParameterValue;
                        val = spv.Value;
                    }
                    if (pv.GetType().ToString() == "Autodesk.Revit.DB.ElementIdParameterValue")
                    {
                        ElementIdParameterValue spv = pv as ElementIdParameterValue;
                        val = spv.Value;
                    }
                }            
            }
            return val;
        }

        public override string GetParamAsString(string parName)
        {
            string val = "";

            List<ElementId> paramIds = subelem.GetAllParameters().ToList();
            foreach (ElementId paramId in paramIds)
            {
                Element param = subelem.Document.GetElement(paramId);
                if (param == null) continue;
                if (param.Name == parName)
                {
                    ParameterValue pv = subelem.GetParameterValue(paramId);
                    if (pv.GetType().ToString() == "Autodesk.Revit.DB.StringParameterValue")
                    {
                        StringParameterValue spv = pv as StringParameterValue;
                        if (spv.Value == null)
                        {
                            val = "";
                        }
                        else val = spv.Value;
                    }
                    if (pv.GetType().ToString() == "Autodesk.Revit.DB.DoubleParameterValue")
                    {
                        DoubleParameterValue spv = pv as DoubleParameterValue;
                        val = spv.Value.ToString();
                    }
                    if (pv.GetType().ToString() == "Autodesk.Revit.DB.IntegerParameterValue")
                    {
                        IntegerParameterValue spv = pv as IntegerParameterValue;
                        val = spv.Value.ToString();
                    }
                    if (pv.GetType().ToString() == "Autodesk.Revit.DB.ElementIdParameterValue")
                    {
                        ElementIdParameterValue spv = pv as ElementIdParameterValue;
                        val = spv.Value.ToString();
                    }
                }
            }
            return val;
        }
    }
}
