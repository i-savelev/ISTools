using System;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Collections.Generic;
using RVTDocument = Autodesk.Revit.DB.Document;
using Autodesk.Revit.DB.Architecture;



namespace ISTools
{
    /// <summary>
    /// a universal class for any element in revit.
    /// </summary>
    public class ObjRvt
    {
        public Element elem;
        
        /// <summary>
        /// a method that return parameter value by parameter name
        /// </summary>
        public virtual object GetParam(string parName)
        {
            Element type = elem.Document.GetElement(elem.GetTypeId());
            if (elem.LookupParameter(parName) != null)
            {
                if (elem.LookupParameter(parName).StorageType.ToString() == "String")
                {
                    return elem.LookupParameter(parName).AsValueString();
                }
                else if (elem.LookupParameter(parName).StorageType.ToString() == "Double")
                {
                    return elem.LookupParameter(parName).AsDouble();
                }
                else if (elem.LookupParameter(parName).StorageType.ToString() == "Integer")
                {
                    return elem.LookupParameter(parName).AsInteger();
                }
                else if (elem.LookupParameter(parName).StorageType.ToString() == "ElementId")
                {
                    return elem.LookupParameter(parName).AsElementId();
                }
                else return null;
            }
            else if (type != null)
            {
                if (type.LookupParameter(parName) != null)
                {
                    if (type.LookupParameter(parName).StorageType.ToString() == "String")
                    {
                        if (type.LookupParameter(parName).AsString() != null)
                        {
                            return type.LookupParameter(parName).AsValueString();
                        }
                        return type.LookupParameter(parName).AsString();
                    }
                    else if (type.LookupParameter(parName).StorageType.ToString() == "Double")
                    {
                        return type.LookupParameter(parName).AsDouble();
                    }
                    else if (type.LookupParameter(parName).StorageType.ToString() == "Integer")
                    {
                        return type.LookupParameter(parName).AsInteger();
                    }
                    else if (type.LookupParameter(parName).StorageType.ToString() == "ElementId")
                    {
                        return type.LookupParameter(parName).AsElementId();
                    }
                    else return null;
                }
                else return null;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// a method that return a messedge that the parameter does not have a parameter value
        /// </summary>
        public virtual string CheckParam(string parName)
        {
            try
            {
                ElementId elemId = (ElementId)GetParam(parName);  // проверка материала
                if (elemId.ToString() == "-1")
                {
                    return $"«{parName}» не заполнен";
                }
                else return null;
            }
            catch
            {
                if (GetParam(parName) != null)
                {
                    if (GetParam(parName).ToString() == "")
                    {
                        return $"«{parName}» не заполнен";
                    }
                    else if (GetParam(parName).ToString() == "0")
                    {
                        return $"«{parName}» не заполнен";
                    }
                    else return null;
                }
                else return $"«{parName}» не заполнен";
            }
        }

        /// <summary>
        /// a method that set parameter value by parameter name
        /// </summary>
        virtual public void SetParam(string parName, object parValue)
        {
            if (elem.LookupParameter(parName) != null)
            {
                if (elem.LookupParameter(parName).StorageType.ToString() == "String")
                {
                    elem.LookupParameter(parName).Set(parValue.ToString());
                }
                if (elem.LookupParameter(parName).StorageType.ToString() == "Double")
                {
                    var a = Convert.ToDouble(parValue);
                    elem.LookupParameter(parName).Set(a);
                }
                if (elem.LookupParameter(parName).StorageType.ToString() == "Integer")
                {
                    var a = Convert.ToInt32(parValue);
                    elem.LookupParameter(parName).Set(a);
                }
            }
        }
        public virtual string CheckWorkSet(Dictionary<string, string> worksetDictByCat, Dictionary<string, string> worksetDictByFamyly)
        {
            string check = null;
            var elemTypeId = elem.GetTypeId();
            var elemType = elem.Document.GetElement(elemTypeId) as ElementType;
            try
            {
                if (worksetDictByFamyly.ContainsKey(elemType.FamilyName.Split('_')[0]) || worksetDictByFamyly.ContainsKey(elem.Name.Split('_')[0]))
                {
                    if (worksetDictByFamyly.ContainsKey(elemType.FamilyName.Split('_')[0]))
                    {
                        var val = worksetDictByFamyly[elemType.FamilyName.Split('_')[0]];
                        if (!elem.LookupParameter("Рабочий набор").AsValueString().Contains(val))
                        {
                            check = $"Элемент расположен в рабочем наборе «{elem.LookupParameter("Рабочий набор").AsValueString()}» должен находится в рабочем наборе с «{worksetDictByFamyly[elemType.FamilyName.Split('_')[0]]}» в названии";
                        }
                    }
                    else if (worksetDictByFamyly.ContainsKey(elem.Name.Split('_')[0]))
                    {
                        var val = worksetDictByFamyly[elem.Name.Split('_')[0]];
                        if (!elem.LookupParameter("Рабочий набор").AsValueString().Contains(val))
                        {
                            check = $"Элемент расположен в рабочем наборе «{elem.LookupParameter("Рабочий набор").AsValueString()}» должен находится в рабочем наборе с «{worksetDictByFamyly[elem.Name.Split('_')[0]]}» в названии";
                        }
                    }
                }
                else if (worksetDictByCat.ContainsKey(((BuiltInCategory)elem.Category.Id.IntegerValue).ToString()))
                {
                  
                    if (!elem.LookupParameter("Рабочий набор").AsValueString().Contains(worksetDictByCat[((BuiltInCategory)elem.Category.Id.IntegerValue).ToString()]))
                    {
                        check = $"Элемент расположен в рабочем наборе «{elem.LookupParameter("Рабочий набор").AsValueString()}» должен находится в рабочем наборе с «{worksetDictByCat[((BuiltInCategory)elem.Category.Id.IntegerValue).ToString()]}» в названии";
                    }
                }
            }
            catch
            {
            }

            return check;
        }
        /// <summary>
        /// a method that return parameter value by parameter name
        /// </summary>
        public virtual string GetParamAsString(string parName)
        {
            Element type = elem.Document.GetElement(elem.GetTypeId());
            if (elem.LookupParameter(parName) != null)
            {
                if (elem.LookupParameter(parName).StorageType.ToString() == "String")
                {
                    if (elem.LookupParameter(parName).AsValueString() == null)
                    { return ""; }
                    return elem.LookupParameter(parName).AsValueString();
                }
                else if (elem.LookupParameter(parName).StorageType.ToString() == "Double")
                {
                    return elem.LookupParameter(parName).AsDouble().ToString();
                }
                else if (elem.LookupParameter(parName).StorageType.ToString() == "Integer")
                {
                    return elem.LookupParameter(parName).AsValueString();
                }
                else if (elem.LookupParameter(parName).StorageType.ToString() == "ElementId")
                {
                    return elem.LookupParameter(parName).AsValueString();
                }
                else return "";
            }
            else if (type != null)
            {
                if (type.LookupParameter(parName) != null)
                {
                    if (type.LookupParameter(parName).StorageType.ToString() == "String")
                    {
                        if (type.LookupParameter(parName).AsString() != null)
                        {
                            return type.LookupParameter(parName).AsValueString();
                        }
                        return type.LookupParameter(parName).AsValueString();
                    }
                    else if (type.LookupParameter(parName).StorageType.ToString() == "Double")
                    {
                        return type.LookupParameter(parName).AsValueString();
                    }
                    else if (type.LookupParameter(parName).StorageType.ToString() == "Integer")
                    {
                        return type.LookupParameter(parName).AsValueString();
                    }
                    else if (type.LookupParameter(parName).StorageType.ToString() == "ElementId")
                    {
                        return type.LookupParameter(parName).AsValueString();
                    }
                    else return null;
                }
                else return "";
            }
            else
            {
                return "";
            }
        }
        public void Set_element_color(Color color)
        {
            FillPatternElement solid_pattern = null;
            var all_patterns = new FilteredElementCollector(elem.Document).OfClass(typeof(FillPatternElement)).ToElements();
            foreach (FillPatternElement pattern in all_patterns)
            {
                if (pattern.GetFillPattern().IsSolidFill)
                {
                    solid_pattern = pattern;
                    break;
                }
            }

            var active_view =elem.Document.ActiveView;
            var override_settings = new OverrideGraphicSettings();
            override_settings.SetSurfaceForegroundPatternColor(color);
            override_settings.SetCutForegroundPatternId(solid_pattern.Id);
            override_settings.SetCutForegroundPatternColor(color);
            override_settings.SetSurfaceForegroundPatternId(solid_pattern.Id);
            active_view.SetElementOverrides(elem.Id, override_settings);
        }

        public void Clear_overrides()
        {
            try
            {
                var active_view = elem.Document.ActiveView;
                var override_settings = new OverrideGraphicSettings();
                active_view.SetElementOverrides(elem.Id, override_settings);
            }
            catch { }
            
        }
    }
}
