using System;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using RVTDocument = Autodesk.Revit.DB.Document;
using ASDocument = Autodesk.AdvanceSteel.DocumentManagement.Document;
using Autodesk.Revit.DB.Steel;
using Autodesk.AdvanceSteel.DocumentManagement;
using Autodesk.AdvanceSteel.CADAccess;
using Document = Autodesk.Revit.DB.Document;



namespace ISTools
{
    /// <summary>
    /// the specific class that contains methods for steel elements in revit
    /// </summary>
    internal class ObjKm : ObjRvt
    {
        /// <summary>
        /// method that return the mass of element that is calculated according the logic of the ADSK-KM template
        /// </summary>
        public virtual double GetMass()
        { 
            double mass = 0;
            if (GetParam("ADSK_Способ подсчета массы") != null)
            {
                switch ((int)GetParam("ADSK_Способ подсчета массы"))
                {
                    case 1:
                        if (GetParam("ADSK_Длина балки истинная") !=null & GetParam("ADSK_Масса на единицу длины") != null)
                        {
                            mass = (double)GetParam("ADSK_Длина балки истинная") * (double)GetParam("ADSK_Масса на единицу длины");
                        }
                        break;
                    case 2:
                        if (GetParam("ADSK_Размер_Длина") != null & GetParam("ADSK_Масса на единицу длины") != null)
                        {
                            mass = (double)GetParam("ADSK_Размер_Длина") * (double)GetParam("ADSK_Масса на единицу длины");
                        }
                        break;
                    case 3:
                        if (GetParam("ADSK_Размер_Толщина") != null & GetParam("ADSK_Размер_Ширина") != null & GetParam("ADSK_Размер_Длина") != null)
                        {
                            mass = (double)GetParam("ADSK_Размер_Толщина") * (double)GetParam("ADSK_Размер_Ширина") * (double)GetParam("ADSK_Размер_Длина") * 7850;
                        }
                        break;
                    case 4:
                        if (GetParam("ADSK_Размер_Толщина") != null & GetParam("ADSK_Размер_Ширина") != null & GetParam("ADSK_Размер_Длина") != null)
                        {
                            mass = 0;
                        }
                        break;
                    case 5:
                        if (GetParam("ADSK_Масса элемента") != null)
                        {
                            mass = (double)GetParam("ADSK_Масса элемента");
                        }
                        break;
                    case 6:
                        if (GetParam("ADSK_Масса на единицу площади") != null & GetParam("ABIM_Площадь Поверхности") != null)
                        {
                            mass = (double)GetParam("ABIM_Площадь Поверхности") * (double)GetParam("ADSK_Масса на единицу площади");
                        }
                        break;
                }
            }
            return mass;
        }

        /// <summary>
        /// method that return a profile name of the element
        /// </summary>
        public virtual string GetProfileName()
        {
            return $"{(string)GetParam("ADSK_Наименование профиля")}. {(string)GetParam("ADSK_Обозначение")}";
        }

        /// <summary>
        /// method that return a profile size of the element
        /// </summary>
        public virtual string GetProfileSize()
        {
            double t = Math.Round(Convert.ToDouble(GetParam("ADSK_Размер_Толщина"))* 304.8, 0);

            return $"{GetParam("ADSK_Наименование_Префикс")}{GetParam("ADSK_Наименование_Текст1")}{GetParam("ADSK_Наименование")}{t.ToString()}";
        }

        /// <summary>
        /// method that return a material of the element
        /// </summary>
        public virtual string GetMaterial()
        {
            if (GetParam("ADSK_Материал") != null)
            {
                ElementId elemId = (ElementId)GetParam("ADSK_Материал");
                Element mat = elem.Document.GetElement(elemId);
                if (elemId.ToString() == "-1" ^ elemId == null)
                {
                    return $"'ADSK_Материал' не заполнен";
                }
                else
                {
                    return $"{mat.LookupParameter("ADSK_Материал наименование").AsString()} {mat.LookupParameter("ADSK_Материал обозначение").AsString()}";
                }
            }
            else return $"'ADSK_Материал' не заполнен";
        }

        /// <summary>
        /// method that return a construction group of the element
        /// </summary>
        public virtual string GetСonstructionGroup()
        {
            if (GetParam("ADSK_Группа конструкций") != null)
            {
                switch ((int)GetParam("ADSK_Группа конструкций"))
                {
                    case 1:
                        return "Балки";
                    case 2:
                        return "Колонны";
                    case 3:
                        return "Связи";
                    case 4:
                        return "Фермы";
                    case 5:
                        return "Фахверк и прочее";
                    default:
                        return "'ADSK_Группа конструкций' не заполнен";
                }
            }
            else return "'ADSK_Группа конструкций' не заполнен";
        }

        /// <summary>
        /// method that return a mark of the element
        /// </summary>
        public virtual string GetMark()
        {
            if ((string)GetParam("ADSK_Марка конструкции") != null)
            {
                return $"{(string)GetParam("ADSK_Марка конструкции")}";
            }
            else return $"'ADSK_Марка конструкции' не заполнен";
        }

        /// <summary>
        /// method that return a advance steel uid of the element
        /// </summary>
        public virtual Guid GetUid()
        {
            FilerObject filerObj = GetFilerObj(elem.Document, new Reference(elem));
            return filerObj.GetUniqueId();
        }

        /// <summary>
        /// method that return a advance steel FilerObject of the element. FilerObject is needed to use advance steel methods for revit elements
        /// </summary>
        public FilerObject GetFilerObj(RVTDocument doc, Reference eRef)
        {
            FilerObject filerObject = null;
            ASDocument curDocAS = DocumentManager.GetCurrentDocument();
            if (null != curDocAS)
            {
                OpenDatabase currentDatabase = curDocAS.CurrentDatabase;
                if (null != currentDatabase)
                {
                    Guid uid = SteelElementProperties.GetFabricationUniqueID(doc, eRef);
                    string asHandle = currentDatabase.getUidDictionary().GetHandle(uid);
                    filerObject = FilerObject.GetFilerObjectByHandle(asHandle);
                }
            }
            return filerObject;
        }

        /// <summary>
        /// method that return a material density
        /// </summary>
        public double GetMaterialDensity(ElementId materialId, Document doc)
        {
            Material material = doc.GetElement(materialId) as Material;
            PropertySetElement materialStructuralParams = doc.GetElement(material.StructuralAssetId) as PropertySetElement;
            double density = materialStructuralParams.get_Parameter(BuiltInParameter.PHY_MATERIAL_PARAM_STRUCTURAL_DENSITY).AsDouble();
            return density;
        }
    }
}
