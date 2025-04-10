using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Linq;
using System.Data;

namespace Plugin
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Sheets : IExternalCommand
    {
        //-********-//
        public static string IS_TAB_NAME => "01.Общее";
        public static string IS_NAME => "Заполнить штамп - КР";
        public static string IS_IMAGE => "Plugin.Resources.Sheet32.png";
        public static string IS_DESCRIPTION => "Перенос значений из параметров листов в параметры основных надписей для заполнения фамилий и даты";
        //-********-//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            var shets = new FilteredElementCollector(doc).
                OfCategory(BuiltInCategory.OST_Sheets).
                WhereElementIsNotElementType().
                Cast<Element>().
                ToList();

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Заполнение штампа");
                foreach (var el in shets)
                {
                    var titleBlockId = new FilteredElementCollector(doc, el.Id).OfCategory(BuiltInCategory.OST_TitleBlocks);
                    foreach (var block in titleBlockId)
                    {
                        var id = block.GetTypeId();
                        var elementType = (ElementType)doc.GetElement(id);
                        var familyName = elementType.FamilyName;
                        try
                        {
                            block.LookupParameter("ADSK_Штамп_1 фамилия").Set(el.LookupParameter("ADSK_Штамп_1 фамилия").AsString());
                            block.LookupParameter("ADSK_Штамп_2 фамилия").Set(el.LookupParameter("ADSK_Штамп_2 фамилия").AsString());
                            block.LookupParameter("ADSK_Штамп_3 фамилия").Set(el.LookupParameter("ADSK_Штамп_3 фамилия").AsString());
                            block.LookupParameter("ADSK_Штамп_4 фамилия").Set(el.LookupParameter("ADSK_Штамп_4 фамилия").AsString());
                            block.LookupParameter("ADSK_Штамп_5 фамилия").Set(el.LookupParameter("ADSK_Штамп_5 фамилия").AsString());
                            block.LookupParameter("ADSK_Штамп_6 фамилия").Set(el.LookupParameter("ADSK_Штамп_6 фамилия").AsString());
                            block.LookupParameter("ADSK_Штамп_1 должность").Set(el.LookupParameter("ADSK_Штамп_1 должность").AsString());
                            block.LookupParameter("ADSK_Штамп_2 должность").Set(el.LookupParameter("ADSK_Штамп_2 должность").AsString());
                            block.LookupParameter("ADSK_Штамп_3 должность").Set(el.LookupParameter("ADSK_Штамп_3 должность").AsString());
                            block.LookupParameter("ADSK_Штамп_4 должность").Set(el.LookupParameter("ADSK_Штамп_4 должность").AsString());
                            block.LookupParameter("ADSK_Штамп_5 должность").Set(el.LookupParameter("ADSK_Штамп_5 должность").AsString());
                            block.LookupParameter("ADSK_Штамп_6 должность").Set(el.LookupParameter("ADSK_Штамп_6 должность").AsString());
                            block.LookupParameter("Без подписи").Set((int)el.LookupParameter("Без подписи").AsInteger());
                            block.LookupParameter("Строка1_Дата").Set((int)el.LookupParameter("Строка1_Дата").AsInteger());
                            block.LookupParameter("Строка2_Дата").Set((int)el.LookupParameter("Строка2_Дата").AsInteger());
                            block.LookupParameter("Строка3_Дата").Set((int)el.LookupParameter("Строка3_Дата").AsInteger());
                            block.LookupParameter("Строка4_Дата").Set((int)el.LookupParameter("Строка4_Дата").AsInteger());
                            block.LookupParameter("Строка5_Дата").Set((int)el.LookupParameter("Строка5_Дата").AsInteger());
                            block.LookupParameter("Строка6_Дата").Set((int)el.LookupParameter("Строка6_Дата").AsInteger());
                        }
                        catch { }
                    }
                }
                tx.Commit();
            }
            return Result.Succeeded;
        }
    }
}

