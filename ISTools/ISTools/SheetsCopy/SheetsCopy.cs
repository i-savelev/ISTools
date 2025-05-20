using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using ISTools.ISTools.SheetsCopy;


namespace ISTools
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class SheetsCopy : IExternalCommand
    {
        //---PluginsManager---//
        public static string IS_TAB_NAME => "ISTools";
        public static string IS_NAME => "Копирование листов";
        public static string IS_IMAGE => "ISTools.Resources.SheetsCopy32.png";
        public static string IS_DESCRIPTION => "Плагин позволяет копировать листы вместе с видами, аннатциями, текстом и параметрами";
        //---PluginsManager---//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Получаем доступ к текущему документу
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            SheetManager sheetManager = new SheetManager(doc);
            SheetsCopyWindowManager sheetsCopyWindowManager = new SheetsCopyWindowManager(sheetManager);
            //IsDebugWindow.Show();
            return Result.Succeeded;
        }
    }
}

