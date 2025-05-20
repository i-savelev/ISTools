using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;



namespace ISTools
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class SetFilters : IExternalCommand
    {
        //---PluginsManager---//
        public static string IS_TAB_NAME => "ISTools";
        public static string IS_NAME => "Фильтры на листах";
        public static string IS_IMAGE => "ISTools.Resources.SheetsCopy32.png";
        public static string IS_DESCRIPTION => "Позволяет массово перенести фильтры на выбранные виды и шаблоны. Если к виду или шаблону уже применен такой фильтр, будет перезаписан";
        //---PluginsManager---//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            try
            {
                SetFiltersModel viewModel = new SetFiltersModel(doc);
                FilterTransferWindow window = new FilterTransferWindow()
                {
                    DataContext = viewModel
                };

                window.ShowDialog(); // Открываем как диалоговое окно

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}

