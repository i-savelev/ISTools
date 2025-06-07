using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;


namespace ISTools
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class SchedulesTable : IExternalCommand
    {
        //---PluginsManager---//
        public static string IS_TAB_NAME => "#|#test|#bim";
        public static string IS_NAME => "Ведомость спецификаций";
        public static string IS_IMAGE => "";
        public static string IS_DESCRIPTION => "Тестирвоание плагина для ведомости спецификаций";
        //---PluginsManager---//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            
            SchedulesTableModel viewModel = new SchedulesTableModel(doc);
            SchedulesTableForm window = new SchedulesTableForm()
            {
                DataContext = viewModel
            };

                // Открываем как диалоговое окно
                
            window.Show();
            //IsDebugWindow.AddRow($"Количество элементов в дереве: {viewModel.TreeViewItemViewModelList.Count}");
            //IsDebugWindow.Show();
            return Result.Succeeded;
            
            
        }
    }
}
