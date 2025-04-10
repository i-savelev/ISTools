using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;


namespace Plugin
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class BrowserYn : IExternalCommand
    {
        //-***-//
        public static string IS_TAB_NAME => "01.Общее";
        public static string IS_NAME => "02.BIM-info";
        public static string IS_IMAGE => "Plugin.Resources.BrowserYn32.png";
        public static string IS_DESCRIPTION => "";
        //-***-//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            System.Diagnostics.Process.Start("https://53bim.yonote.ru/share/0d711288-7ece-45c3-8fad-7a48ed6d8ac4");
            return Result.Succeeded;
        }
    }
}

