using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;


namespace Plugin
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class BrowserYg : IExternalCommand
    {
        //-***-//
        public static string IS_TAB_NAME => "01.Общее";
        public static string IS_NAME => "01.YouGile";
        public static string IS_IMAGE => "Plugin.Resources.BrowserYg32.png";
        public static string IS_DESCRIPTION => "";
        //-***-//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            System.Diagnostics.Process.Start("https://ru.yougile.com/team/");
            return Result.Succeeded;
        }
    }
}

