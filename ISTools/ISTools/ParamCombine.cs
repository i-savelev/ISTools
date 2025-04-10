using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Windows.Forms;


namespace Plugin
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ParamCombine : IExternalCommand
    {
        //---PluginsManager---//
        public static string IS_TAB_NAME => "#params";
        public static string IS_NAME => "Комбинация параметров";
        public static string IS_DESCRIPTION => "Команда комбинирует указанные параметры";
        public static string IS_IMAGE => "Plugin.Resources.Combine32.png";
        //---PluginsManager---//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            var sParams = new FilteredElementCollector(doc).
                OfClass(typeof(SharedParameterElement)).
                WhereElementIsNotElementType().
                Cast<Element>().
                ToList();

            AutoCompleteStringCollection col = new AutoCompleteStringCollection();
            Form1 window = new Form1();
            foreach (var param in sParams)
            {
                window.comboBox1.Items.Add(param.Name);
                col.Add(param.Name);
            }
            window.comboBox1.AutoCompleteCustomSource = col;
            window.Text = "Комбинирование параметров";
            window.discriptionTextBox1.Text = "Наименование параметра для записи";
            window.discriptionTextBox2.Text = "Для комбинации параметров необходимо указать имена параметров в фигурных скобках: {Параметр}. Между параметрами можно расположить разделители";
            window.inputTextBox1.Text = "<Com>{Параметр1}.{Параметр2}%{Параметр3}/++{Параметр4}";
            window.button1.Click += (s, e) => { ParamCombine(); };
            window.button1.Text = "Записать значения";
            window.ShowDialog();

            void ParamCombine()
            {
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Комбинация параметров");
                    var input = window.inputTextBox1.Text;
                    List<string> listParam = new List<string>();
                    List<string> separatorList = new List<string>();
                    string[] pa = input.Split('{');
                    for (int i = 0; i < pa.Length; i++)
                    {
                        string str = pa[i];
                        if (!str.Contains("}")) continue;
                        string paramName = str.Split('}').First();
                        listParam.Add(paramName);  
                    }

                    string[] se = input.Split('}');
                    for (int i = 0; i < se.Length; i++)
                    {
                        string separator = "";
                        string s = se[i];
                        separator = s.Split('{').First();
                        separatorList.Add(separator);
                    }

                    var elems = new FilteredElementCollector(doc).
                        WhereElementIsNotElementType().
                        Cast<Element>().
                        ToList();
                    ObjPlateCollector platesCollector = new ObjPlateCollector();
                    platesCollector.doc = doc;
                    var platesInJoint = platesCollector.GetPlates();

                    window.progressBar1.Value = 0;
                    window.progressBar1.Maximum = elems.Count + platesInJoint.Count;
                    window.progressBar1.Step = 1;

                    foreach (var elem in elems)
                    {
                        window.progressBar1.PerformStep();
                        ObjRvt obj = new ObjRvt();
                        obj.elem = elem;
                        List<string> elemParamList = new List<string>();
                        try
                        {
                            for (int i = 0; i < separatorList.Count - 1; i++)
                            {
                                elemParamList.Add(separatorList[i] + obj.GetParam(listParam[i]));
                            }
                            elemParamList.Add(separatorList.Last());
                            obj.SetParam(window.comboBox1.Text, string.Join("", elemParamList));
                        }
                        catch { }
                    }

                    foreach(var pij in platesInJoint)
                    {
                        window.progressBar1.PerformStep();
                        List<string> elemParamList = new List<string>();
                        try
                        {
                            for (int i = 0; i < separatorList.Count - 1; i++)
                            {
                                elemParamList.Add(separatorList[i] + pij.GetParam(listParam[i]));
                            }
                            elemParamList.Add(separatorList.Last());
                            pij.SetParam(window.comboBox1.Text, string.Join("", elemParamList));
                        }
                        catch { }
                    }
                    tx.Commit();
                    window.progressBar1.Value = elems.Count + platesInJoint.Count;
                }
            }
            return Result.Succeeded;
        }
    }
}
