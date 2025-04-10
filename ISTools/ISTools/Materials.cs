using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Linq;
using System.Data;

namespace ISTools
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Materials : IExternalCommand
    {
        //-********-//
        public static string IS_TAB_NAME => "ISTools";
        public static string IS_NAME => "Параметры материалов";
        public static string IS_DESCRIPTION => "Команда переносит значение из параметра материала 'Модель' в указанный параметр элемента\nВ случае, если элементу принадлежит несколько материалов, значения параметра будут разделены указанным разделителем";
        public static string IS_IMAGE => "ISTools.Resources.Materials32.png";
        //-********-//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            var elems = new FilteredElementCollector(doc).
                WhereElementIsNotElementType().
                Cast<Element>().
                ToList();

            DataTable dtMaterials = new DataTable();
            dtMaterials.Columns.Add("Название материала", typeof(string));
            dtMaterials.Columns.Add("Элемент", typeof(string));
            dtMaterials.Columns.Add("Материал 'Модель'", typeof(string));
            dtMaterials.Columns.Add("Id", typeof(string));

            foreach (var el in elems)
            {
                var typId = el.GetTypeId();
                var typ = el.Document.GetElement(typId);
                string model = "";
                string matName = "";
                if (el.GetType().ToString() == "Autodesk.Revit.DB.Structure.Rebar" || el.GetType().ToString() == "Autodesk.Revit.DB.Structure.RebarInSystem")
                {
                    var id1 = typ.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM).AsElementId();
                    var mat = el.Document.GetElement(id1) as Autodesk.Revit.DB.Material;
                    try
                    {
                        model = mat.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString();
                        matName = mat.Name;
                    }
                    catch { }
                    
                    if (matName != "")
                    {
                        dtMaterials.Rows.Add(
                            matName,
                            $"{typ.Category.Name}: {el.Name}",
                            model,
                            el.Id
                            );
                    }
                }
                else
                {
                    foreach (ElementId id in el.GetMaterialIds(false))
                    {
                        try
                        {
                            var mat = el.Document.GetElement(id) as Autodesk.Revit.DB.Material;
                            model = mat.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString();
                            matName = mat.Name;
                            var catName = "";
                            var elNmae = "";

                            try
                            {
                                catName = typ.Category.Name;
                                elNmae = el.Name;
                            }
                            catch 
                            {
                                catName = "Сталь";
                            }


                            if (matName != "")
                            {
                                try
                                {
                                    dtMaterials.Rows.Add(
                                        matName,
                                        $"{catName}: {elNmae}",
                                        model,
                                        el.Id
                                        );
                                }
                                catch
                                {
                                    TaskDialog.Show("a", el.Id.ToString());
                                }
                            }
                        } 
                        catch
                        {
                            //TaskDialog.Show("a", el.Id.ToString());
                        }
                    }
                }
            }

            ObjPlateCollector platesCollector = new ObjPlateCollector();
            platesCollector.doc = doc;
            //platesCollector.commandData = commandData;
            var platesInJoint = platesCollector.GetPlates();

            foreach ( var pij in platesInJoint )
            {
                string model = "";
                string matName = "";
                var catName = "Сталь";
                ParameterValue pv = pij.subelem.GetParameterValue(new ElementId(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM));
                ElementIdParameterValue idpv = pv as ElementIdParameterValue;
                ElementId mid = idpv.Value;
                Autodesk.Revit.DB.Material mat = doc.GetElement(mid) as Autodesk.Revit.DB.Material;
                try
                {
                    model = mat.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString();
                    matName = mat.Name;
                }
                catch { }

                if (matName != "")
                {
                    dtMaterials.Rows.Add(
                        matName,
                        $"{catName}: ",
                        model,
                        pij.subelem.GetReference().ElementId
                        );
                }
            }

            IsToolsForm window = new IsToolsForm();
            window.dataGridView1.DataSource = dtMaterials;
            window.Text = "Заполнение параметров материала";
            window.tabPage1.Text = "";
            window.groupBox1.Text = "";
            window.tabControl1.TabPages.RemoveByKey("tabPage2");
            window.tabControl1.TabPages.RemoveByKey("tabPage3");
            window.ExportButton.Text = "Заполнить параметр";
            window.toolStripTextBox2.Text = "<Название параметра>";
            window.textBox3.Text = "В таблице ниже содержится список элементов и назначенных им материалов, которые размещены в модели. Для заполнения параметра элементам модели необходимо указать название параметра и нажать кнопку 'Заполнить параметр'. Значение параметра материала 'Модель' будут записаны в указанный параметр элементов";
            window.textBox3.Height = 80;
            window.ExportButton.Click += (s, e) => { FillParam(); };
            window.toolStripTextBox4.Text = "<Разделитель>";

            window.ShowDialog();
            void FillParam()
            {
                window.toolStripProgressBar1.Value = 0;
                window.toolStripProgressBar1.Maximum = elems.Count + platesInJoint.Count;
                window.toolStripProgressBar1.Step = 1;
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Заполнение параметров материала");
                    foreach (var el in elems)
                    {
                        window.toolStripProgressBar1.PerformStep();
                        var typId = el.GetTypeId();
                        var typ = el.Document.GetElement(typId);
                        string model = "";
                        string param = "";
                        if (el.GetType().ToString() == "Autodesk.Revit.DB.Structure.Rebar" || el.GetType().ToString() == "Autodesk.Revit.DB.Structure.RebarInSystem")
                        {
                            var id1 = typ.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM).AsElementId();
                            var mat = el.Document.GetElement(id1) as Autodesk.Revit.DB.Material;
                            model = mat.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString();
                            el.LookupParameter(window.toolStripTextBox2.Text).Set(model);
                            window.toolStripProgressBar1.PerformStep();
                        }
                        else
                        {
                            foreach (ElementId id in el.GetMaterialIds(false))
                            {
                                var mat = el.Document.GetElement(id) as Autodesk.Revit.DB.Material;
                                model = mat.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString();
                                param = param + window.toolStripTextBox4.Text + model;
                                try
                                {
                                    el.LookupParameter(window.toolStripTextBox2.Text).Set(param.TrimStart(window.toolStripTextBox4.Text.ToCharArray()));
                                }
                                catch { }
                                
                            }
                        }
                    }
                    foreach (var pij in platesInJoint)
                    {
                        window.toolStripProgressBar1.PerformStep();
                        string model = "";
                        ParameterValue pv = pij.subelem.GetParameterValue(new ElementId(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM));
                        ElementIdParameterValue idpv = pv as ElementIdParameterValue;
                        ElementId mid = idpv.Value;
                        Autodesk.Revit.DB.Material mat = doc.GetElement(mid) as Autodesk.Revit.DB.Material;
                        try
                        {
                            model = mat.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString();
                            pij.SetParam(window.toolStripTextBox2.Text, model);
                        }
                        catch { }
                    }
                    tx.Commit();
                    window.toolStripProgressBar1.Value = elems.Count + platesInJoint.Count;
                }
            }
            return Result.Succeeded;
        }
    }
}
