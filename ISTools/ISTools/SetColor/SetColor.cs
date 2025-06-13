using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ISTools.MyObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;

namespace ISTools
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    internal class SetColor : IExternalCommand
    {
        //---PluginsManager---//
        public static string IS_TAB_NAME => "ISTools";
        public static string IS_NAME => "Раскраска элементов";
        public static string IS_IMAGE => "ISTools.Resources.Set_color_32.png";
        public static string IS_DESCRIPTION => "Плагин позволяет раскрасить элементы в цвета по выбранному параметру\nПодробнее: https://github.com/i-savelev/ISTools/wiki/Раскраска-элементов";
        //---PluginsManager---//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            //declaring variables
            var categories_list = IsUtils.GetCategoryModelList(doc);
            SetColorForm window = new SetColorForm();
            ObjColorsData objColorsData = new ObjColorsData();
            List<ObjRvt> objRvtList = new List<ObjRvt>();
            string xmlPath = $@"c:\ProgramData\SetColorTemp.xml";

            //datagrid categories
            DataGridViewComboBoxColumn cmb = new DataGridViewComboBoxColumn();
            cmb.HeaderText = "Категория";
            cmb.DataSource = categories_list;
            window.dataGridView2.Columns.Add(cmb);
            window.button_add.Click += (s, e) => { IsUtils.AddRow(window.dataGridView2); };
            window.button_remove.Click += (s, e) => { IsUtils.DeleteRow(window.dataGridView2); };
            window.toolStripButton_save_settings.Click += (s, e) => { SaveSettings(); };
           
            // datagrid colors
            window.dataGridView1.ColumnCount = 2;
            window.dataGridView1.Columns[0].Name = "Цвет";
            window.dataGridView1.Columns[0].Width = 70;
            window.dataGridView1.Columns[1].Name = "Значение параметра";
            window.dataGridView1.Columns[1].Width = 300;
            window.toolStripButton_change_color.Click += (s, e) => { IsUtils.ChengeDatagridCellsColor(window.dataGridView1, "Цвет"); };
            window.toolStripButton_accept_change.Click += (s, e) => { SaveColorCnange(window.dataGridView1); };

            window.button_save.Click += (s, e) => { IsUtils.SaveXmlDialog<ObjColorsData>(objColorsData); };
            window.button_set_color.Click += (s, e) => {SetColorToElements(); };
            window.button_load.Click += (s, e) => { LoadXml(); };
            window.button_remove_color.Click += (s, e) => { ClearOverrideFromElements(); };
            window.textBox4.Text = "Для раскраски элементов необходимо выбрать категории и нажать \"Сохранить настройки\", казать параметр и нажать кнопку \"Раскрасить\". Для выбранных значений параметра можно переопределить цвет, для сохранения переопределения необходимо нажать \"Применить настройки\". После переопределения цвета можно раскрасить по новой. Примененные настройки можно сохранить и загрузить";

            try
            {
                window.FormClosed += (s, e) => { IsUtils.SaveXml<ObjColorsData>(objColorsData, xmlPath); };
            }
            catch { }

            try
            {
                LoadXmlWhenPpen();
            }
            catch { }

            window.ShowDialog();

            void SaveSettings()
            {
                var allCategories = doc.Settings.Categories;
                List<ObjCategory> categoryList = new List<ObjCategory>();
                for (int i = 0; i < window.dataGridView2.Rows.Count; i++)
                {
                    string categoryName = $"{window.dataGridView2.Rows[i].Cells[0].Value}";
                    ObjCategory objCat = new ObjCategory(categoryName);
                    foreach (Category cat in allCategories)
                    {
                        if (categoryName == cat.Name)
                        {
                            objCat.BuiltIn = (BuiltInCategory)cat.Id.IntegerValue;
                        }
                        if (categoryName == "Выступающие профили")
                        {
                            objCat.BuiltIn = BuiltInCategory.OST_Cornices;
                        }
                    }
                    categoryList.Add(objCat);
                }
                objColorsData.CategoryList = categoryList;
                objColorsData.ParameterName = window.textBox2.Text;
            }

            void SetColorToElements()
            {
                BuiltInCategory[] builtinCategoriesFilter = objColorsData.CategoryList
                    .Select(cat => cat.BuiltIn)
                    .ToArray();

                var elementList = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .WherePasses(new ElementMulticategoryFilter(builtinCategoriesFilter))
                    .WhereElementIsNotElementType()
                    .Cast<Element>()
                    .ToList();

                objRvtList.Clear();

                foreach (Element element in elementList)
                {
                    ObjRvt objRvt = new ObjRvt();
                    objRvt.elem = element;
                    objRvtList.Add(objRvt);
                }

                var paramName = objColorsData.ParameterName;
                var uniqueParamValueList = objRvtList.Select(obj => obj.GetParamAsString(paramName))
                    .Distinct()
                    .ToList();

                uniqueParamValueList.Sort();
                uniqueParamValueList.RemoveAll(value => value == "");
                uniqueParamValueList.RemoveAll(value => value == null);
                CreateColorList(uniqueParamValueList);
                SetColorsToDatagrid(uniqueParamValueList);
                ColorChangeTansaction(objRvtList);
            }

            void ColorChangeTansaction(List<ObjRvt> objList)
            {
                window.progressBar1.Value = 0;
                window.progressBar1.Maximum = objList.Count;
                window.progressBar1.Step = 1;
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                var paramName = objColorsData.ParameterName;
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Раскраска элементов");
                    foreach (ObjRvt objRvt in objList)
                    {
                        try
                        {
                            window.progressBar1.PerformStep();
                            var objColor = objColorsData.ColorList.
                                FirstOrDefault(p => p.DiscriptionKey == objRvt.GetParamAsString(paramName));
                            var rvtColor = IsUtils.ConvSysToRvtColor(objColor.ColorFromList());
                            objRvt.Set_element_color(rvtColor);
                        }
                        catch { }
                        
                    }
                    tx.Commit();
                }
                window.progressBar1.Value = objList.Count;
                stopwatch.Stop();
                TimeSpan elapsedTime = stopwatch.Elapsed;
                window.textBox1.Text = $"Время выполнения: {Math.Round(elapsedTime.TotalSeconds, 3).ToString()} сек.";
            }

            void CreateColorList(List<string> uniqueParamValueList)
            {
                Random rnd = new Random();
                foreach (var paramValue in uniqueParamValueList)
                {
                    if (!objColorsData.ColorList.Any(p => p.DiscriptionKey == paramValue) & paramValue != "")
                    {
                        ObjColorByKey objColorByKey = new ObjColorByKey();
                        objColorByKey.DiscriptionKey = paramValue;
                        var color = IsUtils.GetRandomSystemColor(rnd);
                        objColorByKey.Color_int_list = IsUtils.GetColorListInt(color);
                        objColorsData.ColorList.Add(objColorByKey);
                    }
                }
                objColorsData.ColorList.RemoveAll(color => color.DiscriptionKey == "" );
                objColorsData.ColorList.RemoveAll(color => color.DiscriptionKey == null);
            }

            void SetColorsToDatagrid(List<string> uniqueParamValueList)
            {
                int i = 0;
                window.dataGridView1.Rows.Clear();
                window.dataGridView1.RowCount = uniqueParamValueList.Count;
                foreach (var systemColor in objColorsData.ColorList)
                {
                    if ( uniqueParamValueList.Contains(systemColor.DiscriptionKey) & systemColor.DiscriptionKey != "")
                    {
                        var color = systemColor.ColorFromList();
                        window.dataGridView1.Rows[i].Cells[0].Style.BackColor = color;
                        window.dataGridView1.Rows[i].Cells[1].Value = systemColor.DiscriptionKey;
                        i += 1;
                    }
                }
            }

            void LoadXml()
            {
                var objTemp = IsUtils.LoadXmlDialog<ObjColorsData>();
                if (objTemp != null)
                {
                    objColorsData = objTemp;
                }
                int i = 0;
                window.dataGridView2.Rows.Clear();
                window.dataGridView2.RowCount = objColorsData.CategoryList.Count;
                foreach (var cat in objColorsData.CategoryList)
                {
                    window.dataGridView2.Rows[i].Cells[0].Value = cat.Name;
                    i += 1;
                }
                window.textBox2.Text = objColorsData.ParameterName;
            }

            void LoadXmlWhenPpen()
            {
                var objTemp = IsUtils.LoadXml<ObjColorsData>(xmlPath);
                if (objTemp != null)
                {
                    objColorsData = objTemp;
                }
                int i = 0;
                window.dataGridView2.Rows.Clear();
                window.dataGridView2.RowCount = objColorsData.CategoryList.Count;
                foreach (var cat in objColorsData.CategoryList)
                {
                    window.dataGridView2.Rows[i].Cells[0].Value = cat.Name;
                    i += 1;
                }
                window.textBox2.Text = objColorsData.ParameterName;
            }

            void ClearOverrideFromElements()
            {
                var elementList = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .Cast<Element>()
                    .ToList();

                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Удаление раскраски");
                    foreach (Element elem in elementList)
                    {
                        ObjRvt rvt_obj = new ObjRvt();
                        rvt_obj.elem = elem;
                        rvt_obj.Clear_overrides();
                    }
                    tx.Commit();
                }
            }

            void SaveColorCnange(DataGridView dataGridView)
            {
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    string paramValue = $"{dataGridView.Rows[i].Cells[1].Value}";
                    System.Drawing.Color system_color = dataGridView.Rows[i].Cells[0].Style.BackColor;
                    ObjColorByKey obj_Color_By_Key = new ObjColorByKey();
                    objColorsData.ColorList.FirstOrDefault(color => color.DiscriptionKey == paramValue)
                        .Color_int_list = IsUtils.GetColorListInt(system_color);
                }
            }
            return Result.Succeeded;
        }
    }
}
