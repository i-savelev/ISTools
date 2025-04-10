using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Plugin.MyObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;

namespace Plugin
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    internal class SetColor : IExternalCommand
    {
        //-********-//
        public static string IS_TAB_NAME => "01.Общее";
        public static string IS_NAME => "Раскраска элементов";
        public static string IS_IMAGE => "Plugin.Resources.Set_color_32.png";
        public static string IS_DESCRIPTION => "Плагин позволяет раскрасить элементы в цвета по выбранному параметру";
        //-********-//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            //declaring variables
            var categories_list = IsUtils.Get_category_model_list(doc);
            SetColorForm window = new SetColorForm();
            Obj_Colors_Data obj_Colors_Data = new Obj_Colors_Data();
            List<ObjRvt> obj_rvt_list = new List<ObjRvt>();
            string xml_path = $@"c:\ProgramData\SetColorTemp.xml";

            //datagrid categories
            DataGridViewComboBoxColumn cmb = new DataGridViewComboBoxColumn();
            cmb.HeaderText = "Категория";
            cmb.DataSource = categories_list;
            window.dataGridView2.Columns.Add(cmb);
            window.button_add.Click += (s, e) => { IsUtils.Add_row(window.dataGridView2); };
            window.button_remove.Click += (s, e) => { IsUtils.Delete_row(window.dataGridView2); };
            window.toolStripButton_save_settings.Click += (s, e) => { Save_settings(); };
           
            // datagrid colors
            window.dataGridView1.ColumnCount = 2;
            window.dataGridView1.Columns[0].Name = "Цвет";
            window.dataGridView1.Columns[0].Width = 70;
            window.dataGridView1.Columns[1].Name = "Значение параметра";
            window.dataGridView1.Columns[1].Width = 300;
            window.toolStripButton_change_color.Click += (s, e) => { IsUtils.Chenge_datagrid_cells_color(window.dataGridView1, "Цвет"); };
            window.toolStripButton_accept_change.Click += (s, e) => { Save_color_cnange(window.dataGridView1); };

            window.button_save.Click += (s, e) => { IsUtils.Save_xml_dialog<Obj_Colors_Data>(obj_Colors_Data); };
            window.button_set_color.Click += (s, e) => {Set_color_to_elements(); };
            window.button_load.Click += (s, e) => { Load_xml(); };
            window.button_remove_color.Click += (s, e) => { Clear_override_from_elements(); };
            window.textBox4.Text = "Для раскраски элементов необходимо выбрать категории и нажать \"Сохранить настройки\", казать параметр и нажать кнопку \"Раскрасить\". Для выбранных значений параметра можно переопределить цвет, для сохранения переопределения необходимо нажать \"Применить настройки\". После переопределения цвета можно раскрасить по новой. Примененные настройки можно сохранить и загрузить";

            try
            {
                window.FormClosed += (s, e) => { IsUtils.Save_xml<Obj_Colors_Data>(obj_Colors_Data, xml_path); };
        }
            catch { }

            try
            {
                Load_xml_when_open();
            }
            catch { }

            window.ShowDialog();

            void Save_settings()
            {
                var all_categories = doc.Settings.Categories;
                List<Obj_Category> category_list = new List<Obj_Category>();
                for (int i = 0; i < window.dataGridView2.Rows.Count; i++)
                {
                    string categoryName = $"{window.dataGridView2.Rows[i].Cells[0].Value}";
                    Obj_Category objCat = new Obj_Category(categoryName);
                    foreach (Category cat in all_categories)
                    {
                        if (categoryName == cat.Name)
                        {
                            objCat.BuiltIn = cat.BuiltInCategory;
                        }
                        if (categoryName == "Выступающие профили")
                        {
                            objCat.BuiltIn = BuiltInCategory.OST_Cornices;
                        }
                    }
                    category_list.Add(objCat);
                }
                obj_Colors_Data.Category_list = category_list;
                obj_Colors_Data.Parameter_name = window.textBox2.Text;
            }

            void Set_color_to_elements()
            {
                BuiltInCategory[] builtin_categories_filter = obj_Colors_Data.Category_list
                    .Select(cat => cat.BuiltIn)
                    .ToArray();

                var element_list = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .WherePasses(new ElementMulticategoryFilter(builtin_categories_filter))
                    .WhereElementIsNotElementType()
                    .Cast<Element>()
                    .ToList();

                obj_rvt_list.Clear();

                foreach (Element element in element_list)
                {
                    ObjRvt obj_Rvt = new ObjRvt();
                    obj_Rvt.elem = element;
                    obj_rvt_list.Add(obj_Rvt);
                }

                var param_name = obj_Colors_Data.Parameter_name;
                var unique_param_value_list = obj_rvt_list.Select(obj => obj.GetParamAsString(param_name))
                    .Distinct()
                    .ToList();

                unique_param_value_list.Sort();
                unique_param_value_list.RemoveAll(value => value == "");
                unique_param_value_list.RemoveAll(value => value == null);
                Create_color_list(unique_param_value_list);
                Set_colors_to_datagrid(unique_param_value_list);
                Color_change_tansaction(obj_rvt_list);
            }

            void Color_change_tansaction(List<ObjRvt> obj_list)
            {
                window.progressBar1.Value = 0;
                window.progressBar1.Maximum = obj_list.Count;
                window.progressBar1.Step = 1;
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                var param_name = obj_Colors_Data.Parameter_name;
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Раскраска элементов");
                    foreach (ObjRvt obj_Rvt in obj_list)
                    {
                        try
                        {
                            window.progressBar1.PerformStep();

                            var obj_color = obj_Colors_Data.Color_list.
                                FirstOrDefault(p => p.Discription_key == obj_Rvt.GetParamAsString(param_name));
                            var rvt_color = IsUtils.Conv_sys_to_rvt_color(obj_color.Color_from_list());
                            obj_Rvt.Set_element_color(rvt_color);
                        }
                        catch { }
                        
                    }
                    tx.Commit();
                }
                window.progressBar1.Value = obj_list.Count;
                stopwatch.Stop();
                TimeSpan elapsedTime = stopwatch.Elapsed;
                window.textBox1.Text = $"Время выполнения: {Math.Round(elapsedTime.TotalSeconds, 3).ToString()} сек.";
            }

            void Create_color_list(List<string> unique_param_value_list)
            {
                Random rnd = new Random();
                foreach (var param_value in unique_param_value_list)
                {
                    if (!obj_Colors_Data.Color_list.Any(p => p.Discription_key == param_value) & param_value != "")
                    {
                        Obj_color_by_key obj_Color_By_Key = new Obj_color_by_key();
                        obj_Color_By_Key.Discription_key = param_value;
                        var color = IsUtils.Get_random_system_color(rnd);
                        obj_Color_By_Key.Color_int_list = IsUtils.Get_color_list_int(color);
                        obj_Colors_Data.Color_list.Add(obj_Color_By_Key);
                    }
                }
                obj_Colors_Data.Color_list.RemoveAll(color => color.Discription_key == "" );
                obj_Colors_Data.Color_list.RemoveAll(color => color.Discription_key == null);
            }

            void Set_colors_to_datagrid(List<string> unique_param_value_list)
            {
                int i = 0;
                window.dataGridView1.Rows.Clear();
                window.dataGridView1.RowCount = unique_param_value_list.Count;
                foreach (var system_color in obj_Colors_Data.Color_list)
                {
                    if ( unique_param_value_list.Contains(system_color.Discription_key) & system_color.Discription_key != "")
                    {
                        var color = system_color.Color_from_list();
                        window.dataGridView1.Rows[i].Cells[0].Style.BackColor = color;
                        window.dataGridView1.Rows[i].Cells[1].Value = system_color.Discription_key;
                        i += 1;
                    }
                }
            }

            void Load_xml()
            {
                var obj_temp = IsUtils.Load_xml_dialog<Obj_Colors_Data>();
                if (obj_temp != null)
                {
                    obj_Colors_Data = obj_temp;
                }
                int i = 0;
                window.dataGridView2.Rows.Clear();
                window.dataGridView2.RowCount = obj_Colors_Data.Category_list.Count;
                foreach (var cat in obj_Colors_Data.Category_list)
                {
                    window.dataGridView2.Rows[i].Cells[0].Value = cat.Name;
                    i += 1;
                }
                window.textBox2.Text = obj_Colors_Data.Parameter_name;
            }

            void Load_xml_when_open()
            {
                var obj_temp = IsUtils.Load_xml<Obj_Colors_Data>(xml_path);
                if (obj_temp != null)
                {
                    obj_Colors_Data = obj_temp;
                }
                int i = 0;
                window.dataGridView2.Rows.Clear();
                window.dataGridView2.RowCount = obj_Colors_Data.Category_list.Count;
                foreach (var cat in obj_Colors_Data.Category_list)
                {
                    window.dataGridView2.Rows[i].Cells[0].Value = cat.Name;
                    i += 1;
                }
                window.textBox2.Text = obj_Colors_Data.Parameter_name;
            }

            void Clear_override_from_elements()
            {
                var element_list = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .Cast<Element>()
                    .ToList();

                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Удаление раскраски");
                    foreach (Element elem in element_list)
                    {
                        ObjRvt rvt_obj = new ObjRvt();
                        rvt_obj.elem = elem;
                        rvt_obj.Clear_overrides();
                    }
                    tx.Commit();
                }
            }

            void Save_color_cnange(DataGridView dataGridView)
            {
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    string param_value = $"{dataGridView.Rows[i].Cells[1].Value}";
                    System.Drawing.Color system_color = dataGridView.Rows[i].Cells[0].Style.BackColor;
                    Obj_color_by_key obj_Color_By_Key = new Obj_color_by_key();
                    obj_Colors_Data.Color_list.FirstOrDefault(color => color.Discription_key == param_value)
                        .Color_int_list = IsUtils.Get_color_list_int(system_color);
                }
            }
            return Result.Succeeded;
        }
    }
}
