using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using Document = Autodesk.Revit.DB.Document;
using System.Linq;
using System.Collections.Generic;
using Plugin.Forms;
using System.Data;
using System.Windows.Forms;
using Plugin.MyObjects;
using System.Xml.Serialization;
using System.IO;
using System.Diagnostics;


namespace Plugin
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class SetWorksets : IExternalCommand
    {
        //---PluginsManager---//
        public static string IS_TAB_NAME => "01.Общее";
        public static string IS_NAME => "Рабочие наборы";
        public static string IS_IMAGE => "Plugin.Resources.Worksets32.png";
        public static string IS_DESCRIPTION => "Скрипт автоматически присваевает элементам рабочие наборы по заданным правилам";
        //---PluginsManager---//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<Obj_Workset>worksets_List = new List<Obj_Workset>();
            List<string> categories_list = new List<string>();
            List<string> parameter_condition_list = new List<string>() {"равно", "содержит", "не содержит", "не равно" };
            List<string> parameter_bool_condition_list = new List<string>() { "и", "или" };

            DataTable dt_output = new DataTable();
            dt_output.Columns.Add("1");

            var all_categories = doc.Settings.Categories;
            foreach (Category category in all_categories)
            {
                dt_output.Rows.Add($"{category.Name} - {(BuiltInCategory)category.Id.IntegerValue}"); 
                if (category.CategoryType == CategoryType.Model)
                {
                    categories_list.Add(category.Name);
                   
                }
            }

            categories_list.Add("Оси");
            categories_list.Add("Уровни");
            categories_list.Add("Выступающие профили");
            categories_list.Sort();

            SetWorksetsForm window = new SetWorksetsForm();
            window.Text = "Настройка рабочих наборов";
            //datagrid categories
            window.groupBox1.Text = "Категории";
            DataGridViewComboBoxColumn cmb = new DataGridViewComboBoxColumn();
            cmb.HeaderText = "Категория";
            cmb.DataSource = categories_list;
            window.dataGridView1.Columns.Add(cmb);
            window.button7.Text = "+";
            window.button8.Text = "-";
            window.button7.Click += (s, e) => { IsUtils.Add_row(window.dataGridView1); };
            window.button8.Click += (s, e) => { IsUtils.Delete_row(window.dataGridView1); };

            //datagrid parameters
            window.groupBox2.Text = "Параметры";
            DataGridViewColumn columnParameterName = new DataGridViewColumn();
            columnParameterName.HeaderText = "Название параметра";
            columnParameterName.CellTemplate = new DataGridViewTextBoxCell();
            DataGridViewComboBoxColumn columnParameterCondition = new DataGridViewComboBoxColumn();
            columnParameterCondition.DataSource = parameter_condition_list;
            columnParameterCondition.HeaderText = "Условние";
            DataGridViewComboBoxColumn columnParameterBoolCondition = new DataGridViewComboBoxColumn();
            columnParameterBoolCondition.DataSource = parameter_bool_condition_list;
            columnParameterBoolCondition.HeaderText = "и/или";
            DataGridViewColumn columnParameterValue = new DataGridViewColumn();
            columnParameterValue.HeaderText = "Значение параметра";
            columnParameterValue.CellTemplate = new DataGridViewTextBoxCell();
            window.dataGridView2.Columns.Add(columnParameterBoolCondition);
            window.dataGridView2.Columns.Add(columnParameterName);
            window.dataGridView2.Columns.Add(columnParameterCondition);
            window.dataGridView2.Columns.Add(columnParameterValue);
            window.dataGridView2.Columns[0].Width = 60;
            window.dataGridView2.Columns[2].Width = 120;
            window.button4.Text = "-";
            window.button2.Text = "+";
            window.button4.Click += (s, e) => { IsUtils.Delete_row(window.dataGridView2); };
            window.button2.Click += (s, e) => { IsUtils.Add_row(window.dataGridView2); };

            //datagrid worksets
            window.groupBox4.Text = "Рабчие наборы";
            DataGridViewColumn columnWorksetName = new DataGridViewColumn();
            columnWorksetName.HeaderText = "Название рабочего набора";
            columnWorksetName.CellTemplate = new DataGridViewTextBoxCell();
            window.dataGridView3.Columns.Add(columnWorksetName);
            window.button5.Click += (s, e) => { DeleteWorkSet(window.dataGridView3); };
            window.button1.Click += (s, e) => { AddWorksetDialog(); };
            window.button1.Text = "Добавить";
            window.button5.Text = "Удалить";
            window.dataGridView3.CellClick += dataGridViewWorkset_CellClick;
            window.button10.Click += (s, e) => { RenameWorkSetDialog(); };
            window.button10.Text = "Переименовать";
            window.button12.Text = "Копировать";
            window.button12.Click += (s, e) => { Copy_workset(window.dataGridView3); }; 

            window.comboBox1.DataSource = parameter_bool_condition_list;
            window.groupBox3.Text = "Настройки";
            window.button3.Text = "Сохранить рабочий набор";
            window.button3.Click += (s, e) => { SaveWorkset(); };
            window.button6.Click += (s, e) => { SaveXml(); };
            window.button6.Text = "Сохранить шаблон";
            window.button9.Text = "Загрузить шаблон"; 
            window.button9.Click += (s, e) => { LoadXml(); };
            window.button11.Text = "Распределить элементы по наборам";
            window.button11.Click += (s, e) => { SetWsToElement(); };
            try
            {
                window.FormClosed += (s, e) => { Saving_when_closing(); };
            }
            catch { }

            try
            {
                Load_xml_when_open();
            }
            catch { }

            window.groupBox5.Text = "Информация";
            window.textBox2.Text = "Данный скрипт позволяет автоматически распределять элементы по рабочим наборам. Готовая конфигурация для разделов КР: w:\\00. Ресурсы\\03.Плагины\\05.КР\\53ЦПИ-СПБ-КР\\Рабочие наборы";

            window.ShowDialog();

            void dataGridViewWorkset_CellClick(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0) return;
                window.dataGridView2.Rows.Clear();
                window.dataGridView1.Rows.Clear();
                var rowindex = e.RowIndex;
                foreach (var objWs in worksets_List)
                {
                    if ((string)window.dataGridView3[0, rowindex].Value == objWs.Name)
                    {
                        if (objWs.Conditions != null)
                        {
                            foreach (var cond in objWs.Conditions)
                            {
                                window.dataGridView2.Rows.Add(cond.GetStrings());
                            }
                        }

                        if (objWs.Categories != null)
                        {
                            foreach (var cat in objWs.Categories)
                            {
                                window.dataGridView1.Rows.Add(cat.GetStrings());
                            }
                        }

                        if (objWs.Categories_and_Params != null)
                        {
                            window.comboBox1.Text = objWs.Categories_and_Params;
                        }
                    }
                }
            }

            void RenameWorkset(InputDialogForm inputWindow)
            {
                string name = "";
                List<Obj_Param_Condition> conditions = new List<Obj_Param_Condition>();
                foreach (DataGridViewRow row in window.dataGridView3.SelectedRows)
                {
                    var rowindex = row.Index;
                    name = $"{window.dataGridView3.Rows[rowindex].Cells[0].Value}";
                }

                foreach (var ws in worksets_List)
                {
                    if (ws.Name == name)
                    {
                        ws.Name = inputWindow.textBox2.Text;
                        foreach (DataGridViewRow row in window.dataGridView3.SelectedRows)
                        {
                            var rowindex = row.Index;
                            window.dataGridView3.Rows[rowindex].Cells[0].Value = inputWindow.textBox2.Text;
                        }
                        inputWindow.Close();
                    }
                }
            }

            void RenameWorkSetDialog()
            {
                string name = "";
                foreach (DataGridViewRow row in window.dataGridView3.SelectedRows)
                {
                    var rowindex = row.Index;
                    name = $"{window.dataGridView3.Rows[rowindex].Cells[0].Value}";
                }

                InputDialogForm inputWindow = new InputDialogForm();
                inputWindow.textBox1.Text = "Укажите название рабочего набора";
                inputWindow.Text = "Переименование";
                inputWindow.textBox2.Text = name;
                inputWindow.button1.Text = "Ок";
                inputWindow.button1.Click += (s, e) => { RenameWorkset(inputWindow); };
                inputWindow.ShowDialog();
            }

            void SaveWorkset()
            {
                string name = "";
                List<Obj_Param_Condition> conditions_list = new List<Obj_Param_Condition>();
                List<Obj_Category> category_list = new List<Obj_Category>();
                foreach (DataGridViewRow row in window.dataGridView3.SelectedRows)
                {
                    var rowindex = row.Index;
                    name = $"{window.dataGridView3.Rows[rowindex].Cells[0].Value}";
                }

                for (int i = 0; i < window.dataGridView2.Rows.Count; i++)
                {
                    string boolCondition = $"{window.dataGridView2.Rows[i].Cells[0].Value}";
                    string paramName = $"{window.dataGridView2.Rows[i].Cells[1].Value}";
                    string condition = $"{window.dataGridView2.Rows[i].Cells[2].Value}";
                    string paramValue = $"{window.dataGridView2.Rows[i].Cells[3].Value}";
                    Obj_Param_Condition objPC = new Obj_Param_Condition(boolCondition, paramName, condition, paramValue);
                    conditions_list.Add(objPC);
                }

                for (int i = 0; i < window.dataGridView1.Rows.Count; i++)
                {
                    string categoryName = $"{window.dataGridView1.Rows[i].Cells[0].Value}";

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

                foreach (Obj_Workset objWs in worksets_List)
                {
                    if (name == objWs.Name)
                    {
                        objWs.Conditions = conditions_list;
                        objWs.Categories = category_list;
                        objWs.Categories_and_Params = window.comboBox1.Text;
                    }
                }
            }

            void DeleteWorkSet(DataGridView dataGridView)
            {
                foreach (DataGridViewRow row in dataGridView.SelectedRows)
                {
                    dataGridView.Rows.Remove(row);
                    string name = $"{row.Cells[0].Value}";
                    for (int i = 0; i < worksets_List.Count; i++)
                    {
                        if (worksets_List[i].Name == name)
                        {
                            worksets_List.RemoveAt(i);
                            i++;
                        }
                    }
                }
            }

            void AddWorksetDialog()
            {
                InputDialogForm inputwindow = new InputDialogForm();
                inputwindow.Text = "Добавление рабочего набора";
                inputwindow.textBox1.Text = "Укажите название рабочего набора";
                inputwindow.button1.Click += (s, e) => { AddWorkset(window.dataGridView3, inputwindow); };
                inputwindow.button1.Text = "Ок";
                inputwindow.ShowDialog();
            }

            void AddWorkset(DataGridView dataGridView, InputDialogForm inputwindow)
            {
                bool condition = true;
                string name = inputwindow.textBox2.Text;
                foreach(Obj_Workset obj in worksets_List)
                {
                    if (name == obj.Name) condition = false;
                }
                if (condition)
                {
                    Obj_Workset objWs = new Obj_Workset(name);
                    worksets_List.Add(objWs);
                    string[] str = new string[1];
                    str[0] = objWs.Name;
                    dataGridView.Rows.Add(str);
                    inputwindow.Close();
                }
                else TaskDialog.Show("Предупреждение", "Рабочий набор с таким названием уже есть");
            }

            void SaveXml()
            {
                Obj_Workset[] objWorksets = new Obj_Workset[worksets_List.Count];

                for (int i = 0; i < worksets_List.Count; i++)
                {
                    objWorksets[i] = worksets_List[i];
                }

                XmlSerializer serializer = new XmlSerializer(typeof(Obj_Workset[]));
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Title = "Выберете папку для сохранения отчета";
                saveFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
                saveFileDialog.DefaultExt = ".xml";
                saveFileDialog.FileName = $"Рабочие наборы.xml";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string file_path = saveFileDialog.FileName;
                    string directory = saveFileDialog.InitialDirectory;
                    using (StreamWriter streamwriter = new StreamWriter($"{directory}{file_path}"))
                    {
                        serializer.Serialize(streamwriter, objWorksets);
                    }
                }
            }

            void Saving_when_closing()
            {
                Obj_Workset[] objWorksets = new Obj_Workset[worksets_List.Count];

                for (int i = 0; i < worksets_List.Count; i++)
                {
                    objWorksets[i] = worksets_List[i];
                }

                XmlSerializer serializer = new XmlSerializer(typeof(Obj_Workset[]));

                string userName = Environment.UserName;

                //string file_path = $@"C:\Users\{userName}\Documents\WorksetsTemp.xml";
                string file_path = $@"c:\ProgramData\WorksetsTemp.xml";
                using (StreamWriter streamwriter = new StreamWriter(file_path))
                {
                    serializer.Serialize(streamwriter, objWorksets);
                } 
            }

            void Load_xml_when_open()
            {
                string userName = Environment.UserName;
                window.dataGridView3.Rows.Clear();
                worksets_List.Clear();
                //string file_path = $@"C:\Users\{userName}\Documents\WorksetsTemp.xml";
                string file_path = $@"c:\ProgramData\WorksetsTemp.xml";
                using (StreamReader streamReader = new StreamReader(file_path))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(Obj_Workset[]));
                    Obj_Workset[] objWorksets = serializer.Deserialize(streamReader) as Obj_Workset[];
                    worksets_List = objWorksets.OrderBy(p => p.Name).ToList();

                    if (objWorksets != null)
                    {
                        foreach (Obj_Workset objWs in worksets_List)
                        {
                            string[] str = new string[1];
                            str[0] = objWs.Name;
                            window.dataGridView3.Rows.Add(str);
                        }
                    }
                }
            }

            void Copy_workset(DataGridView dataGridView)
            {
                string name = "";
                foreach (DataGridViewRow row in window.dataGridView3.SelectedRows)
                {
                    var rowindex = row.Index;
                    name = $"{window.dataGridView3.Rows[rowindex].Cells[0].Value}";
                }

                foreach (Obj_Workset objWs in worksets_List)
                {
                    if (name == objWs.Name)
                    {
                        Obj_Workset new_obj = objWs.Clone(objWs);
                        worksets_List.Add(new_obj);
                        string[] str = new string[1];
                        str[0] = new_obj.Name;
                        dataGridView.Rows.Add(str);
                        break;
                    }
                }
            }

            void LoadXml()
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    window.dataGridView3.Rows.Clear();
                    worksets_List.Clear();
                    string file_path = openFileDialog.FileName;
                    string directory = openFileDialog.InitialDirectory;
                    using (StreamReader streamReader = new StreamReader($"{directory}{file_path}"))
                    {
                        XmlSerializer serializer = new XmlSerializer(typeof(Obj_Workset[]));
                        Obj_Workset[] objWorksets = serializer.Deserialize(streamReader) as Obj_Workset[];
                        worksets_List = objWorksets.OrderBy(p => p.Name).ToList();

                        if (objWorksets != null)
                        {
                            foreach (Obj_Workset objWs in worksets_List)
                            {
                                string[] str = new string[1];
                                str[0] = objWs.Name;
                                window.dataGridView3.Rows.Add(str);
                            }
                        }
                    }
                }
            }

            void SetWsToElement()
            {
                dt_output.Clear();

                window.progressBar1.Value = 0;
                window.progressBar1.Maximum = worksets_List.Count;
                window.progressBar1.Step = 1;

                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();

                List<Element> elements_to_set_workset = new List<Element>();
                

                List<BuiltInCategory> all_builtin_categories_filter_list = new List<BuiltInCategory>();

                foreach (Category category in all_categories)
                {
                    if (category.CategoryType == CategoryType.Model
                        & (BuiltInCategory)category.Id.IntegerValue != BuiltInCategory.OST_Materials
                        & (BuiltInCategory)category.Id.IntegerValue != BuiltInCategory.OST_PipeSegments
                        & (BuiltInCategory)category.Id.IntegerValue != BuiltInCategory.OST_ProjectInformation
                        & (BuiltInCategory)category.Id.IntegerValue != BuiltInCategory.OST_DetailComponents
                        & (BuiltInCategory)category.Id.IntegerValue != BuiltInCategory.OST_Schedules
                        & (BuiltInCategory)category.Id.IntegerValue != BuiltInCategory.OST_Sheets
                        & (BuiltInCategory)category.Id.IntegerValue != BuiltInCategory.OST_Lines
                        & (BuiltInCategory)category.Id.IntegerValue != BuiltInCategory.INVALID
                        )
                    {
                        all_builtin_categories_filter_list.Add((BuiltInCategory)category.Id.IntegerValue);
                    }
                }
                all_builtin_categories_filter_list.Add(BuiltInCategory.OST_Grids);
                all_builtin_categories_filter_list.Add(BuiltInCategory.OST_Levels);
                all_builtin_categories_filter_list.Add(BuiltInCategory.OST_Cornices);
                all_builtin_categories_filter_list.Add(BuiltInCategory.OST_IOSModelGroups);

                BuiltInCategory[] all_builtin_categories_filter = new BuiltInCategory[all_builtin_categories_filter_list.Count];

                for (int i = 0; i < all_builtin_categories_filter_list.Count; i++)
                {
                    all_builtin_categories_filter[i] = all_builtin_categories_filter_list[i];
                }

                var all_elements = new FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(new ElementMulticategoryFilter(all_builtin_categories_filter)).Cast<Element>().ToList();

                using (Transaction transaction = new Transaction(doc))
                {
                    transaction.Start("Create Workset");
                    foreach (Obj_Workset objWs in worksets_List)
                    {
                        window.progressBar1.PerformStep();

                        dt_output.Rows.Add($"############# Набор: {objWs.Name} #################");

                        FilteredWorksetCollector user_worksets = new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset);

                        Workset unic_workset = user_worksets.FirstOrDefault(ws => ws.Name.Equals(objWs.Name));

                        if (unic_workset == null)
                        {
                            var new_workset = Workset.Create(doc, objWs.Name);
                        }
                        var worksetId = user_worksets.FirstOrDefault(ws => ws.Name.Equals(objWs.Name)).Id;

                        if (objWs.Categories.Count == 0 & objWs.Conditions.Count==0)
                        {
                            continue;
                        }

                        if (objWs.Categories.Count == 0)
                        {
                            elements_to_set_workset = all_elements;
                        }

                        List<BuiltInCategory> builtin_categories_filter_list = new List<BuiltInCategory>();

                        if (objWs.Categories.Count > 0)
                        {
                            foreach (Obj_Category category in objWs.Categories)
                            {
                                builtin_categories_filter_list.Add(category.BuiltIn);
                            }

                            BuiltInCategory[] builtin_categories_filter = new BuiltInCategory[builtin_categories_filter_list.Count];

                            for (int i = 0; i < builtin_categories_filter_list.Count; i++)
                            {
                                builtin_categories_filter[i] = builtin_categories_filter_list[i];
                            }

                            var elements_by_category = new FilteredElementCollector(doc).WherePasses(new ElementMulticategoryFilter(builtin_categories_filter)).WhereElementIsNotElementType().Cast<Element>().ToList();

                            if (objWs.Categories_and_Params == "и")
                            {
                                elements_to_set_workset = elements_by_category;
                            }
                            else if (objWs.Categories_and_Params == "или")
                            {
                                elements_to_set_workset = all_elements;
                            }
                        }

                        foreach (Element element in elements_to_set_workset)
                        {
                            List<bool> and_Condition = new List<bool>();
                            List<bool> or_Condition = new List<bool>();
                            bool bool_Result = false;

                            ObjRvt objRvt = new ObjRvt();
                            objRvt.elem = element;
                            dt_output.Rows.Add($"### Набор: {objWs.Name}. Элемент {element.Category.Name}: {element.Name}");

                            if (objWs.Categories_and_Params == "или")
                            {
                                if (builtin_categories_filter_list.Contains((BuiltInCategory)element.Category.Id.IntegerValue))
                                {
                                    bool_Result = true;
                                }
                                else
                                {
                                    bool_Result = ParameterCheck(objRvt, objWs);
                                }
                            }
                            else if (objWs.Categories_and_Params == "и")
                            {
                                bool_Result = ParameterCheck(objRvt, objWs);
                            }

                            if (objWs.Conditions.Count == 0) bool_Result = true;

                            if (bool_Result == true)
                            {
                                Parameter wsparam = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                                try
                                {
                                    dt_output.Rows.Add($"------------{bool_Result}");
                                    wsparam.Set(worksetId.IntegerValue);
                                }
                                catch { }
                            }
                        }
                    }
                    transaction.Commit();
                }
                stopwatch.Stop();
                TimeSpan elapsedTime = stopwatch.Elapsed;

                window.textBox1.Text = $"Время выполнения: {Math.Round(elapsedTime.TotalSeconds, 3).ToString()} сек.";
                window.progressBar1.Value = worksets_List.Count;
                window.dataGridDebug.DataSource = dt_output;
            }

            bool ParameterCheck(ObjRvt objRvt, Obj_Workset objWs)
            {
                List<bool> and_Condition = new List<bool>();
                List<bool> or_Condition = new List<bool>();
                bool bool_Result = false;
                
                foreach (Obj_Param_Condition param in objWs.Conditions)
                {
                    bool param_Condition = false;
                    var parameter_value = objRvt.GetParamAsString(param.ParamName);

                    if (param.Condition == "равно")
                    {
                        param_Condition = parameter_value == param.ParamValue;
                    }

                    if (param.Condition == "содержит")
                    {
                        try
                        {
                            param_Condition = parameter_value.Contains(param.ParamValue);
                        }
                        catch { }
                    }

                    if (param.Condition == "не равно")
                    {
                        param_Condition = !(parameter_value == param.ParamValue);
                    }

                    if (param.Condition == "не содержит")
                    {
                        try
                        {
                            param_Condition = !parameter_value.Contains(param.ParamValue);
                        }
                        catch { }
                    }

                    if (param_Condition)
                    {
                        bool temp_Param_Condition = true;
                        if (param.BoolCondition == "и")
                        {
                            and_Condition.Add(temp_Param_Condition);
                        }
                        else
                        {
                            or_Condition.Add(temp_Param_Condition);
                        }
                    }
                    else
                    {
                        bool temp_Param_Condition = false;
                        if (param.BoolCondition == "и")
                        {
                            and_Condition.Add(temp_Param_Condition);
                        }
                        else
                        {
                            or_Condition.Add(temp_Param_Condition);
                        }
                    }
                    dt_output.Rows.Add($"------{param.BoolCondition} {param.ParamName} {param.Condition} \"{param.ParamValue}\". Значение-\"{parameter_value}\"-{param_Condition}");
                }

                bool and_Result = false;
                bool or_Result = false;

                if (and_Condition.Count > 0) and_Result = and_Condition.All(condition => condition);
                else if (and_Condition.Count == 0) and_Result = true;
                if (or_Condition.Count > 0) or_Result = or_Condition.Any(condition => condition);
                else if (or_Condition.Count == 0) or_Result = true;
                
                bool_Result = and_Result & or_Result;
                dt_output.Rows.Add($"---------{bool_Result}");
                return bool_Result;
            }
            return Result.Succeeded;
        }
    }
}
