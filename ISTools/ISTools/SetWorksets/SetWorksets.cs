using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ISTools.Forms;
using ISTools.MyObjects;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Serialization;
using Document = Autodesk.Revit.DB.Document;


namespace ISTools
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class SetWorksets : IExternalCommand
    {
        //---PluginsManager---//
        public static string IS_TAB_NAME => "ISTools";
        public static string IS_NAME => "Рабочие наборы";
        public static string IS_IMAGE => "ISTools.Resources.Worksets32.png";
        public static string IS_DESCRIPTION => "Инструкция:\nСкрипт автоматически присваевает элементам рабочие наборы по заданным правилам";
        //---PluginsManager---//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            List<ObjWorkset> worksetsList = new List<ObjWorkset>();
            List<string> categoriesList = new List<string>();
            List<string> parameterConditionList = new List<string>() { "равно", "содержит", "не содержит", "не равно" };
            List<string> parameterBoolConditionList = new List<string>() { "и", "или" };
            List<string> parameterBoolConditionList_2 = new List<string>() { "и", "или" };

            DataTable dtOutput = new DataTable();
            dtOutput.Columns.Add("1");

            var allCategories = doc.Settings.Categories;
            foreach (Category category in allCategories)
            {
                dtOutput.Rows.Add($"{category.Name} - {(BuiltInCategory)category.Id.IntegerValue}");
                if (category.CategoryType == CategoryType.Model)
                {
                    categoriesList.Add(category.Name);

                }
            }

            categoriesList.Add("Оси");
            categoriesList.Add("Уровни");
            categoriesList.Add("Выступающие профили");
            categoriesList.Sort();

            SetWorksetsForm window = new SetWorksetsForm();
            window.Text = "Настройка рабочих наборов";
            //datagrid categories
            window.groupBox1.Text = "Категории";
            DataGridViewComboBoxColumn cmb = new DataGridViewComboBoxColumn();
            cmb.HeaderText = "Категория";
            cmb.DataSource = categoriesList;
            window.dataGridView1.Columns.Add(cmb);
            window.button7.Text = "+";
            window.button8.Text = "-";
            window.button7.Click += (s, e) => { IsUtils.AddRow(window.dataGridView1); };
            window.button8.Click += (s, e) => { IsUtils.DeleteRow(window.dataGridView1); };

            //datagrid parameters
            window.groupBox2.Text = "Параметры";
            DataGridViewColumn columnParameterName = new DataGridViewColumn();
            columnParameterName.HeaderText = "Название параметра";
            columnParameterName.CellTemplate = new DataGridViewTextBoxCell();
            DataGridViewComboBoxColumn columnParameterCondition = new DataGridViewComboBoxColumn();
            columnParameterCondition.DataSource = parameterConditionList;
            columnParameterCondition.HeaderText = "Условние";
            DataGridViewComboBoxColumn columnParameterBoolCondition = new DataGridViewComboBoxColumn();
            columnParameterBoolCondition.DataSource = parameterBoolConditionList;
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
            window.button4.Click += (s, e) => { IsUtils.DeleteRow(window.dataGridView2); };
            window.button2.Click += (s, e) => { IsUtils.AddRow(window.dataGridView2); };

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
            window.dataGridView3.CellClick += SetDataToOtherDatagrid;
            window.button10.Click += (s, e) => { RenameWorkSetDialog(); };
            window.button10.Text = "Переименовать";
            window.button12.Text = "Копировать";
            window.button12.Click += (s, e) => { CopyWorkset(window.dataGridView3); };

            IsHandlerSwitch handelr = new IsHandlerSwitch((ap) =>
            {
                SetWsToElement();
            });

            window.comboBox1.DataSource = parameterBoolConditionList_2;
            window.groupBox3.Text = "Настройки";
            window.button3.Text = "Сохранить рабочий набор";
            window.button3.Click += (s, e) => { SaveWorkset(); };
            window.button6.Click += (s, e) => { SaveXml(); };
            window.button6.Text = "Сохранить шаблон";
            window.button9.Text = "Загрузить шаблон";
            window.button9.Click += (s, e) => { LoadXml(); };
            window.button11.Text = "Распределить элементы по наборам";
            window.button11.Click += (s, e) => { handelr.ExternalEvent.Raise(); };
            try
            {
                window.FormClosed += (s, e) => { SavingWhenClosing(); };
            }
            catch { }

            try
            {
                LoadXmlWhenPpen();
            }
            catch { }

            window.groupBox5.Text = "Информация";
            window.textBox2.Text = "Данный скрипт позволяет автоматически распределять элементы по рабочим наборам.";

            window.Show();

            void SetDataToOtherDatagrid(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0) return;
                window.dataGridView2.Rows.Clear();
                window.dataGridView1.Rows.Clear();
                var rowindex = e.RowIndex;
                foreach (var objWs in worksetsList)
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

                        if (objWs.CategoriesAndParams != null)
                        {
                            window.comboBox1.Text = objWs.CategoriesAndParams;
                        }
                    }
                }
            }

            void RenameWorkset(InputDialogForm inputWindow)
            {
                string name = "";
                List<ObjParamCondition> conditions = new List<ObjParamCondition>();
                foreach (DataGridViewRow row in window.dataGridView3.SelectedRows)
                {
                    var rowindex = row.Index;
                    name = $"{window.dataGridView3.Rows[rowindex].Cells[0].Value}";
                }

                foreach (var ws in worksetsList)
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
                List<ObjParamCondition> conditionsList = new List<ObjParamCondition>();
                List<ObjCategory> categoryList = new List<ObjCategory>();
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
                    ObjParamCondition objPC = new ObjParamCondition(boolCondition, paramName, condition, paramValue);
                    conditionsList.Add(objPC);
                }

                for (int i = 0; i < window.dataGridView1.Rows.Count; i++)
                {
                    string categoryName = $"{window.dataGridView1.Rows[i].Cells[0].Value}";

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

                foreach (ObjWorkset objWs in worksetsList)
                {
                    if (name == objWs.Name)
                    {
                        objWs.Conditions = conditionsList;
                        objWs.Categories = categoryList;
                        objWs.CategoriesAndParams = window.comboBox1.Text;
                    }
                }
            }

            void DeleteWorkSet(DataGridView dataGridView)
            {
                foreach (DataGridViewRow row in dataGridView.SelectedRows)
                {
                    dataGridView.Rows.Remove(row);
                    string name = $"{row.Cells[0].Value}";
                    for (int i = 0; i < worksetsList.Count; i++)
                    {
                        if (worksetsList[i].Name == name)
                        {
                            worksetsList.RemoveAt(i);
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
                foreach (ObjWorkset obj in worksetsList)
                {
                    if (name == obj.Name) condition = false;
                }
                if (condition)
                {
                    ObjWorkset objWs = new ObjWorkset(name);
                    worksetsList.Add(objWs);
                    string[] str = new string[1];
                    str[0] = objWs.Name;
                    dataGridView.Rows.Add(str);
                    inputwindow.Close();
                }
                else TaskDialog.Show("Предупреждение", "Рабочий набор с таким названием уже есть");
            }

            void SaveXml()
            {
                ObjWorkset[] objWorksets = new ObjWorkset[worksetsList.Count];

                for (int i = 0; i < worksetsList.Count; i++)
                {
                    objWorksets[i] = worksetsList[i];
                }

                XmlSerializer serializer = new XmlSerializer(typeof(ObjWorkset[]));
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

            void SavingWhenClosing()
            {
                ObjWorkset[] objWorksets = new ObjWorkset[worksetsList.Count];

                for (int i = 0; i < worksetsList.Count; i++)
                {
                    objWorksets[i] = worksetsList[i];
                }

                XmlSerializer serializer = new XmlSerializer(typeof(ObjWorkset[]));

                string userName = Environment.UserName;
                string file_path = $@"c:\ProgramData\WorksetsTemp.xml";
                using (StreamWriter streamwriter = new StreamWriter(file_path))
                {
                    serializer.Serialize(streamwriter, objWorksets);
                }
            }

            void LoadXmlWhenPpen()
            {
                string userName = Environment.UserName;
                window.dataGridView3.Rows.Clear();
                worksetsList.Clear();
                string filePath = $@"c:\ProgramData\WorksetsTemp.xml";
                using (StreamReader streamReader = new StreamReader(filePath))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(ObjWorkset[]));
                    ObjWorkset[] objWorksets = serializer.Deserialize(streamReader) as ObjWorkset[];
                    worksetsList = objWorksets.OrderBy(p => p.Name).ToList();

                    if (objWorksets != null)
                    {
                        foreach (ObjWorkset objWs in worksetsList)
                        {
                            string[] str = new string[1];
                            str[0] = objWs.Name;
                            window.dataGridView3.Rows.Add(str);
                        }
                    }
                }
            }

            void CopyWorkset(DataGridView dataGridView)
            {
                string name = "";
                foreach (DataGridViewRow row in window.dataGridView3.SelectedRows)
                {
                    var rowindex = row.Index;
                    name = $"{window.dataGridView3.Rows[rowindex].Cells[0].Value}";
                }

                foreach (ObjWorkset objWs in worksetsList)
                {
                    if (name == objWs.Name)
                    {
                        var newObjWs = objWs.Clone(objWs);
                        worksetsList.Add(newObjWs);
                        string[] str = new string[1];
                        str[0] = newObjWs.Name;
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
                    worksetsList.Clear();
                    string filePath = openFileDialog.FileName;
                    string directory = openFileDialog.InitialDirectory;
                    using (StreamReader streamReader = new StreamReader($"{directory}{filePath}"))
                    {
                        XmlSerializer serializer = new XmlSerializer(typeof(ObjWorkset[]));
                        ObjWorkset[] objWorksets = serializer.Deserialize(streamReader) as ObjWorkset[];
                        worksetsList = objWorksets.OrderBy(p => p.Name).ToList();

                        if (objWorksets != null)
                        {
                            foreach (ObjWorkset objWs in worksetsList)
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
                dtOutput.Clear();

                window.progressBar1.Value = 0;
                window.progressBar1.Maximum = worksetsList.Count;
                window.progressBar1.Step = 1;

                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();

                List<Element> ListElementsToSetWorkset = new List<Element>();


                List<BuiltInCategory> allBuiltinCategoriesFilterList = new List<BuiltInCategory>();

                foreach (Category category in allCategories)
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
                        allBuiltinCategoriesFilterList.Add((BuiltInCategory)category.Id.IntegerValue);
                    }
                }
                allBuiltinCategoriesFilterList.Add(BuiltInCategory.OST_Grids);
                allBuiltinCategoriesFilterList.Add(BuiltInCategory.OST_Levels);
                allBuiltinCategoriesFilterList.Add(BuiltInCategory.OST_Cornices);
                allBuiltinCategoriesFilterList.Add(BuiltInCategory.OST_IOSModelGroups);

                BuiltInCategory[] allBuiltinCategoriesFilter = new BuiltInCategory[allBuiltinCategoriesFilterList.Count];

                for (int i = 0; i < allBuiltinCategoriesFilterList.Count; i++)
                {
                    allBuiltinCategoriesFilter[i] = allBuiltinCategoriesFilterList[i];
                }

                var allElements = new FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(new ElementMulticategoryFilter(allBuiltinCategoriesFilter)).Cast<Element>().ToList();

                using (Transaction transaction = new Transaction(doc))
                {
                    transaction.Start("Create Workset");
                    foreach (ObjWorkset objWs in worksetsList)
                    {
                        window.progressBar1.PerformStep();

                        dtOutput.Rows.Add($"############# Набор: {objWs.Name} #################");

                        FilteredWorksetCollector userWorksets = new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset);

                        Workset unicWorkset = userWorksets.FirstOrDefault(ws => ws.Name.Equals(objWs.Name));

                        if (unicWorkset == null)
                        {
                            var new_workset = Workset.Create(doc, objWs.Name);
                        }
                        var worksetId = userWorksets.FirstOrDefault(ws => ws.Name.Equals(objWs.Name)).Id;

                        if (objWs.Categories.Count == 0 & objWs.Conditions.Count == 0)
                        {
                            continue;
                        }

                        if (objWs.Categories.Count == 0)
                        {
                            ListElementsToSetWorkset = allElements;
                        }

                        List<BuiltInCategory> builtinCategoriesFilterList = new List<BuiltInCategory>();

                        if (objWs.Categories.Count > 0)
                        {
                            foreach (ObjCategory category in objWs.Categories)
                            {
                                builtinCategoriesFilterList.Add(category.BuiltIn);
                            }

                            BuiltInCategory[] builtinCategoriesFilter = new BuiltInCategory[builtinCategoriesFilterList.Count];

                            for (int i = 0; i < builtinCategoriesFilterList.Count; i++)
                            {
                                builtinCategoriesFilter[i] = builtinCategoriesFilterList[i];
                            }

                            var elementsByCategory = new FilteredElementCollector(doc).WherePasses(new ElementMulticategoryFilter(builtinCategoriesFilter)).WhereElementIsNotElementType().Cast<Element>().ToList();

                            if (objWs.CategoriesAndParams == "и")
                            {
                                ListElementsToSetWorkset = elementsByCategory;
                            }
                            else if (objWs.CategoriesAndParams == "или")
                            {
                                ListElementsToSetWorkset = allElements;
                            }
                        }

                        foreach (Element element in ListElementsToSetWorkset)
                        {
                            List<bool> andCondition = new List<bool>();
                            List<bool> orCondition = new List<bool>();
                            bool boolResult = false;

                            ObjRvt objRvt = new ObjRvt();
                            objRvt.elem = element;
                            dtOutput.Rows.Add($"### Набор: {objWs.Name}. Элемент {element.Category.Name}: {element.Name}");

                            if (objWs.CategoriesAndParams == "или")
                            {
                                if (builtinCategoriesFilterList.Contains((BuiltInCategory)element.Category.Id.IntegerValue))
                                {
                                    boolResult = true;
                                }
                                else
                                {
                                    boolResult = ParameterCheck(objRvt, objWs);
                                }
                            }
                            else if (objWs.CategoriesAndParams == "и")
                            {
                                boolResult = ParameterCheck(objRvt, objWs);
                            }

                            if (objWs.Conditions.Count == 0) boolResult = true;

                            if (boolResult == true)
                            {
                                Parameter wsParam = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                                try
                                {
                                    dtOutput.Rows.Add($"------------{boolResult}");
                                    wsParam.Set(worksetId.IntegerValue);
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
                window.progressBar1.Value = worksetsList.Count;
                window.dataGridDebug.DataSource = dtOutput;
            }

            bool ParameterCheck(ObjRvt objRvt, ObjWorkset objWs)
            {
                List<bool> andCondition = new List<bool>();
                List<bool> orCondition = new List<bool>();
                bool boolResult = false;

                foreach (ObjParamCondition param in objWs.Conditions)
                {
                    bool paramCondition = false;
                    var parameterValue = objRvt.GetParamAsString(param.ParamName);

                    if (param.Condition == "равно")
                    {
                        paramCondition = parameterValue == param.ParamValue;
                    }

                    if (param.Condition == "содержит")
                    {
                        try
                        {
                            paramCondition = parameterValue.Contains(param.ParamValue);
                        }
                        catch { }
                    }

                    if (param.Condition == "не равно")
                    {
                        paramCondition = !(parameterValue == param.ParamValue);
                    }

                    if (param.Condition == "не содержит")
                    {
                        try
                        {
                            paramCondition = !parameterValue.Contains(param.ParamValue);
                        }
                        catch { }
                    }

                    if (paramCondition)
                    {
                        bool temp_Param_Condition = true;
                        if (param.BoolCondition == "и")
                        {
                            andCondition.Add(temp_Param_Condition);
                        }
                        else
                        {
                            orCondition.Add(temp_Param_Condition);
                        }
                    }
                    else
                    {
                        bool temp_Param_Condition = false;
                        if (param.BoolCondition == "и")
                        {
                            andCondition.Add(temp_Param_Condition);
                        }
                        else
                        {
                            orCondition.Add(temp_Param_Condition);
                        }
                    }
                    dtOutput.Rows.Add($"------{param.BoolCondition} {param.ParamName} {param.Condition} \"{param.ParamValue}\". Значение-\"{parameterValue}\"-{paramCondition}");
                }

                bool andResult = false;
                bool orResult = false;

                if (andCondition.Count > 0) andResult = andCondition.All(condition => condition);
                else if (andCondition.Count == 0) andResult = true;
                if (orCondition.Count > 0) orResult = orCondition.Any(condition => condition);
                else if (orCondition.Count == 0) orResult = true;

                boolResult = andResult & orResult;
                dtOutput.Rows.Add($"---------{boolResult}");
                return boolResult;
            }
            return Result.Succeeded;
        }
    }
}
