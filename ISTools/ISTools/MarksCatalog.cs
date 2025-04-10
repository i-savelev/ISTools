using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using ISTools.MyObjects;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;


namespace ISTools
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class MarksCatalog : IExternalCommand
    {
        //---PluginsManager---//
        public static string IS_TAB_NAME => "ISTools";
        public static string IS_NAME => "Каталог марок";
        public static string IS_IMAGE => "ISTools.Resources.MarksCatalog32.png";
        public static string IS_DESCRIPTION => "Команда позволяет формировать каталог значений параметров и заполнять указаныне параметры по значениям других параметров";
        //---PluginsManager---//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            string docName = doc.Title.Replace(".rvt", "");

            DataTable dtElems = new DataTable();
            dtElems.Columns.Add("Параметр", typeof(string));
            dtElems.Columns.Add("Значение", typeof(string));
            
            List<Element> elems = new List<Element>();

            IsToolsForm window = new IsToolsForm();
            window.dataGridView1.DataSource = dtElems;
            window.Text = "Обработка данных параметров модели";
            window.dataGridView1.Refresh();
            window.ExportButton.Text = "Собрать данные";
            
            window.toolStripButton1.Text = "Выбрать таблицу";
            window.toolStripButton2.Enabled = false;
            window.toolStripButton4.Enabled = false;
            window.toolStripTextBox1.Text = "<Укажите имя листа>";
            window.toolStripTextBox4.Text = "<Укажите имя листа>";
            window.toolStripTextBox3.Text = "<Укажите имя листа>";
            window.toolStripTextBox2.Text = "<Укажите параметр>";

            window.toolStripButton3.Text = "Загрузить данные";
            window.toolStripButton4.Text = "Записать данные";
            window.toolStripButton2.Text = "Собрать данные";

            window.groupBox1.Text = "Сохранение новой таблицы";
            window.groupBox2.Text = "Список семейтсв, которых нет в указанной таблице excel";
            window.groupBox3.Text = "Запись данных из таблицы в параметры элеменов";
            
            window.tabPage1.Text = "Сохранение новой таблицы";
            window.tabPage2.Text = "Дополнение существующей таблицы";
            window.tabPage3.Text = "Запись параметров";

            window.textBox1.Text = "Чтобы загрузить данные из таблицы необходимо указать имя листа и выбрать файл excel - кнопка \"Загрузить данные\". Список параметров и их значений отобразиться в таблице ниже. Параметры записываются для всех элементов на активном виде или специфиукации. Имена параметров - столбцы таблицы, наичиная с третьего. Для того, чтобы к значениям параметров добавился счетчик (001, 002, 003, ...), нужно добавить к названию столбца с названием параметра знак '*'. Чтобы записать данные в параметры элементов, необходимо нажать кнопку \"Записать данные\". Можно указать символ, который будет заполнен для пустых значений";
            window.textBox2.Text = "Для открытия существующего каталога, необходимо указать имя листа и выбрать файл excel. Для получения списка уникальных значений параметров, которых нет в каталоге, необходимо указать параметр и и нажать кнопку \"Собрать данные\". Для сохранения новых значений в таблицу необходимо нажать кнопку \"Сохранить новые данные в таблицу\"";
            window.textBox3.Text = "Для получения спска уникальных значений параметра, необходимо указать название параметра и нажать кнопку \"Собрать данные\". Для сохранения таблицы необходимо указать имя листа и нажать кнопку \"Сохранить таблицу\"";

            ToolStripSeparator toolStripSeparator = new ToolStripSeparator();
            ToolStripSeparator toolStripSeparator1 = new ToolStripSeparator();
            ToolStripSeparator toolStripSeparator2 = new ToolStripSeparator();
            ToolStripSeparator toolStripSeparator3 = new ToolStripSeparator();
            ToolStripButton stripButton = new ToolStripButton();
            ToolStripButton stripButton2 = new ToolStripButton();
            ToolStripTextBox toolStripTextBox = new ToolStripTextBox();
            ToolStripTextBox toolStripTextBox2 = new ToolStripTextBox();

            window.toolStrip1.Items.Add(toolStripSeparator1);
            window.toolStrip1.Items.Add(stripButton);

            window.toolStrip2.Items.Add(toolStripSeparator2);
            window.toolStrip2.Items.Add(toolStripTextBox);
            window.toolStrip2.Items.Add(toolStripSeparator3);
            window.toolStrip2.Items.Add(stripButton2);

            window.toolStrip3.Items.Add(toolStripSeparator);
            window.toolStrip3.Items.Add(toolStripTextBox2);

            stripButton.Text = "Сохранить таблицу";
            stripButton.BackColor = System.Drawing.SystemColors.ControlLight;
            stripButton.Enabled = false;

            stripButton2.Text = "Сохранить новые данные в таблицу";
            stripButton2.BackColor = System.Drawing.SystemColors.ControlLight;
            stripButton2.Enabled = false;

            toolStripTextBox.BackColor = System.Drawing.Color.White;
            toolStripTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            toolStripTextBox.Font = new System.Drawing.Font("Segoe UI", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            toolStripTextBox.Size = new System.Drawing.Size(170, 27);
            toolStripTextBox.Text = "<Укажите параметр>";

            toolStripTextBox2.BackColor = System.Drawing.Color.White;
            toolStripTextBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            toolStripTextBox2.Font = new System.Drawing.Font("Segoe UI", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            toolStripTextBox2.Size = new System.Drawing.Size(200, 27);
            toolStripTextBox2.Text = "<Заполнение пустых значений>";

            window.toolStripButton3.Click += (s, e) => { PapamPreview(); };
            window.toolStripButton1.Click += (s, e) => { TableOverride(); };
            window.ExportButton.Click += (s, e) => { SetParameterValuesToDt(); };
            stripButton.Click += (s, e) => { TableExport(); };

            window.ShowDialog();

            void SetParameterValuesToDt()
            {
                dtElems.Clear();
                var unique_param_value_list = CollectUniqueParameterValues(window.toolStripTextBox2.Text);
                window.toolStripProgressBar1.Value = 0;
                window.toolStripProgressBar1.Maximum = unique_param_value_list.Count;
                window.toolStripProgressBar1.Step = 1;
                foreach (var str in unique_param_value_list)
                {
                    dtElems.Rows.Add(
                        str[0],
                        str[1]
                        );
                    window.toolStripProgressBar1.PerformStep();
                }

                stripButton.Enabled = true;
                window.toolStripProgressBar1.Value = unique_param_value_list.Count;
            }

            List<string[]> CollectUniqueParameterValues(string paramName, List< string[]> outsideParamValueList = null)
            {
                if (outsideParamValueList == null)
                {
                    outsideParamValueList = new List<string[]>();
                }
                    
                var all_elements = new FilteredElementCollector(doc, doc.ActiveView.Id).
                    WhereElementIsNotElementType().
                    Cast<Element>().
                    ToList();
                ObjPlateCollector platesCollector = new ObjPlateCollector();
                List<ObjRvt> objRvtList = new List<ObjRvt>();
                platesCollector.doc = doc;
                var platesInJoint = platesCollector.GetPlates();

                foreach (Element elem in all_elements)
                {
                    ObjRvt objRvt = new ObjRvt();
                    objRvt.elem = elem;
                    objRvtList.Add(objRvt);
                }

                List<string> paramValueList = objRvtList.Select(obj => obj.GetParamAsString(paramName)).ToList();
                paramValueList.AddRange(platesInJoint.Select(obj => obj.GetParamAsString(paramName)).ToList());
                var uniqueParamValueList = paramValueList.Distinct().ToList();
                List<string[]> paramValueOutlist = new List<string[]>();
                foreach (var paramValue in uniqueParamValueList)
                    if (!outsideParamValueList.Any(str => $"{str[0]}_{str[1]}" == $"{paramName}_{paramValue}"))
                    {
                        paramValueOutlist.Add(new string[] { paramName, paramValue });
                        outsideParamValueList.Add(new string[] { paramName, paramValue });
                    }
                return paramValueOutlist;
            }

            // export database to oexcel
            void TableExport()
            {
                window.toolStripProgressBar1.Value = 0;
                window.toolStripProgressBar1.Maximum = 1;
                window.toolStripProgressBar1.Step = 1;
                ObjExcelTable table = new ObjExcelTable();
                table.filename = $"Каталог параметров_{docName}";
                table.dt = dtElems;
                table.worksheets = window.toolStripTextBox4.Text;
                table.TableExport();
                window.toolStripProgressBar1.Value = 1;
            }

            void TableOverride()
            {
                OpenFileDialog saveFileDialog1 = new OpenFileDialog();
                window.dataGridView2.Rows.Clear();
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    FileInfo existingFile = new FileInfo(saveFileDialog1.FileName);
                    using (ExcelPackage excelPackage = new ExcelPackage(existingFile))
                    {
                        ExcelPackage excelpackage = new ExcelPackage(existingFile);
                        stripButton2.Click += (s, e) => { IsUtils.Save_opened_excel(excelpackage, window.toolStripProgressBar2); };
                        IsUtils.AddColumnHeaderFromExcel(excelpackage,
                            window.toolStripTextBox1.Text,
                            window.dataGridView2, 1, 1, 2
                            );
                        var rowCount = IsUtils.GetRowExcelCount(excelpackage, window.toolStripTextBox1.Text);
                        var excelData = IsUtils.GetListOfStringFromExcel(excelpackage, window.toolStripTextBox1.Text);
                        window.toolStripButton2.Click += (s, e) => { CollectUniqueParameterValuesToOverride(excelpackage, rowCount, excelData); }; 
                        window.toolStripButton2.Enabled = true;
                    }
                }
            }

            void CollectUniqueParameterValuesToOverride(ExcelPackage exel_package, int row_count, List<string[]> excelВata)
            {
                var uniqueParamValueList = CollectUniqueParameterValues(toolStripTextBox.Text, excelВata);
                IsUtils.SetStringForExcel(uniqueParamValueList, exel_package, window.toolStripTextBox1.Text, row_count);
                
                foreach (string[] str in uniqueParamValueList)
                {
                    window.dataGridView2.Rows.Add(str);
                }
                stripButton2.Enabled = true;
            }

            void PapamPreview()
            {
                OpenFileDialog saveFileDialog1 = new OpenFileDialog();
                
                window.dataGridView3.Rows.Clear();
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    FileInfo existingFile = new FileInfo(saveFileDialog1.FileName);
                    using (ExcelPackage excelPackage = new ExcelPackage(existingFile))
                    {
                        ExcelPackage exlPackage = new ExcelPackage(existingFile);

                        ExcelPackage excel = new ExcelPackage(existingFile);

                        IsUtils.AddColumnHeaderFromExcel(excel,
                            window.toolStripTextBox3.Text,
                            window.dataGridView3, 1
                            );
                        var parametersDict = CreateParametersDict(excel, window.toolStripTextBox3.Text);
                        var rowList = IsUtils.GetListOfStringFromExcel(excel, window.toolStripTextBox3.Text);
                        foreach (string[] str in rowList)
                        {
                            window.dataGridView3.Rows.Add(str);
                        }
                        window.toolStripButton4.Enabled = true;
                        window.toolStripButton4.Click += (s, e) => { Fill(parametersDict); };
                    }
                }
            }

            void Fill(Dictionary<string, Dictionary<string, string>> parametersDict)
            {
                var allElems = new FilteredElementCollector(doc, doc.ActiveView.Id).
                    WhereElementIsNotElementType().
                    Cast<Element>().
                    ToList();

                ObjPlateCollector platesCollector = new ObjPlateCollector();
                platesCollector.doc = doc;
                var platesInJoint = platesCollector.GetPlates();

                window.toolStripProgressBar3.Value = 0;
                window.toolStripProgressBar3.Maximum = allElems.Count;
                window.toolStripProgressBar3.Step = 1;

                Dictionary<string, int> paramValueCountDict = new Dictionary<string, int>();
                List<string> unicueParameters = new List<string>();
                foreach (var param in parametersDict)
                {
                    foreach (string key in param.Value.Keys)
                    {
                        if (!unicueParameters.Contains(key.Split('&')[0]))
                        {
                            unicueParameters.Add(key.Split('&')[0]);
                        }
                    }
                }
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Заполнение параметров");
                    foreach (var el in allElems)
                    {
                        ObjRvt objRvt = new ObjRvt();
                        objRvt.elem = el;
                        window.toolStripProgressBar3.PerformStep();
                        foreach (var param in parametersDict)
                        {
                            var str = "";
                            var count = "";
                            foreach (string par in unicueParameters)
                            {
                                try
                                {
                                    if (objRvt.GetParam(par).ToString() != "")
                                    {
                                        string vol = $"{par}&{objRvt.GetParam(par).ToString()}";
                                        if (!paramValueCountDict.ContainsKey(param.Value[vol]))
                                        {
                                            paramValueCountDict.Add(param.Value[vol], 0);
                                        }
                                        paramValueCountDict[param.Value[vol]] += 1;
                                        if (param.Key.Contains("*"))
                                        {
                                            count = $"{paramValueCountDict[param.Value[vol]]:000}";
                                        }
                                        objRvt.SetParam(param.Key.Replace("*", ""), $"{param.Value[vol]}{count}");
                                        str = $"{param.Value[vol]}{count}";
                                    }
                                }
                                catch { }
                            }
                            try
                            {
                                if (str == "")
                                {
                                    objRvt.SetParam(param.Key.Replace("*", ""), $"{toolStripTextBox2.Text}");
                                }
                            }
                            catch { }
                        }
                    }

                    foreach (var pij in platesInJoint)
                    {
                        foreach (var param in parametersDict)
                        {
                            var str = "";
                            var count = "";
                            foreach (string par in unicueParameters)
                            {
                                try
                                {
                                    if (pij.GetParam(par).ToString() != "")
                                    {
                                        string vol = $"{par}&{pij.GetParam(par).ToString()}";
                                        if (!paramValueCountDict.ContainsKey(param.Value[vol]))
                                        {
                                            paramValueCountDict.Add(param.Value[vol], 0);
                                        }
                                        paramValueCountDict[param.Value[vol]] += 1;
                                        if (param.Key.Contains("*"))
                                        {
                                            count = $"{paramValueCountDict[param.Value[vol]]:000}";
                                        }
                                        pij.SetParam(param.Key.Replace("*", ""), $"{param.Value[vol]}{count}");
                                        str = $"{param.Value[vol]}{count}";
                                    }
                                }
                                catch{}
                            }
                            try
                            {
                                if (str == "")
                                {
                                    pij.SetParam(param.Key.Replace("*", ""), $"{toolStripTextBox2.Text}");
                                }
                            }
                            catch { }
                        }
                    }
                    window.toolStripProgressBar3.Value = allElems.Count;
                    tx.Commit();
                }
            }
            return Result.Succeeded;
            }

        Dictionary<string, Dictionary<string, string>> CreateParametersDict(ExcelPackage excelPackage, string worksheetName)
        {
            Dictionary<string, Dictionary<string, string>> parametersDict = new Dictionary<string, Dictionary<string, string>>();
            foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
            {
                if (worksheet.Name == worksheetName)
                {
                    for (int col = worksheet.Dimension.Start.Column + 2; col <= worksheet.Dimension.End.Column; col++)
                    {
                        if (worksheet.Cells[1, col].Value != null)
                        {
                            if (worksheet.Cells[1, col].Value.ToString() != "")
                            {
                                Dictionary<string, string> parameterDict = new Dictionary<string, string>();
                                for (int row = worksheet.Dimension.Start.Row + 1; row <= worksheet.Dimension.End.Row; row++)
                                {
                                    if (worksheet.Cells[row, 1].Value != null)
                                    {
                                        if (worksheet.Cells[row, 1].Value.ToString() != "")
                                        {
                                            try
                                            {
                                                parameterDict.Add($"{worksheet.Cells[row, 1].Value}&{worksheet.Cells[row, 2].Value}", $"{worksheet.Cells[row, col].Value}");
                                            }
                                            catch { }

                                        }
                                    }
                                }
                                parametersDict.Add(worksheet.Cells[1, col].Value.ToString(), parameterDict);
                            }
                        }
                    }
                }
            }
            return parametersDict;
        }
    }
}

