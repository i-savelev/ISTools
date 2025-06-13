using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;

namespace ISTools
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ParamByCat : IExternalCommand
    {
        //---PluginsManager---//
        public static string IS_TAB_NAME => "ISTools";
        public static string IS_NAME => "Парметры по категории";
        public static string IS_IMAGE => "ISTools.Resources.ParamByCat32.png";
        public static string IS_DESCRIPTION => "Команда заполняет параметры по категориям\nПодробнее: https://github.com/i-savelev/ISTools/wiki/Каталог-семейств";
        //---PluginsManager---//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            DataTable dtCategories = new DataTable();
            dtCategories.Columns.Add("Название параметра", typeof(string));
            dtCategories.Columns.Add("Категория 1", typeof(string));
            dtCategories.Columns.Add("Категория 2", typeof(string));
            dtCategories.Columns.Add("Категория n", typeof(string));
            Dictionary<string, Dictionary<string, string>> parametersDict = new Dictionary<string, Dictionary<string, string>>();
            dtCategories.Rows.Add(
                "Парамемтр 1",
                "Значение 1;Значение 2...",
                "Значение 1;Значение 2...",
                "Значение 1 ..."
                );

            IsToolsForm window = new IsToolsForm();
            window.Text = "Заполнение параметров по категории";
            window.tabPage1.Text = "";
            window.groupBox1.Text = "";
            window.tabControl1.TabPages.RemoveByKey("tabPage2");
            window.tabControl1.TabPages.RemoveByKey("tabPage3");
            window.ExportButton.Text = "Сохранить шаблон";
            window.toolStripTextBox2.Text = "<Укажите имя листа>";
            window.textBox3.Text = "Кнопка \"Cохранить\" шаблон позволяет сохранить шаблон таблицы эксель для заполнения. Кнопка \"Загрузить\" загружает выбранную таблицу. Кнопка \"Записать\" записывает значения в параметры элементов. Плагин дает воможность добавлять \"счетчик\", для этого нужно дополнить название параметра символом \"*\". Если указать значения параметров через точку с запятой \";\", то при записи будет выбрано случайное из них. Плагин заполняет параметры для элементов на активном виде или спецификации";
            window.textBox3.Height = 90;
            window.ExportButton.Click += (s, e) => { TableExport(); };
            window.toolStripTextBox4.Text = "<Разделитель>";
            window.toolStrip1.Items.RemoveByKey("toolStripTextBox4");
            ToolStripButton stripButton = new ToolStripButton();
            window.toolStrip1.Items.Add(stripButton);
            stripButton.Text = "Загрузить";
            stripButton.BackColor = System.Drawing.SystemColors.ControlLight;
            ToolStripButton stripButton1 = new ToolStripButton();
            window.toolStrip1.Items.Add(stripButton1);
            stripButton1.Text = "Записать";
            stripButton1.BackColor = System.Drawing.SystemColors.ControlLight;
            stripButton1.Enabled = false;
            stripButton.Click += (s, e) => { PapamFill(); };

            window.ShowDialog();

            void TableExport()
            {
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    ObjExcelTable table = new ObjExcelTable();
                    table.filename = "Шаблон";
                    table.dt = dtCategories;
                    table.worksheets = window.toolStripTextBox2.Text;
                    table.TableExport();
                }
            }

            void PapamFill()
            {
                OpenFileDialog saveFileDialog1 = new OpenFileDialog();
                window.dataGridView3.Rows.Clear();
                window.dataGridView1.Rows.Clear();
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    FileInfo existingFile = new FileInfo(saveFileDialog1.FileName);
                    using (ExcelPackage excelPackage = new ExcelPackage(existingFile))
                    {
                        ExcelPackage excel = new ExcelPackage(existingFile);
                        foreach (ExcelWorksheet worksheet in excel.Workbook.Worksheets)
                        {
                            if (worksheet.Name == window.toolStripTextBox2.Text)
                            {
                                IsUtils.AddColumnHeaderFromExcel(excel,
                                    window.toolStripTextBox2.Text,
                                    window.dataGridView1, 1
                                    );
                                for (int j = worksheet.Dimension.Start.Column + 1; j <= worksheet.Dimension.End.Column; j++)
                                {
                                    if (worksheet.Cells[1, j].Value != null) 
                                    {
                                        if (worksheet.Cells[1, j].Value.ToString() != "")
                                        {
                                            Dictionary<string, string> paramDict = new Dictionary<string, string>();
                                            for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                                            {
                                                if (worksheet.Cells[i, 1].Value != null)
                                                {
                                                    if (worksheet.Cells[i, 1].Value.ToString() != "")
                                                    {
                                                        try
                                                        {
                                                            paramDict.Add($"{worksheet.Cells[i, 1].Value}", $"{worksheet.Cells[i, j].Value}");
                                                        }
                                                        catch { }
                                                    }
                                                }
                                            }
                                            parametersDict.Add(worksheet.Cells[1, j].Value.ToString(), paramDict);
                                        }
                                    }
                                }
                                var rowList = IsUtils.GetListOfStringFromExcel(excel, window.toolStripTextBox2.Text);
                                foreach (string[] str in rowList)
                                {
                                    try
                                    {
                                        window.dataGridView1.Rows.Add(str);
                                    }
                                    catch{}
                                }
                            }
                        }
                        stripButton1.Enabled = true;
                        stripButton1.Click += (s, e) => { Fill(); };
                    }
                }
            }
            void Fill()
            {
                Random rnd = new Random();
                var allElems = new FilteredElementCollector(doc, doc.ActiveView.Id).
                    WhereElementIsNotElementType().
                    Cast<Element>().
                    ToList();

                ObjPlateCollector platesCollector = new ObjPlateCollector();
                platesCollector.doc = doc;
                var platesInJoint = platesCollector.GetPlates();

                window.toolStripProgressBar1.Value = 0;
                window.toolStripProgressBar1.Maximum = allElems.Count + platesInJoint.Count;
                window.toolStripProgressBar1.Step = 1;

                Dictionary<string, int> paramCount = new Dictionary<string, int>();
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Заполнение параметров семейств");
                    foreach (var el in allElems)
                    {
                        window.toolStripProgressBar1.PerformStep();
                        ObjRvt objRvt = new ObjRvt();
                        objRvt.elem = el;
                        var catName = "";
                        try
                        {
                            catName = el.Category.Name;
                        }
                        catch{}
                        if (parametersDict.ContainsKey(catName))
                        {
                            foreach (var param in parametersDict[catName])
                            {
                                int index = rnd.Next(param.Value.Split(';').Length);
                                window.toolStripProgressBar1.PerformStep();
                                var count = "";

                                if (!paramCount.ContainsKey(param.Value.Split(';')[index]))
                                {
                                    paramCount.Add(param.Value.Split(';')[index], 0);
                                }
                                paramCount[param.Value.Split(';')[index]] += 1;
                                if (param.Key.Contains("*"))
                                {
                                    count = $"{paramCount[param.Value.Split(';')[index]]:000}";
                                }
                                try
                                {
                                    objRvt.SetParam(param.Key.Replace("*", ""), $"{param.Value.Split(';')[index]}{count}");
                                }
                                catch { }
                            }
                        }
                    }

                    foreach (var pij in platesInJoint)
                    {
                        window.toolStripProgressBar1.PerformStep();
                        var catName = "Пластины";
                        if (parametersDict.ContainsKey(catName))
                        {
                            foreach (var param in parametersDict[catName])
                            {
                                int index = rnd.Next(param.Value.Split(';').Length);
                                window.toolStripProgressBar3.PerformStep();
                                var count = "";

                                if (!paramCount.ContainsKey(param.Value.Split(';')[index]))
                                {
                                    paramCount.Add(param.Value.Split(';')[index], 0);
                                }
                                paramCount[param.Value.Split(';')[index]] += 1;
                                if (param.Key.Contains("*"))
                                {
                                    count = $"{paramCount[param.Value.Split(';')[index]]:000}";
                                }
                                try
                                {
                                    pij.SetParam(param.Key.Replace("*", ""), $"{param.Value.Split(';')[index]}{count}");
                                }
                                catch { }
                            }
                        }
                    }
                    window.toolStripProgressBar1.Value = allElems.Count + platesInJoint.Count;
                    tx.Commit();
                }
            }
            return Result.Succeeded;
        }
    }
}
