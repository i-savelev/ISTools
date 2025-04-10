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
    public class ParamMapping : IExternalCommand
    {
        //---PluginsManager---//
        public static string IS_TAB_NAME => "ISTools";
        public static string IS_NAME => "Мэппинг параметров";
        public static string IS_IMAGE => "ISTools.Resources.ParamMapping32.png";
        public static string IS_DESCRIPTION => "Команда заполняет согласно правилам мэппинга";
        //---PluginsManager---//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            DataTable dtCategories = new DataTable();
            dtCategories.Columns.Add("Исходный параметр", typeof(string));
            dtCategories.Columns.Add("Параметр для записи", typeof(string));

            dtCategories.Rows.Add(
                "Исходный параметр1",
                "Параметр для записи1"
                );

            IsToolsForm window = new IsToolsForm();
            window.Text = "Мэппинг параметров";
            window.tabPage1.Text = "";
            window.groupBox1.Text = "";
            window.tabControl1.TabPages.RemoveByKey("tabPage2");
            window.tabControl1.TabPages.RemoveByKey("tabPage3");
            window.ExportButton.Text = "Сохранить шаблон";
            window.toolStripTextBox2.Text = "<Укажите имя листа>";
            window.textBox3.Text = "Кнопка \"Cохранить шаблон\"  позволяет сохранить шаблон таблицы эксель для заполнения. Кнопка \"Загрузить\" загружает выбранную таблицу. Кнопка \"Записать\" записывает значения в параметры элементов. Плагин записывает значения параметров из первого столбца в параметры из второго столбца. Для сохранения корретных значений длин, объемов и площадей оба параметра должны иметь одинаковый тип данных. Плагин заполняет параметры для элементов на активном виде или спецификации";
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
                Dictionary<string, string> parametersDict = new Dictionary<string, string>();
                parametersDict.Clear();
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
                                for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)  // param value in row
                                {
                                    if (worksheet.Cells[i, 1].Value != null)
                                    {
                                        if (worksheet.Cells[i, 1].Value.ToString() != "")
                                        {
                                            try
                                            {
                                                parametersDict.Add($"{worksheet.Cells[i, 1].Value}", $"{worksheet.Cells[i, 2].Value}");
                                            }
                                            catch { }
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
                    window.toolStripProgressBar1.Maximum = (allElems.Count + platesInJoint.Count) * parametersDict.Count;
                    window.toolStripProgressBar1.Step = 1;
                    Dictionary<string, int> paramCount = new Dictionary<string, int>();
                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("Заполнение параметров семейств");
                        foreach (var el in allElems)
                        {
                            ObjRvt objRvt = new ObjRvt();
                            objRvt.elem = el;
                            window.toolStripProgressBar1.PerformStep();
                            foreach (var param in parametersDict)
                            {
                                try
                                {
                                    objRvt.SetParam(param.Value, objRvt.GetParam(param.Key));
                                }
                                catch { }
                            }
                        }

                        foreach (var pij in platesInJoint)
                        {
                            window.toolStripProgressBar1.PerformStep();
                            foreach (var param in parametersDict)
                            {
                                try
                                {
                                    pij.SetParam(param.Value, pij.GetParam(param.Key));
                                }
                                catch { }
                            }
                        }
                        window.toolStripProgressBar1.Value = (allElems.Count + platesInJoint.Count) * parametersDict.Count;
                        tx.Commit();
                    }
                }
            }
            return Result.Succeeded;
        }
    }
}
