using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ISTools
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class FamilyCatalog : IExternalCommand
    {
        //---PluginsManager---//
        public static string IS_TAB_NAME => "ISTools";
        public static string IS_NAME => "Каталог семейств";
        public static string IS_IMAGE => "ISTools.Resources.FamilyCatalog32.png";
        public static string IS_DESCRIPTION => "Данная команда имеет три функции:\n- создание каталога семейств и типоразмеров системных элементов модели\n- дополнение существующего каталога\n- запись значений параметров из каталога в параметры элементов модели\nПодробнее: https://github.com/i-savelev/ISTools/wiki/Каталог-семейств";
        //---PluginsManager---//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            string docName = doc.Title.Replace(".rvt", "");
            List<Type> types = new List<Type>() {
                typeof(Family),
                typeof(FloorType),
                typeof(WallFoundationType),
                typeof(WallType),
                typeof(RoofType),
                typeof(RebarBarType),
                typeof(FabricSheetType),
                typeof(FabricWireType),
                typeof(FlexDuctType),
                typeof(FlexPipeType),
                typeof(MullionType),
                typeof(PanelType),
                typeof(PipeType),
                typeof(DuctType),
                typeof(RailingType),
                typeof(StairsLandingType),
                typeof(StairsRunType),
                typeof(StairsType),
                typeof(WallFoundationType),
                typeof(PipeInsulationType),
                typeof(DuctInsulationType),
                typeof(CableTrayType),
                typeof(ConduitType),
                typeof(ConduitType),
                typeof(MechanicalSystemType),
                typeof(PipingSystemType),
                typeof(HostedSweep),
                typeof(Stairs),
            };

            DataTable dataTableElements = new DataTable();
            dataTableElements.Columns.Add("Категория", typeof(string));
            dataTableElements.Columns.Add("Семейство", typeof(string));
            List<Element> elems = new List<Element>();
            List<string[]> dataRows = new List<string[]>();

            CollectTypesOfElements(elems, types);

            IsToolsForm window = new IsToolsForm();
            window.dataGridView1.DataSource = dataTableElements;
            window.Text = "Обработка данных семейств и параметров модели";
            window.dataGridView1.Refresh();
            window.ExportButton.Text = "Сохранить таблицу";
            window.ExportButton.Click += (s, e) => { TableExport(); };
            window.toolStripButton1.Click += (s, e) => { TableOverride(); };
            window.toolStripButton3.Click += (s, e) => { PapamPreview(); };
            window.toolStripButton1.Text = "Выбрать таблицу";
            window.toolStripButton2.Enabled = false;
            window.toolStripButton4.Enabled = false;
            window.toolStripTextBox1.Text = "<Укажите имя листа>";
            window.toolStripTextBox2.Text = "<Укажите имя листа>";
            window.toolStripTextBox3.Text = "<Укажите имя листа>";
            window.toolStripButton3.Text = "Загрузить данные";
            window.toolStripButton4.Text = "Записать данные";
            window.groupBox1.Text = "Сохранение новой таблицы";
            window.groupBox2.Text = "Список семейтсв, которых нет в указанной таблице excel";
            window.groupBox3.Text = "Запись данных из таблицы в параметры элеменов";
            window.toolStripButton2.Text = "Сохранить новые данные в таблицу";
            window.tabPage1.Text = "Сохранение новой таблицы";
            window.tabPage2.Text = "Дополнение существующей таблицы";
            window.tabPage3.Text = "Запись параметров";
            window.textBox1.Text = "Для записи значений параметров необходимо указать имя листа и выбрать файл excel. Список параметров и их значений отобразиться в таблице ниже. Параметры записываются для всех элементов на активном виде. Имена параметров - столбцы таблицы, наичиная с третьего. Для того, чтобы к значениям параметров добавился счетчик (001, 002, 003, ...), нужно добавить к названию столбца с названием параметра знак '*'";
            window.textBox2.Text = "Перед открытием существующего каталога, необходимо указать имя листа и выбрать файл excel. Элементы, которых нет в указанном каталоге, но есть в диспетчере модели, отобразятся в таблице ниже. Эти элементы можно сохранить в указанный каталог - они будут записаны в конец списка";
            window.textBox3.Text = "Ниже отобразится список всех элементов в диспетчере модели. Их можно сохранить в файл excel. При сохранеии можно указать имя листа";
            window.toolStrip1.Items.RemoveByKey("toolStripTextBox4");

            window.ShowDialog();

            void CollectTypesOfElements(List<Element> elemsList, List<Type> classes)
            {
                var familyList = new FilteredElementCollector(doc).
                    WhereElementIsNotElementType().
                    OfClass(typeof(Family)).
                    Cast<Family>().
                    ToList();

                foreach (Family family in familyList)
                {
                    elemsList.Add(family);
                    var placementType = family.FamilyPlacementType;
                    var familyCategory = family.FamilyCategory.Name;
                    var familyBuiltIn = ((BuiltInCategory)family.FamilyCategory.Id.IntegerValue).ToString();

                    if (placementType.ToString() != "ViewBased" & familyBuiltIn != "OST_DetailComponents" & familyBuiltIn != "OST_RebarShape" & familyBuiltIn != "OST_BoundaryConditions")
                    {
                        dataTableElements.Rows.Add(
                            familyCategory,
                            family.Name
                            );
                        string[] row = new string[] { familyCategory, family.Name };
                        dataRows.Add(row);
                    }
                }
                foreach (Type typ in classes)
                {
                    var typesList = new FilteredElementCollector(doc).
                        WhereElementIsElementType().
                        OfClass(typ).
                        Cast<Element>().
                        ToList();

                    foreach (Element type in typesList)
                    {
                        elemsList.Add(type);
                        dataTableElements.Rows.Add(
                            type.Category.Name,
                            type.Name
                            );
                        string[] row = new string[] { type.Category.Name, type.Name };
                        dataRows.Add(row);
                    }
                }
                dataTableElements.Rows.Add(
                            "Пластины",
                            "Пластина"
                            );
            }

            void TableExport()
            {
                window.toolStripProgressBar1.Value = 0;
                window.toolStripProgressBar1.Maximum = 1;
                window.toolStripProgressBar1.Step = 1;
                ObjExcelTable table = new ObjExcelTable();
                table.filename = $"Каталог семейств_{docName}";
                table.dt = dataTableElements;
                table.worksheets = window.toolStripTextBox2.Text;
                table.TableExport();
                window.toolStripProgressBar1.Value = 1;
            }

            void TableOverride()
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                List<string> excelDataString = new List<string>();
                window.dataGridView2.Rows.Clear();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    FileInfo existingFile = new FileInfo(openFileDialog1.FileName);
                    using (ExcelPackage excelPackage = new ExcelPackage(existingFile))
                    {
                        ExcelPackage excel = new ExcelPackage(existingFile);
                        window.toolStripButton2.Click += (s, e) => {
                            IsUtils.Save_opened_excel(excel, window.toolStripProgressBar2);
                        };
                        IsUtils.AddColumnHeaderFromExcel(excelPackage,
                            window.toolStripTextBox1.Text,
                            window.dataGridView2, 1, 1, 2
                            );
                        var rowCount = IsUtils.GetRowExcelCount(excelPackage, window.toolStripTextBox1.Text);
                        var excelData = IsUtils.GetListOfStringFromExcel(excelPackage, window.toolStripTextBox1.Text);

                        foreach (ExcelWorksheet worksheet in excel.Workbook.Worksheets)
                        {
                            if (worksheet.Name == window.toolStripTextBox1.Text)
                            {
                                foreach (string[] str in dataRows)
                                {
                                    if (!excelData.Any(d => $"{d[0]}_{d[1]}" == $"{str[0]}_{str[1]}"))
                                    {
                                        var n = 1;
                                        rowCount += 1;
                                        window.dataGridView2.Rows.Add(str);
                                        foreach (string s in str)
                                        {
                                            IsDebugWindow.AddRow(s);
                                            worksheet.Cells[rowCount, n].Value = s;
                                            n += 1;
                                        }
                                    }
                                }
                            }
                        }
                        window.toolStripButton2.Enabled = true;
                    }
                }
            }

            void PapamPreview()
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                window.dataGridView3.Rows.Clear();
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    FileInfo existingFile = new FileInfo(openFileDialog.FileName);
                    using (ExcelPackage excelPackage = new ExcelPackage(existingFile))
                    {
                        //ExcelPackage excelPackage = new ExcelPackage(existingFile);

                        IsUtils.AddColumnHeaderFromExcel(excelPackage,
                            window.toolStripTextBox3.Text,
                            window.dataGridView3, 1
                            );

                        var parametersDict = CreateParametersDict(excelPackage, window.toolStripTextBox3.Text);
                        var rowList = IsUtils.GetListOfStringFromExcel(excelPackage, window.toolStripTextBox3.Text);

                        foreach (string[] str in rowList)
                        {
                            window.dataGridView3.Rows.Add(str);
                        }
                        window.toolStripButton4.Enabled = true;
                        window.toolStripButton4.Click += (s, e) => { SetParamToElements(parametersDict); };
                    }
                }
            }

            void SetParamToElements(Dictionary<string, Dictionary<string, string>> parametersDict)
            {
                var allElementsOnView = new FilteredElementCollector(doc, doc.ActiveView.Id).
                    WhereElementIsNotElementType().
                    Cast<Element>().
                    ToList();

                ObjPlateCollector platesCollector = new ObjPlateCollector();
                platesCollector.doc = doc;
                var platesInJoint = platesCollector.GetPlates();

                window.toolStripProgressBar3.Value = 0;
                window.toolStripProgressBar3.Maximum = allElementsOnView.Count;
                window.toolStripProgressBar3.Step = 1;

                Dictionary<string, int> paramValueCountDict = new Dictionary<string, int>();
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Заполнение параметров семейств");

                    foreach (var elem in allElementsOnView)
                    {
                        window.toolStripProgressBar3.PerformStep();
                        try
                        {
                            if (elem.GetType().Name == "FamilyInstance")
                            {
                                var fam = doc.GetElement(elem.GetTypeId()) as FamilySymbol;
                                var categoryNameKey = $"{elem.Category.Name}_{fam.FamilyName}";
                                SetParamToElement(parametersDict, elem, paramValueCountDict, categoryNameKey);
                            }
                            if (elem.GetType().Name == "Rebar")
                            {
                                var categoryNameKey = $"{elem.Category.Name}_{elem.Name.Split(':')[0].TrimEnd(' ')}";
                                SetParamToElement(parametersDict, elem, paramValueCountDict, categoryNameKey);
                            }
                            else if (((BuiltInCategory)elem.Category.Id.IntegerValue).ToString() == "OST_StructConnectionPlates")
                            {
                                var categoryNameKey = $"{"Пластины"}_{"Пластина"}";
                                SetParamToElement(parametersDict, elem, paramValueCountDict, categoryNameKey);
                            }
                            else
                            {
                                var category_name_key = $"{elem.Category.Name}_{elem.Name}";
                                SetParamToElement(parametersDict, elem, paramValueCountDict, category_name_key);
                            }
                        }
                        catch { }
                    }
                    foreach (var pij in platesInJoint)
                    {
                        var category_name_key = $"{"Пластины"}_{"Пластина"}";
                        SetParamToPlatesInJoint(parametersDict, pij, paramValueCountDict, category_name_key);
                    }
                    window.toolStripProgressBar3.Value = allElementsOnView.Count;
                    tx.Commit();
                }
            }
            return Result.Succeeded;
        }

        void SetParamToElement(Dictionary<string, Dictionary<string, string>> parametersDict,
            Element elem,
            Dictionary<string, int> paramValueCountDict,
            string categoryNameKey
            )
        {
            try
            {
                foreach (var param in parametersDict)
                {
                    var count = "";
                    if (param.Value.ContainsKey(categoryNameKey))
                    {
                        if (!paramValueCountDict.ContainsKey(param.Value[categoryNameKey]))
                        {
                            paramValueCountDict.Add(param.Value[categoryNameKey], 0);
                        }
                        paramValueCountDict[param.Value[categoryNameKey]] += 1;
                        if (param.Key.Contains("*"))
                        {
                            count = $"{paramValueCountDict[param.Value[categoryNameKey]]:000}";
                        }
                        try
                        {
                            elem.LookupParameter(param.Key.Replace("*", ""))
                                .Set($"{param.Value[categoryNameKey]}{count}");
                        }
                        catch { }
                    }
                }
            }
            catch { }
        }

        void SetParamToPlatesInJoint(Dictionary<string, Dictionary<string, string>> parametersDict,
            ObjPlateInJoint pij,
            Dictionary<string, int> paramValueCountDict,
            string categoryNameKey
            )
        {
            try
            {
                foreach (var param in parametersDict)
                {
                    var count = "";
                    try
                    {
                        if (!paramValueCountDict.ContainsKey(param.Value[$"{"Пластины"}_{"Пластина"}"]))
                        {
                            paramValueCountDict.Add(param.Value[$"{"Пластины"}_{"Пластина"}"], 0);
                        }
                        paramValueCountDict[param.Value[$"{"Пластины"}_{"Пластина"}"]] += 1;
                        if (param.Key.Contains("*"))
                        {
                            count = $"{paramValueCountDict[param.Value[$"{"Пластины"}_{"Пластина"}"]]:000}";
                        }
                        pij.SetParam(param.Key.Replace("*", ""), $"{param.Value[$"{"Пластины"}_{"Пластина"}"]}{count}");
                    }
                    catch { }
                }
            }
            catch { }
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
                                                parameterDict.Add($"{worksheet.Cells[row, 1].Value}_{worksheet.Cells[row, 2].Value}", $"{worksheet.Cells[row, col].Value}");
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