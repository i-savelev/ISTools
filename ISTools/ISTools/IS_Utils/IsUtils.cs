using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Windows;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using System.Xml.Serialization;


namespace ISTools
{
    public class IsUtils
    {
        /// <summary>
        /// Add one row to System.Windows.Forms.DataGridView
        /// </summary>
        /// <param name="dataGridView"></param>
        public static void DeleteRow(DataGridView dataGridView)
        {
            foreach (DataGridViewRow row in dataGridView.SelectedRows)
            {
                dataGridView.Rows.Remove(row);
            }
        }

        /// <summary>
        /// Add one row to System.Windows.Forms.DataGridView
        /// </summary>
        /// <param name="dataGridView"></param>
        public static void AddRow(DataGridView dataGridView)
        {
            dataGridView.Rows.Add();
        }

        /// <summary>
        /// Serialize object to xml file using the SaveFileDialog and XmlSerializer
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        public static void SaveXmlDialog<T>(T obj)
            where T : class
        {
            XmlSerializer serializer = new XmlSerializer(typeof(T));
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Выберете папку для сохранения отчета";
            saveFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
            saveFileDialog.DefaultExt = ".xml";
            saveFileDialog.FileName = $"Рабочие наборы.xml";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;
                string directory = saveFileDialog.InitialDirectory;
                using (StreamWriter streamwriter = new StreamWriter($"{directory}{filePath}"))
                {
                    serializer.Serialize(streamwriter, obj);
                }
            }
        }

        /// <summary>
        /// Serialize object to xml file using XmlSerializer
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <param name="path"></param>
        public static void SaveXml<T>(T obj, string path)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(T));
            string userName = Environment.UserName;
            string filePath = path;
            using (StreamWriter streamwriter = new StreamWriter(filePath))
            {
                serializer.Serialize(streamwriter, obj);
            }
        }

        /// <summary>
        /// Deserialize xml file to object using XmlSerializer
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="file_path"></param>
        /// <returns></returns>
        public static T LoadXml<T>(string file_path)
            where T : class
        {
            string userName = Environment.UserName;
            using (StreamReader streamReader = new StreamReader(file_path))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(T));
                T deserializedObject = (T)serializer.Deserialize(streamReader);
                return deserializedObject;
            }
        }

        /// <summary>
        /// Deserialize xml file to object using the OpenFileDialog and XmlSerializer
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static T LoadXmlDialog<T>()
            where T : class
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = openFileDialog.FileName;
                string fileDirectory = openFileDialog.InitialDirectory;
                var filePath = $"{fileDirectory}{fileName}";
                return LoadXml<T>(filePath);
            }
            else { return null; }
        }

        /// <summary>
        /// Return all CategoryType.Model name
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static List<string> GetCategoryModelList(Document doc)
        {
            List<string> categorieslist = new List<string>();
            var allCategories = doc.Settings.Categories;
            foreach (Category category in allCategories)
            {
                if (category.CategoryType == CategoryType.Model)
                {
                    categorieslist.Add(category.Name);
                }
            }
            categorieslist.Add("Оси");
            categorieslist.Add("Уровни");
            categorieslist.Add("Выступающие профили");
            categorieslist.Sort();
            return categorieslist;
        }

        /// <summary>
        /// Return random color
        /// </summary>
        /// <param name="rnd"></param>
        /// <returns></returns>
        public static System.Drawing.Color GetRandomSystemColor(Random rnd)
        {
            int R = rnd.Next(255);
            int G = rnd.Next(255);
            int B = rnd.Next(255);
            System.Drawing.Color system_color = System.Drawing.Color.FromArgb(R, G, B);

            return system_color;
        }

        /// <summary>
        /// Covert System.Drawing.Color to Autodesk.Revit.DB.Color
        /// </summary>
        /// <param name="color"></param>
        /// <returns></returns>
        public static Color ConvSysToRvtColor(System.Drawing.Color color)
        {
            int R = color.R;
            int G = color.G;
            int B = color.B;
            Color rvt_color = new Color((byte)R, (byte)G, (byte)B);

            return rvt_color;
        }

        /// <summary>
        /// Return int R, int G, int B list from System.Drawing.Color color
        /// </summary>
        /// <param name="color"></param>
        /// <returns></returns>
        public static List<int> GetColorListInt(System.Drawing.Color color)
        {
            var colorList = new List<int> { color.R, color.G, color.B };
            return colorList;
        }

        /// <summary>
        /// Add button to existing tab and existing paanel. 
        /// If the tab does not exist, it will be created.
        /// If the panel does not exist, it will be created.
        /// </summary>
        /// <param name="uiApp"></param>
        /// <param name="buttonName"></param>
        /// <param name="className"></param>
        /// <param name="tabName"></param>
        /// <param name="panelName"></param>
        /// <param name="largeImage"></param>
        /// <param name="smallImage"></param>
        /// <param name="tooltip"></param>
        /// <param name="contextualHelp"></param>
        public static void AddButtonToExistTab(
            UIApplication uiApp,
            string buttonName,
            string className,
            string tabName,
            string panelName,
            string largeImage,
            string smallImage,
            string tooltip = null,
            string contextualHelp = null
            )
        {
            var tabs = ComponentManager.Ribbon.Tabs;
            if (tabs.Any(tab => tab.Name == tabName))
            {
                if (uiApp.GetRibbonPanels(tabName).Any(panel => panel.Name == panelName))
                {
                    Autodesk.Revit.UI.RibbonPanel exist_panel = uiApp.GetRibbonPanels(tabName).FirstOrDefault(panel => panel.Name == panelName);
                    AddButtonToPanel(uiApp, buttonName, className, exist_panel, largeImage, smallImage, tooltip, contextualHelp);
                }
                else
                {
                    Autodesk.Revit.UI.RibbonPanel new_panel = uiApp.CreateRibbonPanel(tabName, panelName);
                    AddButtonToPanel(uiApp, buttonName, className, new_panel, largeImage, smallImage, tooltip, contextualHelp);
                }
            }
            else
            {
                uiApp.CreateRibbonTab(tabName);
                Autodesk.Revit.UI.RibbonPanel panel = uiApp.CreateRibbonPanel(tabName, panelName);
                AddButtonToPanel(uiApp, buttonName, className, panel, largeImage, smallImage, tooltip, contextualHelp);
            }
        }

        /// <summary>
        /// Create a button on the tab and panel
        /// </summary>
        /// <param name="uiApp"></param>
        /// <param name="buttonName"></param>
        /// <param name="className"></param>
        /// <param name="panel"></param>
        /// <param name="largeImage"></param>
        /// <param name="smallImage"></param>
        /// <param name="tooltip"></param>
        /// <param name="contextualHelp"></param>
        public static void AddButtonToPanel(
            UIApplication uiApp,
            string buttonName,
            string className,
            Autodesk.Revit.UI.RibbonPanel panel,
            string largeImage,
            string smallImage,
            string tooltip = null,
            string contextualHelp = null
            )
        {
            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData buttonData = new PushButtonData(buttonName, buttonName, assemblyPath, className);
            var button = panel.AddItem(buttonData) as PushButton;
            button.LargeImage = GetImageFromResources(largeImage);
            button.Image = GetImageFromResources(smallImage);
            if (tooltip != null) button.ToolTip = tooltip;
            if (contextualHelp != null) button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));
        }

        /// <summary>
        /// get bitmap image from resources
        /// </summary>
        /// <param name="resourceName"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static BitmapImage GetImageFromResources(string resourceName)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream($"{resourceName}"))
            {
                if (stream is null)
                {
                    throw new ArgumentException($"Ресурс '{resourceName}' не найден.");
                }

                var bitmap = new BitmapImage();

                bitmap.BeginInit();
                bitmap.StreamSource = stream;
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();

                return bitmap;
            }
        }

        /// <summary>
        /// Change select cells color in DataGridView
        /// </summary>
        /// <param name="dataGridView"></param>
        /// <param name="columnName"></param>
        public static void ChengeDatagridCellsColor(DataGridView dataGridView, string columnName)
        {
            ColorDialog MyDialog = new ColorDialog();
            if (MyDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (DataGridViewCell cell in dataGridView.SelectedCells)
                {
                    var colindex = dataGridView.Columns[columnName].Index;
                    var rowindex = cell.RowIndex;
                    dataGridView.Rows[rowindex].Cells[colindex].Style.BackColor = MyDialog.Color;
                    dataGridView.ClearSelection();
                    int R = MyDialog.Color.R;
                    int G = MyDialog.Color.G;
                    int B = MyDialog.Color.B;
                }
            }
        }

        /// <summary>
        /// Add column headers to DataGridView from excel table 
        /// </summary>
        /// <param name="exel_package"></param>
        /// <param name="worksheet_name"></param>
        /// <param name="data_grid"></param>
        /// <param name="row"></param>
        /// <param name="start_col"></param>
        /// <param name="end_col"></param>
        public static void AddColumnHeaderFromExcel(ExcelPackage exel_package, string worksheet_name, DataGridView data_grid, int row, int start_col = 0, int end_col = 0)
        {
            foreach (ExcelWorksheet worksheet in exel_package.Workbook.Worksheets)
            {
                if (worksheet.Name == worksheet_name)
                {
                    int start = 1;
                    int end = 1;

                    if (start_col != 0) start = start_col;
                    else if (start_col == 0) start = worksheet.Dimension.Start.Column;

                    if (end_col != 0) end = end_col;
                    else if (end_col == 0) end = worksheet.Dimension.End.Column;

                    for (int j = start; j <= end; j++)
                    {
                        data_grid.ColumnCount = end;
                        if (worksheet.Cells[row, j].Value != null)
                        {
                            data_grid.Columns[j - 1].Name = worksheet.Cells[row, j].Value.ToString();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Return count of rows from excel table
        /// </summary>
        /// <param name="exelPackage"></param>
        /// <param name="worksheetName"></param>
        /// <returns></returns>
        public static int GetRowExcelCount(ExcelPackage exelPackage, string worksheetName)
        {

            int count = 0;
            foreach (ExcelWorksheet worksheet in exelPackage.Workbook.Worksheets)
            {
                if (worksheet.Name == worksheetName)
                {
                    for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                    {
                        if (worksheet.Cells[i, 1].Value != null)
                        {
                            if (worksheet.Cells[i, 1].Value.ToString() != "") count += 1;
                        }
                    }
                }
            }
            return count;
        }

        /// <summary>
        /// Return all table rows from excel as List<string[]>
        /// </summary>
        /// <param name="exelPackage"></param>
        /// <param name="worksheetName"></param>
        /// <returns></returns>
        public static List<string[]> GetListOfStringFromExcel(ExcelPackage exelPackage, string worksheetName)
        {
            List<string[]> data = new List<string[]>();

            foreach (ExcelWorksheet worksheet in exelPackage.Workbook.Worksheets)
            {
                if (worksheet.Name == worksheetName)
                {
                    for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                    {
                        string[] row = new string[worksheet.Dimension.End.Column];
                        //loop all columns in a row
                        for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                        {
                            //add the cell data to the List
                            if (worksheet.Cells[i, j].Value != null)
                            {
                                if (worksheet.Cells[i, j].Value.ToString() != "")
                                {
                                    row[j - 1] = worksheet.Cells[i, j].Value.ToString();
                                }
                            }
                        }
                        data.Add(row);
                    }
                }
            }
            return data;
        }

        /// <summary>
        /// Save opened excel file
        /// </summary>
        /// <param name="exelPackage"></param>
        /// <param name="progressBar"></param>
        public static void Save_opened_excel(ExcelPackage exelPackage, ToolStripProgressBar progressBar)
        {
            progressBar.Value = 0;
            progressBar.Maximum = 1;
            progressBar.Step = 1;
            exelPackage.Save();
            progressBar.Value = 1;
        }

        /// <summary>
        /// Set list of string string[] in to excel table
        /// </summary>
        /// <param name="dataRows"></param>
        /// <param name="exelPackage"></param>
        /// <param name="worksheetName"></param>
        /// <param name="startRow"></param>
        public static void SetStringForExcel(List<string[]> dataRows, ExcelPackage exelPackage, string worksheetName, int startRow)
        {
            foreach (ExcelWorksheet worksheet in exelPackage.Workbook.Worksheets)
            {
                if (worksheet.Name == worksheetName)
                {
                    foreach (string[] str in dataRows)
                    {
                        var n = 1;
                        startRow += 1;
                        foreach (string s in str)
                        {
                            worksheet.Cells[startRow, n].Value = s;
                            n += 1;
                        }
                    }
                }
            }
        }
    }
}
