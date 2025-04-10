using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Diagnostics;
using Autodesk.Revit.DB.Architecture;
using System.Linq;
using System.Data;
using System.Collections.Generic;
using System.Windows.Forms;
using Document = Autodesk.Revit.DB.Document;
using OfficeOpenXml;
using System.IO;


namespace ISTools
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ParamFromRoom : IExternalCommand
    {
        //---PluginsManager---//
        public static string IS_TAB_NAME => "ISTools";
        public static string IS_NAME => "Параметры из помещений";
        public static string IS_IMAGE => "ISTools.Resources.Rooms32.png";
        public static string IS_DESCRIPTION => "Команда позволяет передавать значения параметров из помещений в элементы модели";
        //---PluginsManager---//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            var app = doc.Application;

            DataTable dtParams = new DataTable();
            dtParams.Columns.Add("Параметр помещения", typeof(string));
            dtParams.Columns.Add("Параметр элемента", typeof(string));

            dtParams.Rows.Add(
                "Парметр1",
                "Параметр2"
                );

            var elementLinks = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_RvtLinks)
                .WhereElementIsNotElementType()
                .ToElements()
                .Cast<RevitLinkInstance>()
                .ToList();

            List<ObjRvtLink> objRvtLinks = new List<ObjRvtLink>();
            List<string> linkList = new List<string>() {doc.Title};
            Dictionary<string, string> parametersDict = new Dictionary<string, string>();

            foreach (Document linkDoc in app.Documents)
            {
                ObjRvtLink linkObj = new ObjRvtLink();
                linkObj.currentDoc = doc;
                linkObj.linkName = linkDoc.Title;
                objRvtLinks.Add(linkObj);
                if (linkDoc.IsLinked)
                {
                    linkList.Add(linkDoc.Title);
                }
            }
            ParamFromRoomForm window = new ParamFromRoomForm();
            window.checkedListBox1.DataSource = linkList;
            window.Text = "Работа с помещениями";
            window.groupBox1.Text = "Выбор моделей";
            window.groupBox2.Text = "Настройка";
            window.groupBox3.Text = "Передача параметров элементам";
            window.groupBox4.Text = "Сздание и удаление геометрии помещений";
            window.groupBox5.Text = "Информация";
            window.toolStripTextBox1.Text = "<имя листа>";
            window.button4.Text = "Сохранить шаблон таблицы";
            window.toolStripButton2.Text = "Загрузить параметры";
            window.button1.Text = "Заполнить параметры";
            window.button2.Text = "Создать геометрию помещений";
            window.button3.Text = "Удалить геометрию помещений";
            window.textBox2.Text = "Укажите размер смещения контура и основания помещения. Данный параметр необходим, чтобы геометрия помещения пересеклась с стенами и перекрытием, которые окружают помещение";
            window.textBox3.Text = "Укажите толщину перекрытия. Данный параметр необходим, чтобы верхняя грань геометии помещения не пересекалась с перекрытием над ним";
            window.textBox1.Text = "";
            window.textBox4.Text = "Для заполнения параметров и формирования геометрии помещений необходим выбрать файлы, помещения из которых должны быть обработаны. Для заполнения параметров необходимо выбрать таблицу сопоставления параметров. Параметры будут заполнны только тем элеменам, которые отображены на активном виде или в спецификации. Для элементов, которые пересекают несколько помещений будут записаны параметры из всех элементов через <разделитель>";
            window.button1.Enabled = false;
            window.button4.Click += (s, e) => { TableExport(); };
            window.toolStripButton2.Click += (s, e) => { DownloadMap(); };
            window.button1.Click += (s, e) => { ParamFill(); };
            window.button3.Click += (s, e) => { DeleteRoomSolid(); };
            window.button2.Click += (s, e) => { SetRoomSolid(); };
            window.toolStripTextBox2.Text = "<разделитель>";
            window.textBox5.Text = "мм";
            window.textBox6.Text = "мм";
            window.ShowDialog();
            void TableExport()
            {
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    ObjExcelTable table = new ObjExcelTable();
                    table.filename = "Шаблон";
                    table.dt = dtParams;
                    table.worksheets = window.toolStripTextBox1.Text;
                    table.TableExport();
                }
            }

            void DownloadMap()
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                window.dataGridView1.Rows.Clear();
                parametersDict.Clear();
                window.dataGridView1.Rows.Clear();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    FileInfo existingFile = new FileInfo(openFileDialog1.FileName);
                    using (ExcelPackage excelPackage = new ExcelPackage(existingFile))
                    {
                        ExcelPackage excel = new ExcelPackage(existingFile);
                        foreach (ExcelWorksheet worksheet in excel.Workbook.Worksheets)
                        {
                            if (worksheet.Name == window.toolStripTextBox1.Text)
                            {
                                IsUtils.AddColumnHeaderFromExcel(excel,
                                    window.toolStripTextBox1.Text,
                                    window.dataGridView1, 1
                                    );

                                for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                                {
                                    if (worksheet.Cells[i, 1].Value != null &
                                        worksheet.Cells[i, 1].Value.ToString() != "" &
                                        worksheet.Cells[i, 2].Value != null &
                                        worksheet.Cells[i, 2].Value.ToString() != "")
                                    {
                                        parametersDict.Add(worksheet.Cells[i, 1].Value.ToString(), worksheet.Cells[i, 2].Value.ToString());
                                    }
                                }
                                var rowList = IsUtils.GetListOfStringFromExcel(excel, window.toolStripTextBox1.Text);
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
                        window.button1.Enabled = true;
                    }
                }
            }

            void DeleteRoomSolid()
            {
                var roomSolids = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .WhereElementIsNotElementType()
                    .Cast<Element>()
                    .ToList();

                window.progressBar1.Value = 0;
                window.progressBar1.Maximum = roomSolids.Count;
                window.progressBar1.Step = 1;
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Удаление помещений");
                    foreach (var roomSolid in roomSolids)
                    {
                        window.progressBar1.PerformStep();
                        try
                        {
                            if (roomSolid.get_Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString().Contains("##room_"))
                            {
                                doc.Delete(roomSolid.Id);
                            }
                        }
                        catch { }
                    }
                    window.progressBar1.Value = roomSolids.Count;
                    tx.Commit();
                }
            }

            void SetRoomSolid()
            {
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                List<string> outlist = new List<string>();   
                var offset = (double)window.numericUpDown1.Value / 304.8;
                var thickness = (double)window.numericUpDown2.Value / 304.8;
                List<string> selectLinkList = new List<string>();
                foreach (var i in window.checkedListBox1.CheckedItems)
                {
                    selectLinkList.Add((string)i);
                };

                List<ObjRoom> rooms = new List<ObjRoom>();
                foreach (var link in elementLinks)
                {
                    if (link.GetLinkDocument() != null)
                    {
                        if ((String.Join("", selectLinkList).Contains(link.GetLinkDocument().Title)))
                        {
                            FilteredElementCollector spatialElem = new FilteredElementCollector(link.GetLinkDocument()).OfClass(typeof(SpatialElement));
                            foreach (SpatialElement e in spatialElem)
                            {
                                Room room = e as Room;

                                if (null != room)
                                {
                                    if (room.Area != 0)
                                    {
                                        ObjRoom objroom = new ObjRoom();
                                        objroom.transform = link.GetTransform();
                                        objroom.roomDoc = link.GetLinkDocument();
                                        objroom.doc = doc;
                                        objroom.room = room;
                                        rooms.Add(objroom);
                                    }
                                }
                            }
                        }
                    }
                }
                if ((String.Join("", selectLinkList).Contains(doc.Title)))
                {
                    FilteredElementCollector spatialElem = new FilteredElementCollector(doc).OfClass(typeof(SpatialElement));
                    foreach (SpatialElement e in spatialElem)
                    {
                        Room room = e as Room;
                        if (null != room)
                        {
                            if (room.Area != 0)
                            {
                                var translation = Transform.CreateTranslation(new XYZ(0, 0, 0));
                                ObjRoom objroom = new ObjRoom();
                                objroom.transform = translation;
                                objroom.doc = doc;
                                objroom.roomDoc = doc;
                                objroom.room = room;
                                rooms.Add(objroom);
                            }
                        }
                    }
                }
                window.progressBar1.Value = 0;
                window.progressBar1.Maximum = rooms.Count;
                window.progressBar1.Step = 1;
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Параметры из помещений");
                    foreach (ObjRoom r in rooms)
                    {
                        window.progressBar1.PerformStep();
                        try
                        {
                            r.SetRoomSolid(offset, thickness);
                        }
                        catch
                        {
                            outlist.Add($"{r.room.Name}_Id {r.room.Id}");
                        }
                    }
                    tx.Commit();
                    TaskDialog.Show("Список необработанных помещений", string.Join("\n", outlist));
                    window.progressBar1.Value = rooms.Count;
                    stopwatch.Stop();
                    TimeSpan elapsedTime = stopwatch.Elapsed;
                    window.textBox1.Text = $"Геометрия создана. Время выполнения: {Math.Round(elapsedTime.TotalSeconds, 3).ToString()} сек. ";
                }
            }

            void ParamFill()
            {
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                List<string> outlist = new List<string>();
                List<string> selectLinkList = new List<string>();
                foreach (var i in window.checkedListBox1.CheckedItems)
                {
                    selectLinkList.Add((string)i);
                };

                List<ObjRoom> rooms = new List<ObjRoom>();
                foreach (var link in elementLinks)
                {
                    if (link.GetLinkDocument() != null)
                    {
                        if ((String.Join("", selectLinkList).Contains(link.GetLinkDocument().Title)))
                        {
                            FilteredElementCollector spatialElem = new FilteredElementCollector(link.GetLinkDocument()).OfClass(typeof(SpatialElement));
                            foreach (SpatialElement e in spatialElem)
                            {
                                Room room = e as Room;

                                if (null != room)
                                {
                                    if (room.Area != 0)
                                    {
                                        ObjRoom objroom = new ObjRoom();
                                        objroom.transform = link.GetTransform();
                                        objroom.roomDoc = link.GetLinkDocument();
                                        objroom.doc = doc;
                                        objroom.roomDoc = doc;
                                        objroom.room = room;
                                        rooms.Add(objroom);
                                    }
                                }
                            }
                        }
                    }   
                }
                if ((String.Join("", selectLinkList).Contains(doc.Title)))
                {
                    FilteredElementCollector spatialElem = new FilteredElementCollector(doc).OfClass(typeof(SpatialElement));

                    foreach (SpatialElement e in spatialElem)
                    {
                        Room room = e as Room;

                        if (null != room)
                        {
                            if (room.Area != 0)
                            {
                                var translation = Transform.CreateTranslation(new XYZ(0, 0, 0));
                                ObjRoom objroom = new ObjRoom();
                                objroom.transform = translation;
                                objroom.doc = doc;
                                objroom.roomDoc = doc;
                                objroom.room = room;
                                rooms.Add(objroom);
                            }
                        }
                    }
                }
                window.progressBar1.Value = 0;
                window.progressBar1.Maximum = rooms.Count;
                window.progressBar1.Step = 1;
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Параметры из помещений");
                    foreach (var param in parametersDict)
                    {
                        Dictionary<ElementId, List<string>> map = new Dictionary<ElementId, List<string>>();
                        
                        foreach (ObjRoom r in rooms)
                        {
                            window.progressBar1.PerformStep();
                            try
                            {
                                var offset = (double)window.numericUpDown1.Value / 304.8;
                                var thickness = (double)window.numericUpDown2.Value / 304.8;
                                var roomsolid = r.GetRoomSolid(offset, thickness);
                                var bbox = roomsolid.GetBoundingBox();
                                var t = bbox.Transform;

                                Outline outline = new Outline(bbox.Min + t.Origin, bbox.Max + t.Origin);
                                ElementIntersectsSolidFilter solidfFilter = new ElementIntersectsSolidFilter(roomsolid);
                                BoundingBoxIntersectsFilter bbfilter = new BoundingBoxIntersectsFilter(outline);

                                var allElems = new FilteredElementCollector(doc, doc.ActiveView.Id).
                                   WhereElementIsNotElementType().
                                   WherePasses(bbfilter).
                                   WherePasses(solidfFilter).
                                   Cast<Element>().
                                   ToList();

                                foreach (Element element in allElems)
                                {
                                    try
                                    {
                                        ObjRvt obj = new ObjRvt();
                                        obj.elem = element;
                                        var p = r.room.LookupParameter(param.Key).AsValueString();
                                        if (!map.ContainsKey(element.Id)) map.Add(element.Id, new List<string>());
                                        if (!map[element.Id].Contains(p)) map[element.Id].Add(p);
                                        obj.SetParam(param.Value, string.Join(window.toolStripTextBox2.Text, map[element.Id]));
                                    }
                                    catch { }
                                }
                            }
                            catch
                            {
                                outlist.Add($"{r.room.Name}_Id {r.room.Id}");
                            }
                        }
                    }
                    window.progressBar1.Value = rooms.Count;
                    TaskDialog.Show("Список необработанных помещений", string.Join("\n", outlist));
                    tx.Commit();
                    stopwatch.Stop();
                    TimeSpan elapsedTime = stopwatch.Elapsed;
                    window.textBox1.Text = $"Параметры заполнены. Время выполнения: {Math.Round(elapsedTime.TotalSeconds, 3).ToString()} сек.";
                }
            }
            return Result.Succeeded;
        }
    }
}
