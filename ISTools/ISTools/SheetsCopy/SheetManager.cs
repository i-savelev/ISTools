using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ISTools.ISTools.SheetsCopy
{
    internal class SheetManager
    {
        public Document Doc { get; set; }
        public List<ObjSheet> ObjSheetList = new List<ObjSheet>();  
        public SheetManager(Document doc)
        {
            Doc = doc;
            CollectAllSheets(Doc);
        }
        public void  CollectAllSheets(Document doc)
        {
            var sheets = new FilteredElementCollector(doc).
                OfCategory(BuiltInCategory.OST_Sheets).
                WhereElementIsNotElementType().
                Cast<ViewSheet>().
                ToList();
            ObjSheetList.Clear();
            BrowserOrganization org = BrowserOrganization.GetCurrentBrowserOrganizationForSheets(doc);
            foreach (var sheet in sheets)
            {
                ObjSheet sheet_obj = new ObjSheet(sheet.Name, sheet.SheetNumber, sheet);

                List<FolderItemInfo> folderfields = org.GetFolderItems(sheet.Id).ToList();

                foreach (FolderItemInfo info in folderfields)
                {
                    string groupheader = info.Name;
                    sheet_obj.GroupList.Add(groupheader);
                }
                ObjSheetList.Add(sheet_obj);
            }
        }

        List<ObjSheet> CollectSelectedSheets(List<string> selectedNode)
        {
            List<ObjSheet> selectedObjSheetList = new List<ObjSheet>();
            foreach (var str in selectedNode)
            {
                var sheet = ObjSheetList.FirstOrDefault(p => p.Number == str.Split(':')[0]);
                selectedObjSheetList.Add(sheet);

            }
            return selectedObjSheetList;
        }

        public void CopySheets(List<string> selectedNode, SheetsCopyWindowManager window)
        {
            using (Transaction trans = new Transaction(Doc, "Копирование листа"))
            {
                var prefix = window.Window.textBox2.Text;
                window.Window.progressBar1.Value = 0;
                window.Window.progressBar1.Maximum = selectedNode.Count;
                window.Window.progressBar1.Step = 1;
                trans.Start();
                var sheetList = CollectSelectedSheets(selectedNode);
                foreach (var sheetId in sheetList)
                {
                    var newSheet = DuplicateSheet(Doc, sheetId.Elem.Id, prefix);
                    Element sourceSheetElement = Doc.GetElement(sheetId.Elem.Id);
                    ViewSheet sourceSheet = sourceSheetElement as ViewSheet;
                    DuplivcateAllViwesInSheet(Doc, sourceSheet, newSheet, prefix);
                    DuplivcateAllLegendInSheet(Doc, sourceSheet, newSheet);
                    DuplicateSchedule(Doc, sourceSheet, newSheet);
                    CopyTextNotes(Doc, sourceSheet, newSheet);
                    CopyGenericAnnotation(Doc, sourceSheet, newSheet);
                    window.Window.progressBar1.PerformStep();
                }
                trans.Commit();
                window.Window.progressBar1.Value = selectedNode.Count;
            }
            CollectAllSheets(Doc);
        }

        void DuplivcateAllViwesInSheet(Document doc, ViewSheet sourceSheet, ViewSheet newSheet, string prefix)
        {
            // Копируем виды
            ICollection<ElementId> viewIds = sourceSheet.GetAllPlacedViews();
            foreach (ElementId viewId in viewIds)
            {
                View originalView = doc.GetElement(viewId) as View;
                try
                {
                    if (originalView != null & originalView.ViewType != ViewType.Legend)
                    {
                        // Копируем вид
                        ElementId copiedViewId = originalView.Duplicate(ViewDuplicateOption.WithDetailing);
                        View copiedView = doc.GetElement(copiedViewId) as View;

                        copiedView.Name = prefix + originalView.Name;
                        if (copiedView != null)
                        {
                            // Получаем координаты размещения вида на исходном листе
                            UV viewPosition = GetViewPositionOnSheet(sourceSheet, viewId);

                            // Добавляем скопированный вид на новый лист
                            if (viewPosition != null)
                            {
                                var newViewport = Viewport.Create(doc, newSheet.Id, copiedViewId, new XYZ(viewPosition.U, viewPosition.V, 0));
                                CopyViewportType(doc, sourceSheet, viewId, newViewport);
                            }
                        }
                    }
                }
                catch { IsDebugWindow.AddRow(viewId.ToString()); }

            }
        }


        private void CopyViewportType(Document doc, ViewSheet sourceSheet, ElementId viewId, Viewport newViewport)
        {
            // Находим исходный видовой экран на исходном листе
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Viewports).OwnedByView(sourceSheet.Id);

            foreach (Element element in collector)
            {
                Viewport originalViewport = element as Viewport;
                if (originalViewport != null && originalViewport.ViewId == viewId)
                {
                    // Получаем тип видового экрана
                    Parameter originalViewportTypeParam = originalViewport.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM);
                    if (originalViewportTypeParam != null)
                    {
                        ElementId originalViewportTypeId = originalViewportTypeParam.AsElementId();
                        if (originalViewportTypeId != ElementId.InvalidElementId)
                        {
                            // Применяем тот же тип видового экрана к новому видовому экрану
                            newViewport.ChangeTypeId(originalViewportTypeId);
                            try
                            {
                                newViewport.LabelOffset = originalViewport.LabelOffset;
                            }
                            catch { }


                        }
                    }
                    break;
                }
            }
        }

        void DuplivcateAllLegendInSheet(Document doc, ViewSheet sourceSheet, ViewSheet newSheet)
        {
            // Копируем виды
            ICollection<ElementId> viewIds = sourceSheet.GetAllPlacedViews();
            foreach (ElementId viewId in viewIds)
            {
                View originalView = doc.GetElement(viewId) as View;
                try
                {
                    if (originalView != null & originalView.ViewType == ViewType.Legend)
                    {
                        // Получаем координаты размещения вида на исходном листе
                        UV viewPosition = GetViewPositionOnSheet(sourceSheet, viewId);

                        // Добавляем скопированный вид на новый лист
                        if (viewPosition != null)
                        {
                            Viewport.Create(doc, newSheet.Id, viewId, new XYZ(viewPosition.U, viewPosition.V, 0));
                        }
                    }
                }
                catch { IsDebugWindow.AddRow(viewId.ToString()); }

            }
        }

        void DuplicateSchedule(Document doc, ViewSheet sourceSheet, ViewSheet newSheet)
        {
            FilteredElementCollector scheduleCollector = new FilteredElementCollector(doc);
            scheduleCollector.OfCategory(BuiltInCategory.OST_ScheduleGraphics).OwnedByView(sourceSheet.Id);

            foreach (Element scheduleElement in scheduleCollector)
            {
                ScheduleSheetInstance originalSchedule = scheduleElement as ScheduleSheetInstance;
                if (originalSchedule != null)
                {
                    // Получаем саму спецификацию
                    ElementId scheduleId = originalSchedule.ScheduleId;
                    ViewSchedule schedule = doc.GetElement(scheduleId) as ViewSchedule;

                    if (schedule != null & !schedule.IsTitleblockRevisionSchedule)
                    {
                        // Получаем позицию спецификации на исходном листе
                        XYZ schedulePosition = originalSchedule.Point;

                        // Создаем новую спецификацию на новом листе
                        ScheduleSheetInstance.Create(doc, newSheet.Id, scheduleId, schedulePosition);
                    }
                }
            }
        }

        private void CopyTextNotes(Document doc, ViewSheet sourceSheet, ViewSheet newSheet)
        {
            FilteredElementCollector textCollector = new FilteredElementCollector(doc, sourceSheet.Id);
            textCollector.OfCategory(BuiltInCategory.OST_TextNotes);

            foreach (TextNote textNote in textCollector)
            {
                XYZ position = textNote.Coord;
                string text = textNote.Text;
                TextNote.Create(doc, newSheet.Id, position, text, textNote.GetTypeId());
            }
        }

        private void CopyGenericAnnotation(Document doc, ViewSheet sourceSheet, ViewSheet newSheet)
        {
            var annotationCollector = new FilteredElementCollector(doc, sourceSheet.Id).OfCategory(BuiltInCategory.OST_GenericAnnotation).ToElementIds().ToList();
            if (annotationCollector.Count != 0 )
            {
                ElementTransformUtils.CopyElements(sourceSheet, annotationCollector, newSheet, null, null);
            }
            
        }

        ViewSheet DuplicateSheet(Document doc, ElementId sourceSheetId, string prefix)
        {
            var sourceTitleBlock = new FilteredElementCollector(doc, sourceSheetId).OfCategory(BuiltInCategory.OST_TitleBlocks).FirstElement();
            Element sourceSheetElement = doc.GetElement(sourceSheetId);
            ViewSheet sourceSheet = sourceSheetElement as ViewSheet;
            ViewSheet newSheet = ViewSheet.Create(doc, sourceTitleBlock.GetTypeId());
            newSheet.SheetNumber = prefix + sourceSheet.SheetNumber;
            newSheet.Name = sourceSheet.Name;
            var newTitleBlock = new FilteredElementCollector(doc, newSheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).FirstElement();
            CopyParameters(sourceTitleBlock, newTitleBlock, doc, new List<string>() { "Номер листа" });
            CopyParameters(sourceSheet, newSheet, doc, new List<string>() { "Номер листа" });
            return newSheet;
        }

        private UV GetViewPositionOnSheet(ViewSheet sheet, ElementId viewId)
        {
            FilteredElementCollector collector = new FilteredElementCollector(sheet.Document);
            collector.OfCategory(BuiltInCategory.OST_Viewports).OwnedByView(sheet.Id);

            foreach (Element viewport in collector)
            {
                Viewport vp = viewport as Viewport;
                if (vp != null && vp.ViewId == viewId)
                {
                    // Преобразуем XYZ в UV
                    XYZ center = vp.GetBoxCenter();
                    return new UV(center.X, center.Y);
                }
            }
            return null;
        }

        private void CopyParameters(Element sourceElement, Element targetElement, Document doc, List<string> excludeParams)
        {
            foreach (Parameter sourceParam in sourceElement.Parameters)
            {
                // Пропускаем параметры, которые нельзя редактировать
                if (!sourceParam.IsReadOnly)
                {
                    Parameter targetParam = targetElement.LookupParameter(sourceParam.Definition.Name);
                    if (targetParam != null & !targetParam.IsReadOnly & !excludeParams.Contains(sourceParam.Definition.Name))
                    {
                        try
                        {
                            switch (sourceParam.StorageType)
                            {
                                case StorageType.String:
                                    targetParam.Set(sourceParam.AsString());
                                    break;
                                case StorageType.Integer:
                                    targetParam.Set(sourceParam.AsInteger());
                                    break;
                                case StorageType.Double:
                                    targetParam.Set(sourceParam.AsDouble());
                                    break;
                                case StorageType.ElementId:
                                    targetParam.Set(sourceParam.AsElementId());
                                    break;
                                default:
                                    break;
                            }
                        }
                        catch { }
                    }
                }
            }
        }
    }
}
