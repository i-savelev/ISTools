using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;

namespace ISTools
{
    public class SchedulesTableModel : INotifyPropertyChanged
    {
        public Document Document;
        public List<ObjSheet> objSheetList { get; set; }
        public ObservableCollection<ObjTreeViewItemViewModel> TreeViewItemViewModelList { get; set; }

        public SchedulesTableModel(Document doc)
        {
            Document = doc;
            objSheetList = GetSheets(doc);
            TreeViewItemViewModelList = BuildTreeView(objSheetList);
        }

        public event PropertyChangedEventHandler PropertyChanged;


        private List<ObjSheet> GetSheets(Document doc)
        {
            List<ScheduleSheetInstance> schedules_in_sheet = new FilteredElementCollector(Document)
                .OfClass(typeof(ScheduleSheetInstance))
                .Cast<ScheduleSheetInstance>()
                .ToList();

            List<ObjSheet> objSheets = new List<ObjSheet>();
            var shets = new FilteredElementCollector(doc).
                OfCategory(BuiltInCategory.OST_Sheets).
                WhereElementIsNotElementType().
                Cast<ViewSheet>().
                ToList();

            BrowserOrganization org = BrowserOrganization.GetCurrentBrowserOrganizationForSheets(doc);
            foreach (var sheet in shets)
            {
                ObjSheet sheet_obj = new ObjSheet(sheet.Name, sheet.SheetNumber, sheet);
                sheet_obj.GetSchedules(schedules_in_sheet);

                List<FolderItemInfo> folderfields = org.GetFolderItems(sheet.Id).ToList();

                foreach (FolderItemInfo info in folderfields)
                {
                    string groupheader = info.Name;
                    sheet_obj.GroupList.Add(groupheader);
                }
                objSheets.Add(sheet_obj);
            }
            return objSheets;
        }


        public ObservableCollection<ObjTreeViewItemViewModel> BuildTreeView(List<ObjSheet> objSheetList)
        {
            var rootItems = new ObservableCollection<ObjTreeViewItemViewModel>();
            objSheetList.Sort(new NaturalComparer<ObjSheet>(s => s.Number));
            foreach (var sheet in objSheetList)
            {
                var sheetNode = new ObjTreeViewItemViewModel
                {
                    Header = $"{sheet.Number}: {sheet.Name}",
                    Sheet = sheet,
                    IsSelected = sheet.IsSelected
                };


                // Добавляем спецификации как дочерние узлы
                foreach (var schedule in sheet.Schedules)
                {
                    sheetNode.Children.Add(new ObjTreeViewItemViewModel
                    {
                        Header = $"📋 {schedule.Name}",
                        Schedule = schedule,
                        IsSelected = false, // Можно привязать к чему-то ещё
                        Parent = sheetNode
                    });

                }

                if (!sheet.GroupList.Any())
                {
                    rootItems.Add(sheetNode);
                    continue;
                }

                var currentLevel = rootItems;

                for (int i = 0; i < sheet.GroupList.Count; i++)
                {
                    string groupName = sheet.GroupList[i];

                    var existingNode = currentLevel.FirstOrDefault(x => x.Header == groupName);

                    if (existingNode == null)
                    {
                        existingNode = new ObjTreeViewItemViewModel { Header = groupName };
                        currentLevel.Add(existingNode);
                    }

                    if (i == sheet.GroupList.Count - 1)
                    {
                        existingNode.Children.Add(sheetNode);
                        sheetNode.Parent = existingNode;
                    }

                    currentLevel = existingNode.Children;

                }
            }

            return rootItems;
        }
    }
}
