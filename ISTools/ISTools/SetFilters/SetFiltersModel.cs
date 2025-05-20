using System.Collections.ObjectModel;
using System.Windows.Input;
using Autodesk.Revit.DB;
using System.Linq;
using System.ComponentModel;
using System;
using View = Autodesk.Revit.DB.View;

namespace ISTools
{
    public class SetFiltersModel : INotifyPropertyChanged
    {
        public Document Document;
        public ObservableCollection<ObjView> Views { get; set; }
        public ObservableCollection<ObjView> ViewTemplates { get; set; }
        public ObservableCollection<ObjView> FilteredViewsWithFilters { get; private set; }
        public ObservableCollection<ObjView> FilteredViews { get; private set; }
        public ObservableCollection<ObjView> FilteredViewTemplates { get; private set; }
        public ICommand SetFilterToViewsCommand { get; }
        public ICommand SelectAllViewCommand { get; }
        public ICommand UnSelectAllViewCommand { get; }
        public ICommand SelectAllViewTemplateCommand { get; }
        public ICommand UnSelectAllViewTemplateCommand { get; }
        private string _searchTextFilters;
        private string _searchTextViews;
        private string _searchTextViewTemplates;

        public string SearchTextFilters
        {
            get => _searchTextFilters;
            set
            {
                if (_searchTextFilters != value)
                {
                    _searchTextFilters = value;
                    OnPropertyChanged(nameof(SearchTextFilters));
                    var filtered = WPFHelper.ApplyFilter(
                        this,
                        PropertyChanged,
                        Views, 
                        FilteredViewsWithFilters,
                        SearchTextFilters,
                        v => v.Name
                        );
                    FilteredViewsWithFilters.Clear();
                    foreach (var item in filtered)
                    {
                        FilteredViewsWithFilters.Add(item);
                    }
                }
            }
        }

        public string SearchTextViews
        {
            get => _searchTextViews;
            set
            {
                if (_searchTextViews != value)
                {
                    _searchTextViews = value;
                    OnPropertyChanged(nameof(SearchTextViews));
                    var filtered = WPFHelper.ApplyFilter(
                        this,
                        PropertyChanged,
                        Views,
                        FilteredViews,
                        SearchTextViews,
                        v => v.Name
                        );
                    FilteredViews.Clear();
                    foreach (var item in filtered)
                    {
                        FilteredViews.Add(item);
                    }
                }
            }
        }

        public string SearchTextViewTemplates
        {
            get => _searchTextViewTemplates;
            set
            {
                if (_searchTextViewTemplates != value)
                {
                    _searchTextViewTemplates = value;
                    OnPropertyChanged(nameof(SearchTextViewTemplates));
                    var filtered = WPFHelper.ApplyFilter(
                        this,
                        PropertyChanged,
                        ViewTemplates,
                        FilteredViewTemplates,
                        SearchTextViewTemplates,
                        v => v.Name
                        );
                    FilteredViewTemplates.Clear();
                    foreach (var item in filtered)
                    {
                        FilteredViewTemplates.Add(item);
                    }
                }
            }
        }

        public SetFiltersModel(Document doc)
        {
            Views = LoadViewsAndFilters(doc);
            Views.OrderBy(x => x.Name);
            ViewTemplates = LoadViewTemplates(doc);
            ViewTemplates.OrderBy(x => x.Name); 
            Document = doc; 
            SetFilterToViewsCommand = new RelayCommand(SetFilterToViews);
            SelectAllViewCommand = new RelayCommand(SelectAllView);
            UnSelectAllViewCommand = new RelayCommand(UnSelectAllView);
            SelectAllViewTemplateCommand = new RelayCommand(SelectAllViewTemplate);
            UnSelectAllViewTemplateCommand = new RelayCommand(UnSelectAllViewTemplate);
            FilteredViewsWithFilters = new ObservableCollection<ObjView>(Views);
            FilteredViews = new ObservableCollection<ObjView>(Views);
            FilteredViewTemplates = new ObservableCollection<ObjView>(ViewTemplates);
        }

        public event PropertyChangedEventHandler PropertyChanged;
        void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private ObservableCollection<ObjView> LoadViewsAndFilters(Document doc)
        {
            var views = new ObservableCollection<ObjView>();

            var viewsList = new FilteredElementCollector(doc).
                OfCategory(BuiltInCategory.OST_Views).
                Cast<View>().
                Where(v => v.IsTemplate == false).
                ToList();

            foreach (var view in viewsList)
            {
                ObjView objView = new ObjView(view);
                views.Add(objView);
            }

            return views;
        }

        private ObservableCollection<ObjView> LoadViewTemplates(Document doc)
        {
            var views = new ObservableCollection<ObjView>();
            var viewsList = new FilteredElementCollector(doc).
                OfCategory(BuiltInCategory.OST_Views).
                Cast<View>().
                Where(v => v.IsTemplate == true).
                ToList();

            foreach (var view in viewsList)
            {
                ObjView objView = new ObjView(view);
                views.Add(objView);
            }
            return views;
        }

        private void SelectAllView()
        {
            foreach (var view in FilteredViews)
            {
                view.IsSelected = true;
            }
        }
        private void UnSelectAllView()
        {
            foreach (var view in FilteredViews)
            {
                view.IsSelected = false;
            }
        }

        private void SelectAllViewTemplate()
        {
            foreach (var view in FilteredViewTemplates)
            {
                view.IsSelected = true;
            }
        }

        private void UnSelectAllViewTemplate()
        {
            foreach (var view in FilteredViewTemplates)
            {
                view.IsSelected = false;
            }
        }

        private void SetFilterToViews()
        {
            IsDebugWindow.DtSheets.Clear();
            var filters = new ObservableCollection<ObjViewFilter>();
            foreach (var view in Views)
            {
                if (view.Filters != null)
                {
                    var selectedFilters = view.Filters.Where(f => f.IsSelected).ToList();
                    foreach (var filter in selectedFilters)
                    {
                        filters.Add(filter);
                    }
                }
            }
            var selectedViews = Views.Where(f => f.IsSelected).ToList();
            var selectedTemplate = ViewTemplates.Where(f => f.IsSelected).ToList();
            using (Transaction tx = new Transaction(Document))
            {
                tx.Start("test");
                foreach (var filter in filters)
                {
                    foreach (var view in selectedViews)
                    {
                        try
                        {
                            if (view.View.GetFilters().Contains(filter.Id))
                            {
                                view.View.RemoveFilter(filter.Id);
                                filter.SetFilterToView(view.View);
                                IsDebugWindow.AddRow($"Фильтр \"{filter.Name}\" перезаписан для вида \"{view.Name}\"");
                            }
                            else
                            {
                                filter.SetFilterToView(view.View);
                                IsDebugWindow.AddRow($"Фильтр \"{filter.Name}\" добавлен для вида \"{view.Name}\"");
                            }
                        }
                        catch (Exception ex) { IsDebugWindow.AddRow($"Ошибка: [{ex.Message}] фильтр \"{filter.Name}\" не добавлен к виду \"{view.Name}\""); }
                    }
                    foreach (var template in selectedTemplate)
                    {
                        try
                        {
                            if (template.View.GetFilters().Contains(filter.Id))
                            {
                                template.View.RemoveFilter(filter.Id);
                                filter.SetFilterToView(template.View);
                                IsDebugWindow.AddRow($"Фильтр \"{filter.Name}\" перезаписан для шаблона \"{template.Name}\"");
                            }
                            else
                            {
                                filter.SetFilterToView(template.View);
                                IsDebugWindow.AddRow($"Фильтр \"{filter.Name}\" добавлен для шаблона \"{template.Name}\"");
                            }
                            
                        }
                        catch (Exception ex) { IsDebugWindow.AddRow($"Ошибка: [{ex.Message}] фильтр \"{filter.Name}\" не добавлен к шаблону \"{template.Name}\""); }
                    }
                }
                tx.Commit();
            }
            UpdateFilteredViews();
            IsDebugWindow.Show("Результат");
        }
        private void UpdateFilteredViews()
        {
            Views.Clear();
            Views = LoadViewsAndFilters(Document);
            FilteredViewsWithFilters = new ObservableCollection<ObjView>(Views);
            OnPropertyChanged(nameof(FilteredViewsWithFilters));
            FilteredViews = new ObservableCollection<ObjView>(Views);
            OnPropertyChanged(nameof(FilteredViews));
        }
    }
}
