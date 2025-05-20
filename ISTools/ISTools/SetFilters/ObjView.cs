using Autodesk.Revit.DB;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Controls;

namespace ISTools
{
    public class ObjView : INotifyPropertyChanged
    {
        private bool _isSelected;
        public string Name { get; set; }
        public ElementId Id { get; set; }
        public View View { get; set; }
        public ObservableCollection<ObjViewFilter> Filters { get; set; }

        public ObjView(View view)
        {
            View = view;
            Name = view.Name;
            Filters = new ObservableCollection<ObjViewFilter>();
            foreach (ElementId filterId in view.GetFilters())
            {
                var filter = view.Document.GetElement(filterId) as FilterElement;
                if (filter != null)
                {
                    Filters.Add(new ObjViewFilter
                    {
                        Name = filter.Name,
                        Id = filter.Id,
                        Element = filter,
                        OverrideGraphicSettings = view.GetFilterOverrides(filterId),
                        Visability = view.GetFilterVisibility(filterId),
                        Enabled = view.GetIsFilterEnabled(filterId)
                    });
                }
            }
        }
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
