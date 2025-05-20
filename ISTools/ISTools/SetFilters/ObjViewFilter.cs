using Autodesk.Revit.DB;
using System.ComponentModel;

namespace ISTools
{
    public class ObjViewFilter : INotifyPropertyChanged
    {
        private bool _isSelected;
        public string Name { get; set; }
        public ElementId Id { get; set; }
        public FilterElement Element { get; set; }
        public OverrideGraphicSettings OverrideGraphicSettings { get; set; }
        public bool Visability { get; set; }
        public bool Enabled { get; set; }

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

        public void SetFilterToView(View view)
        {
            view.AddFilter(Id);
            view.SetFilterOverrides(Id, OverrideGraphicSettings);
            view.SetFilterVisibility(Id, Visability);
            view.SetIsFilterEnabled(Id, Enabled);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
