using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;


namespace ISTools
{
    public class ObjTreeViewItemViewModel : INotifyPropertyChanged
    {
        public string Header { get; set; }

        public ObservableCollection<ObjTreeViewItemViewModel> Children { get; set; } =
            new ObservableCollection<ObjTreeViewItemViewModel>();

        public ObjTreeViewItemViewModel Parent { get; set; }

        private bool? _isSelected = false;
        public bool? IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value) return;

                _isSelected = value;

                // Обновляем дочерние элементы только если значение не null
                if (value.HasValue)
                {
                    foreach (var child in Children)
                    {
                        child.IsSelected = value;
                    }
                }

                // Уведомляем родителя о изменении
                OnPropertyChanged();

                // Если есть родитель — вызываем пересчёт его состояния
                Parent?.UpdateSelectionFromChildren();
            }
        }


        // Ссылка на оригинальный объект
        public ObjSheet Sheet { get; set; }
        public ObjSchedule Schedule { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        internal void UpdateSelectionFromChildren()
        {
            bool allSelected = Children.All(c => c.IsSelected == true);
            bool anySelected = Children.Any(c => c.IsSelected == true);

            if (allSelected)
                _isSelected = true;
            else if (!anySelected)
                _isSelected = false;
            else
                _isSelected = null; // Неопределённое состояние

            OnPropertyChanged(nameof(IsSelected));
        }
    }
}
