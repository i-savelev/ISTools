using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace ISTools
{
    public class ObjSheet : INotifyPropertyChanged
    {
        public string Name { get; set; } = null;

        public string Number { get; set; } = null;

        public List<string> GroupList = new List<string>();

        public ViewSheet Elem { get; set; }

        public List<ObjSchedule> Schedules { get; set; } = new List<ObjSchedule>();


        private bool _isSelected;
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

        public ObjSheet(string name, string number, ViewSheet elem)
        {
            Name = name;
            Number = number;
            Elem = elem;
        }

        public ObjSheet()
        {

        }

        public void GetSchedules(List<ScheduleSheetInstance> scheduleSheetInstanceList)
        {
            foreach (var ssi in scheduleSheetInstanceList)
            {
                if (ssi.OwnerViewId == Elem.Id)
                {
                    var schedule = Elem.Document.GetElement(ssi.ScheduleId) as ViewSchedule;
                    if (schedule != null)
                    {
                        Schedules.Add(new ObjSchedule
                        {
                            Name = schedule.Name,
                            Id = ssi.Id,
                            Elem = ssi
                        });
                    }
                }
            }
            Schedules.Sort(new NaturalComparer<ObjSchedule>(s => s.Name));
        }
    }
}
