using System.Collections.Generic;
using System;
using System.Linq;
using System.ComponentModel;

public static class WPFHelper
{
    public static IEnumerable<T> ApplyFilter<T>(
        object sender,
        PropertyChangedEventHandler PropertyChanged,
        IEnumerable<T> sourceList,
        IEnumerable<T> filteredList,
        string searchText,
        Func<T, string> textSelector)
    {
        if (string.IsNullOrWhiteSpace(searchText))
        {
            return sourceList;
        }

        var lowerSearch = searchText.ToLower();

        var newfilteredList = sourceList.Where(item =>
        {
            var text = textSelector(item) ?? string.Empty;
            return text.ToLower().Contains(lowerSearch);
        });
        OnPropertyChanged(sender, PropertyChanged, filteredList);
        return newfilteredList;
    }

    public static void OnPropertyChanged<T>(object sender, PropertyChangedEventHandler PropertyChanged, IEnumerable<T> source)
    {
        PropertyChanged?.Invoke(sender, new PropertyChangedEventArgs(nameof(source)));
    }
}