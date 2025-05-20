using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

public class NaturalComparer<T> : IComparer<T>
{
    private readonly Func<T, string> _selector;

    public NaturalComparer(Func<T, string> selector)
    {
        _selector = selector ?? throw new ArgumentNullException(nameof(selector));
    }

    public int Compare(T x, T y)
    {
        if (x == null && y == null) return 0;
        if (x == null) return -1;
        if (y == null) return 1;

        string leftText = _selector(x) ?? "";
        string rightText = _selector(y) ?? "";

        // Извлекаем префикс и числовые части
        string leftPrefix = ExtractPrefix(leftText);
        string rightPrefix = ExtractPrefix(rightText);

        int prefixComparison = string.Compare(leftPrefix, rightPrefix, StringComparison.Ordinal);
        if (prefixComparison != 0)
            return prefixComparison;

        string[] leftParts = ExtractNumberParts(leftText);
        string[] rightParts = ExtractNumberParts(rightText);

        int maxLength = Math.Max(leftParts.Length, rightParts.Length);
        for (int i = 0; i < maxLength; i++)
        {
            int leftNumber = i < leftParts.Length && IsNumeric(leftParts[i]) ? int.Parse(leftParts[i]) : 0;
            int rightNumber = i < rightParts.Length && IsNumeric(rightParts[i]) ? int.Parse(rightParts[i]) : 0;

            if (leftNumber != rightNumber)
                return leftNumber - rightNumber;
        }

        return 0;
    }

    private string ExtractPrefix(string text)
    {
        return Regex.Replace(text, @"[\d\.]+", "").Trim();
    }

    private string[] ExtractNumberParts(string text)
    {
        string numericPart = Regex.Replace(text, @"[^\d\.]", "").Trim('.');
        return numericPart.Split('.');
    }

    private bool IsNumeric(string value)
    {
        return int.TryParse(value, out _);
    }
}