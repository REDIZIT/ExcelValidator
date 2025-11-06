using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

public static class Extensions
{
    public static ExcelRange At(this ExcelWorksheet s, int x_column, int y_row)
    {
        return s.Cells[y_row + 1, x_column + 1];
    }

    public static void Set(this ExcelWorksheet s, int x_column, int y_row, string value)
    {
        s.Cells[y_row + 1, x_column + 1].SetCellValue(0, 0, value);
    }

    public static string ToStrings<T>(this IEnumerable<T> ls, Func<T, string> selector)
    {
        return string.Join(", ", ls.Select(e => selector(e)));
    }
}