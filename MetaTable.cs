using System.Globalization;
using OfficeOpenXml;

public class MetaTable
{
    public ExcelWorksheet s;
    public List<Column> columns = new();
    public int rows;

    public MetaTable(ExcelWorksheet s)
    {
        this.s = s;

        Console.WriteLine("Чтение таблицы...");

        int i = -1;
        while (true)
        {
            i++;
            if (i >= 1_000) throw new("Too many columns, seems something went wrong");

            var cell = s.At(i, 0);
            if (cell.Value == null) break;

            Column c = new()
            {
                name = cell.Value.ToString()!.Trim(),
                x = i,
                alphabetName = cell.ToString()[..1]
            };
            columns.Add(c);
            // Console.WriteLine($"- [{c.x}] {c.name}");
        }

        int rowsByMetaInfo = s.Dimension.End.Row;

        int rowsByActualData = rowsByMetaInfo;
        while (rowsByActualData > 0 && string.IsNullOrEmpty(s.At(0, rowsByActualData).Text))
        {
            rowsByActualData--;
        }

        if (rowsByMetaInfo != rowsByActualData)
        {
            rows = rowsByActualData - 1; // Exclude header row
            Console.WriteLine($"Размер таблицы: {columns.Count}x{rows} (Метаданные искажены: {columns.Count}x{rowsByMetaInfo})");
        }
        else
        {
            rows = rowsByActualData - 1; // Exclude header row
            Console.WriteLine($"Размер таблицы: {columns.Count}x{rows}");
        }
    }

    public Column GetColumn(string name)
    {
        Column? column = columns.FirstOrDefault(c => c.name == name);
        if (column == null) throw new($"Column '{name}' not found");
        return column;
    }

    public float AtFloat(Column column, int row)
    {
        object value = At<string>(column, row);
        string str = value.ToString().Replace(" ", "");
        float flt = float.Parse(str, NumberStyles.Any, CultureInfo.InvariantCulture);

        // Console.WriteLine($"'{str}' => {flt}");
        return flt;
    }

    public T At<T>(Column column, int row)
    {
        return (T)At(column, row);
    }
    public object At(Column column, int row)
    {
        return s.At(column.x, 1 + row).Value; // +1 due to header
    }
}