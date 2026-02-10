using System.Drawing;
using System.Text;
using OfficeOpenXml;

public static class Program
{
    public static void Main(string[] args)
    {
        try
        {
            UnsafeMain();
        }
        catch (Exception err)
        {
            Console.WriteLine();

            Console.BackgroundColor = ConsoleColor.Red;
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write(" Panic! ");
            Console.ResetColor();

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(" Program crashed with exception below:");
            Console.WriteLine(err.ToString());

            Console.Read();
        }
    }

    private static void UnsafeMain()
    {
        Console.WriteLine("Укажите путь до файла выгрузки (таблица.xlsx) (можно просто перетащить)");
        string path = Console.ReadLine()!.Replace("\"", "");

        if (File.Exists(path) == false) throw new($"File '{path}' does not exist");

        Console.WriteLine("Чтение файла...");

        ExcelPackage.License.SetNonCommercialPersonal("AAA");

        ExcelPackage package = new(path);
        var book = package.Workbook;
        var sheets = book.Worksheets;

        var s = sheets[0];
        MetaTable t = new(s);
        var tests = Rules.Check(t);

        ExportProblems(tests, t);

        Console.WriteLine(" - сохранение файла...");

        string exportPath = Path.GetDirectoryName(path) + "/" + Path.GetFileNameWithoutExtension(path) + "_export.xlsx";

        while (true)
        {
            try
            {
                package.SaveAs(exportPath);
                Console.WriteLine($"\nОтчёт сохранён по пути: '{exportPath}'");
                break;
            }
            catch (InvalidOperationException e)
            {
                if (e.InnerException.InnerException.GetType() == typeof(IOException))
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"Кажется, файл '{Path.GetFileName(exportPath)}' открыт в другой программе. Попытка сохранить файл через 5 секунд...");
                    Console.ResetColor();
                    Thread.Sleep(5000);
                }
                else
                {
                    throw;
                }
            }
            catch
            {
                throw;
            }
        }

        Console.WriteLine("\nНажмите Enter чтобы закрыть окно");
        Console.ReadLine();
    }

    private static void ExportProblems(List<TestContext> tests, MetaTable t)
    {
        Console.WriteLine("Экспорт отчёта о проблемах...");

        if (t.columns.Any(c => c.name == "Количество проблем" || c.name == "Описание проблем"))
        {
            throw new("Excel table already has export columns. (overwrite protection)");
        }

        int x_flag = t.columns.Count;
        int x_desc = x_flag + 1;

        t.s.Set(x_flag, 0, "Количество проблем");
        t.s.Set(x_desc, 0, "Описание проблем");

        Console.WriteLine(" - подготовка отчёта...");

        Dictionary<int, List<Problem>> problemsByRow = new();

        foreach (TestContext test in tests)
        {
            if (test.status == TestContext.Status.ReadyToRun) continue;

            foreach (Problem problem in test.result.problems)
            {
                int y = problem.y + 1;

                if (problemsByRow.ContainsKey(y) == false)
                {
                    problemsByRow.Add(y, new());
                }
                problemsByRow[y].Add(problem);
            }
        }

        Console.WriteLine($" - - обнаружено {problemsByRow.Sum(kv => kv.Value.Count)} проблем в {problemsByRow.Count} строках");

        Console.WriteLine(" - формирование отчёта...");

        StringBuilder b = new();
        Dictionary<int, StringBuilder> commentByColumnX = new();
        foreach (KeyValuePair<int, List<Problem>> kv in problemsByRow)
        {
            // Set problems number
            t.s.Set(x_flag, kv.Key, kv.Value.Count.ToString());

            foreach (Problem problem in kv.Value)
            {
                b.Append("[");
                b.Append(problem.testCtx.attr.name);
                b.Append(":'");
                b.Append(problem.column.name);
                b.Append("']: ");
                b.AppendLine(problem.explanation);

                if (commentByColumnX.ContainsKey(problem.column.x) == false)
                {
                    commentByColumnX.Add(problem.column.x, new());
                }

                var commentB = commentByColumnX[problem.column.x]!;

                commentB.Append("[");
                commentB.Append(problem.testCtx.attr.name);
                commentB.Append("] ");
                commentB.AppendLine(problem.explanation);
            }

            // Set problems overall description column
            t.s.Set(x_desc, kv.Key, b.ToString());
            b.Clear();

            // Remove all conditional formatting
            t.s.ConditionalFormatting.RemoveAll();

            // Set comments
            foreach (KeyValuePair<int, StringBuilder> ckv in commentByColumnX)
            {
                var commentB = ckv.Value!;

                if (commentB.Length == 0) continue;

                var cell = t.s.At(ckv.Key, kv.Key);

                cell.StyleName = "";
                cell.Style.Fill.SetBackground(Color.Firebrick);

                var comment = cell.AddComment(commentB.ToString(), "ExcelValidator");
                comment.AutoFit = true;

                ckv.Value.Clear();
            }
        }
    }
}