using System.Reflection;

public static class Rules
{
    public static List<TestContext> Check(MetaTable t)
    {
        var tests = CreateTestsContexts(t);

        
        Dictionary<string, List<TestContext>> testsByCategory = new();

        foreach (TestContext test in tests)
        {
            if (testsByCategory.ContainsKey(test.attr.category) == false)
            {
                testsByCategory.Add(test.attr.category, new());
            }
            testsByCategory[test.attr.category].Add(test);
        }


        Console.WriteLine($"\nКатегории тестов:");
        int i = 0;
        foreach (string category in testsByCategory.Keys)
        {
            i++;
            Console.WriteLine($"{i}. {category} ({testsByCategory[category].Count} тестов)");
            foreach (TestContext test in testsByCategory[category])
            {
                Console.WriteLine($" - {test.attr.name}");
            }
        }


        Console.WriteLine("\nВведите номер (или номера через запятую) категорий для запуска (или оставьте пустым для запуска всех тестов):");

        string input = Console.ReadLine()!;
        string[] inputs = input.Split(',').Select(s => s.Trim()).ToArray();

        if (inputs.Length == 0 || (inputs.Length == 1 && string.IsNullOrWhiteSpace(inputs[0])))
        {
            RunAll(tests);
        }
        else
        {
            foreach (string inputStr in inputs)
            {
                int inputIndex = int.Parse(inputStr) - 1;
                foreach (TestContext inputTest in testsByCategory[testsByCategory.Keys.ElementAt(inputIndex)])
                {
                    RunSingle(inputTest);
                }
            }
        }

        return tests;
    }

    private static void RunSingle(TestContext test)
    {
        string label = $"Test '{test.attr.name}':";
        Console.ResetColor();
        Console.WriteLine(label);

        test.Run();

        if (test.status == TestContext.Status.Ok)
        {
            Console.BackgroundColor = ConsoleColor.Green;
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write(" OK ");
            Console.ResetColor();
            Console.WriteLine();
        }
        else
        {
            Console.BackgroundColor = ConsoleColor.Red;
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write($" {test.status.ToString().ToUpper()} ");
            Console.ResetColor();
            Console.WriteLine();

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Проблемы ({test.result.problems.Count}):");
            int j = 0;
            foreach (Problem problem in test.result.problems)
            {
                j++;
                if (j >= 20)
                {
                    Console.WriteLine($" {j}..{test.result.problems.Count} (вывод скрыт из-за большого количества элементов)");
                    break;
                }
                Console.WriteLine($" {j}. {problem.ToString()}");
            }
            Console.ResetColor();

            if (test.result.exception != null)
            {
                Console.ResetColor();
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(test.result.exception.ToString());
            }
        }

        Console.WriteLine();
    }
    private static void RunAll(List<TestContext> tests)
    {
        for (int i = 0; i < tests.Count; i++)
        {
            RunSingle(tests[i]);
        }
    }

    private static List<TestContext> CreateTestsContexts(MetaTable table)
    {
        List<TestContext> tests = new();

        Type type = typeof(TestContext);
        foreach (MethodInfo method in type.GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
        {
            TestAttribute? attr = method.GetCustomAttribute<TestAttribute>();
            if (attr == null) continue;

            tests.Add(new(table)
            {
                attr = attr,
                testMethod = method,
                status = TestContext.Status.ReadyToRun
            });
        }

        return tests;
    }
}