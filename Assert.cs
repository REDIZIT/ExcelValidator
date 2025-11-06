using System;
using System.Globalization;

public class Assert
{
    public Result r;
    public Context ctx = new();

    public Assert(Result r)
    {
        this.r = r;
    }


    public void Equal(Column column, string value) => Equal(column, ctx.y, value);
    public void Equal(Column column, int y, string value)
    {
        That(column, y, s => IsEqual(column, value), actual => $"Ожидалось '{value}', но получено '{actual}'");
    }

    public void NotEqual(Column column, string value) => NotEqual(column, ctx.y, value);
    public void NotEqual(Column column, int y, string value)
    {
        That(column, y, s => IsNotEqual(column, value), actual => $"Ожидалось НЕ '{value}', но получено '{actual}'");
    }

    public void Empty(Column column) => Empty(column, ctx.y);
    public void Empty(Column column, int y)
    {
        That(column, y, s => IsEmpty(column), actual => $"Ожидалось <пусто>, но получено '{actual}'");
    }
    public void NotEmpty(Column column)
    {
        That(column, s => IsNotEmpty(column), actual => $"Ожидалось НЕ <пусто>, но получено '{actual}'");
    }


    public void ZeroOrEmpty(Column column)
    {
        That(column, ctx.y, s => IsZeroOrEmpty(column), actual => $"Ожидалось 0 или <пусто>, но получено '{actual}'");
    }
    public void NotZeroOrEmpty(Column column, Func<string, string>? explaination = null)
    {
        That(column, ctx.y, s => IsNotZeroOrEmpty(column), explaination ?? (actual => $"Ожидалось НЕ 0 и НЕ <пусто>, но получено '{actual}'"));
    }

    public void Contains(Column column, string str)
    {
        That(column, ctx.y, s => IsContains(column, str), actual => $"Ожидалось, что ячейка будет содержать '{str}'");
    }
    public void NotContains(Column column, string str)
    {
        That(column, ctx.y, s => IsContains(column, str) == false, actual => $"Ожидалось, что ячейка НЕ будет содержать '{str}'");
    }


    public bool IsEmpty(Column column) => Is(column, s => string.IsNullOrWhiteSpace(s));
    public bool IsNotEmpty(Column column) => !IsEmpty(column);

    public bool IsZeroOrEmpty(Column column)
    {
        return Is(column, s =>
        {
            if (string.IsNullOrWhiteSpace(s)) return true;
            if (s == "0") return true;

            if (float.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out float flt))
            {
                return flt == 0;
            }

            return false;
        });
    }
    public bool IsNotZeroOrEmpty(Column column) => !IsZeroOrEmpty(column);

    public bool IsEqual(Column column, string value) => Is(column, s => s == value);
    public bool IsNotEqual(Column column, string value) => !IsEqual(column, value);

    public bool IsContains(Column column, string value) => Is(column, s => s.ToLower().Contains(value.ToLower()));


    public bool Is(Column column, Func<string, bool> predicate) => Is(column, ctx.y, predicate);
    public bool Is(Column column, int y, Func<string, bool> predicate)
    {
        string cell = r.t.At<string>(column, y);
        return predicate(cell);
    }

    public void That(Column column, Func<string, bool> goodPredicate, Func<string, string>? explainCallback = null) => That(column, ctx.y, goodPredicate, explainCallback);
    public void That(Column column, int y, Func<string, bool> goodPredicate, Func<string, string>? explainCallback = null)
    {
        if (Is(column, y, goodPredicate)) return;

        if (explainCallback != null)
        {
            string explaination = explainCallback(r.t.At<string>(column, y));
            r.Mark(column, y, explaination);
        }
        else
        {
            r.Mark(column, y);
        }
    }
    public void That(Column column, string explanation)
    {
        r.Mark(column, ctx.y, explanation);
    }
}

public class Context
{
    public int y;
}