using System;
using System.Collections.Generic;

public class Result
{
    public bool IsGood => problems.Count == 0;

    public Assert assert;
    public MetaTable t;
    public Exception? exception;

    public List<Problem> problems = new();
    public TestContext ctx;

    public Result(MetaTable t)
    {
        this.t = t;
        assert = new(this);
    }

    public void Mark(Column column, int row, string? explanation = null)
    {
        problems.Add(new ()
        {
            column = column,
            y = row,
            explanation = explanation,
            testCtx = ctx
        });
    }
}