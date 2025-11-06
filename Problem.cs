public class Problem
{
    public Column column;
    public int y;
    public string? explanation;
    public TestContext testCtx;

    public override string ToString()
    {
        // string address = $"{column.alphabetName}${y + 2}";
        //string address = $"{y + 2}";
        //return $"[{address}:{address}] '{column.name}':{y + 2}" + (explanation != null ? $" - {explanation}" : null);
        return $"'{column.name}':{y + 2}" + (explanation != null ? $" - {explanation}" : null);
    }
}