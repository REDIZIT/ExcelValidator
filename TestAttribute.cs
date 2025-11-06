using System;

[AttributeUsage(AttributeTargets.Method)]
public class TestAttribute : Attribute
{
    public string name;
    public string category;

    public TestAttribute(string name, string category)
    {
        this.name = name;
        this.category = category;
    }
}