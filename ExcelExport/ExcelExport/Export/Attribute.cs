namespace FTeam.Excel.Export;

[AttributeUsage(AttributeTargets.All, Inherited = false, AllowMultiple = true)]
public class ExcelColumn : Attribute
{
    public string Name { get; set; } = null!;

    public bool Ignore { get; set; } = false;

    public Type Type { get; set; }
}