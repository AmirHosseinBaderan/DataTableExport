using System.Data;
using System.Reflection;

List<User> users = [new(Guid.NewGuid(), "Test", DateTime.Now)];
var table = new DataTableExporter(users.ExportAsTable())
                                .SetCellsValue("UserName", "Hossein")
                                .SetCellsValue("CreateDate", (current) => (DateTime.Parse(current.ToString() ?? "").ToString("yyyy:MM:dd")))
                                .Export();

Console.WriteLine(table.Rows.Count);


record User
{
    public User(Guid id, string name, DateTime date)
    {
        Id = id;
        Name = name;
        Date = date;
    }

    [ExcelColumn(Name = "UserId")]
    public Guid Id { get; set; }

    [ExcelColumn(Name = "UserName")]
    public string Name { get; set; }

    [ExcelColumn(Name = "CreateDate")]
    public DateTime Date { get; set; }
}


public static class DataTableExtension
{
    public static DataTable ExportAsTable<T>(this IEnumerable<T> data)
    {
        DataTable table = new();
        PropertyInfo[] props = typeof(T).GetProperties();
        table.Columns.AddRange(props.Select(x => new DataColumn(x.ColName())).ToArray());

        foreach (var item in data)
        {
            DataRow row = table.NewRow();
            foreach (var prop in props)
                if (!prop.ColIgnore())
                    row[prop.ColName()] = prop.GetValue(item) ?? DBNull.Value;
            table.Rows.Add(row);
        }

        return table;
    }

    private static string ColName(this PropertyInfo prop)
    {
        try
        {
            var column = (ExcelColumn)Attribute.GetCustomAttribute(prop, typeof(ExcelColumn));
            return column != null ? column.Name : prop.Name;
        }
        catch
        {
            return prop.Name;
        }
    }

    private static bool ColIgnore(this PropertyInfo prop)
    {
        try
        {
            var column = (ExcelColumn)Attribute.GetCustomAttribute(prop, typeof(ExcelColumn));
            return column != null && column.Ignore;
        }
        catch
        {
            return false;
        }
    }
}


public class DataTableExporter
{
    private readonly DataTable _dataTable;

    public DataTableExporter(DataTable dataTable)
    {
        _dataTable = dataTable;
    }

    public DataTableExporter SetCellValue(int row, int column, object value)
    {
        _dataTable.Rows[row][column] = value;
        return this;
    }

    public DataTableExporter SetCellValue(int row, string columnName, object value)
    {
        _dataTable.Rows[row][columnName] = value;
        return this;
    }

    public DataTableExporter SetCellsValue(string columnName, object value)
    {
        foreach (DataRow item in _dataTable.Rows)
            item[columnName] = value;

        return this;
    }

    public DataTableExporter SetCellsValue(string columnName, Func<object, object> callBack)
    {
        foreach (DataRow item in _dataTable.Rows)
        {
            var def = item[columnName];
            var value = callBack(def);
            item[columnName] = value;
        }

        return this;
    }

    // Add other methods for sorting, filtering, etc.

    public DataTable Export()
    {
        // Perform any final operations before exporting the modified DataTable.
        return _dataTable;
    }
}

[System.AttributeUsage(AttributeTargets.All, Inherited = true, AllowMultiple = true)]
sealed class ExcelColumn : Attribute
{

    public string Name { get; set; } = null!;

    public bool Igonre = {get;set;} = false;
}
