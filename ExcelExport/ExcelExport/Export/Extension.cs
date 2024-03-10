using System.Data;
using System.Reflection;

namespace FTeam.Excel.Export;

public static class DataTableExtension
{
    public static DataTable ExportAsTable<T>(this IEnumerable<T> data)
    {
        DataTable table = new();
        PropertyInfo[] props = typeof(T).GetProperties();
        table.Columns.AddRange(props
                                .Where(p => !p.ColIgnore())
                                .Select(x => new DataColumn(x.ColName()))
                                            .ToArray());

        foreach (var item in data)
        {
            DataRow row = table.NewRow();
            foreach (var prop in props)
                if (!prop.ColIgnore())
                    row[prop.ColName()] = prop.ColValue(item) ?? DBNull.Value;
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

    private static object ColValue<T>(this PropertyInfo prop, T obj)
    {
        var column = (ExcelColumn)Attribute.GetCustomAttribute(prop, typeof(ExcelColumn));
        return column is null || column.Type is null
                ? prop.GetValue(obj)
                : Convert.ChangeType(prop.GetValue(obj), column.Type);
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
