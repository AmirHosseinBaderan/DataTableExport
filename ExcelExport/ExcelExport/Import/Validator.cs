using FTeam.Excel.Export;
using System.Data;
using System.Reflection;

namespace ExcelExport.Import;

public static class ColumnValidator
{
    public static bool ValidateColumnNames<T>(this DataTable table, out List<string> errors) where T : new()
    {
        errors = new List<string>();
        PropertyInfo[] props = typeof(T).GetProperties();

        // Get the property names of the model
        List<string> propertyNames = props.Select(p =>
        {
            var attr = p.GetCustomAttribute<ExcelColumn>();
            return attr != null && !attr.Ignore ? attr.Name : p.Name;
        }).ToList();

        // Check if all DataTable columns match the model properties
        foreach (DataColumn column in table.Columns)
            if (!propertyNames.Contains(column.ColumnName, StringComparer.CurrentCultureIgnoreCase))
                errors.Add($"Column '{column.ColumnName}' does not match any property in the model.");

        return errors.Count == 0;
    }
}
