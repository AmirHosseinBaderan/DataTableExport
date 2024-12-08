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
        List<(string name, bool required, bool ignore)> propertyNames = props.Select(p =>
        {
            var attr = p.GetCustomAttribute<ExcelColumn>();

            if (attr is null)
            {
                var nullable = p.IsNullable();
                return (p.Name, !nullable, false);
            }
            if (attr.Ignore)
                return ("", false, true);
            var name = attr.Name;
            return (name, attr.Required, false);
        }).ToList();
        propertyNames = propertyNames.Where(x => !string.IsNullOrEmpty(x.name) && !x.ignore).ToList();

        // Check if all DataTable columns match the model properties
        foreach (DataColumn column in table.Columns)
            if (!propertyNames.Select(x => x.name)
                 .Contains(column.ColumnName, StringComparer.CurrentCultureIgnoreCase))
                errors.Add($"Column '{column.ColumnName}' does not match any property in the model.");

        return errors.Count == 0;
    }

    public static bool IsNullable(this PropertyInfo property)
    {
        // Check if the property type is a reference type or a Nullable<T>
        if (!property.PropertyType.IsValueType)
        {
            return true; // Reference types are nullable
        }

        // Check if the property type is a Nullable<T>
        return Nullable.GetUnderlyingType(property.PropertyType) != null;
    }

}
