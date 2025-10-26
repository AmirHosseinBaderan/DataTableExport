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

        // Only keep properties that are not ignored and have a valid name
        propertyNames = propertyNames.Where(x => !string.IsNullOrEmpty(x.name) && !x.ignore).ToList();

        // Check that all **required properties** exist in the DataTable columns
        foreach (var prop in propertyNames.Where(x => x.required))
            if (!table.Columns.Contains(prop.name))
                errors.Add($"Required column '{prop.name}' is missing from the Excel file.");


        return errors.Count == 0;
    }

    public static bool IsNullable(this PropertyInfo property)
    {
        // Reference types are nullable
        if (!property.PropertyType.IsValueType)
            return true;

        // Nullable<T> types
        return Nullable.GetUnderlyingType(property.PropertyType) != null;
    }
}