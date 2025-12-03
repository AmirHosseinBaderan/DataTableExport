using ExcelDataReader;
using ExcelExport.Export;
using ExcelExport.Import;
using FTeam.Excel.Export;
using Microsoft.AspNetCore.Http;
using System.Data;
using System.Reflection;
using System.Text;

namespace FTeam.Excel.Import;

public static class ExcelExtension
{
    // ------------------------ EXCEL READ (IFormFile) ------------------------

    public static IEnumerable<T> ReadExcel<T>(this IFormFile file, bool validator = false) where T : new()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using Stream stream = file.OpenReadStream();
        return ReadExcel<T>(stream, validator);
    }

    public static IEnumerable<T> ReadExcel<T>(this IFormFile file, Func<T, bool>? validator) where T : new()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using Stream stream = file.OpenReadStream();
        return ReadExcel(stream, validator);
    }

    // ------------------------ EXCEL READ (Stream, validator flag) ------------------------

    public static IEnumerable<T> ReadExcel<T>(this Stream stream, bool validator = false) where T : new()
    {
        List<T> result = new();

        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration
               {
                   FallbackEncoding = Encoding.UTF8,
                   LeaveOpen = true,
               }))
        {
            DataSet dataSet = reader.ExportAsDataSet();

            // Use strict validator OR normal loader
            DataTable dataTable = validator
                ? dataSet.ExportDataTableValidateColumns<T>()
                : dataSet.ExportDataTable<T>();

            result = dataTable.ReadFromDataTable<T>(validator);
        }

        return result;
    }

    // ------------------------ EXCEL READ (Stream, row validator) ------------------------

    public static IEnumerable<T> ReadExcel<T>(this Stream stream, Func<T, bool>? validator) where T : new()
    {
        List<T> result = new();

        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration
               {
                   FallbackEncoding = Encoding.UTF8,
                   LeaveOpen = true,
               }))
        {
            DataSet dataSet = reader.ExportAsDataSet();
            DataTable dataTable = dataSet.ExportDataTable<T>();
            result = dataTable.ReadFromDataTable(validator);
        }

        return result;
    }

    // ------------------------ BASE CONVERTER ------------------------

    static DataTable ExportDataTable<TModel>(this DataSet dataSet) where TModel : new()
    {
        DataTable table = dataSet.Tables[0];
        DataTable newTable = new DataTable();

        PropertyInfo[] props = typeof(TModel).GetProperties();

        foreach (var prop in props)
        {
            var attr = prop.GetCustomAttribute<ExcelColumn>();
            if (attr != null && attr.Ignore)
                continue;

            string columnName = attr?.Name ?? prop.Name;
            Type columnType = attr?.Type ?? prop.PropertyType;

            newTable.Columns.Add(columnName, Nullable.GetUnderlyingType(columnType) ?? columnType);
        }

        foreach (DataRow row in table.Rows)
        {
            DataRow newRow = newTable.NewRow();

            foreach (var prop in props)
            {
                var attr = prop.GetCustomAttribute<ExcelColumn>();
                if (attr != null && attr.Ignore)
                    continue;

                string columnName = attr?.Name ?? prop.Name;

                if (table.Columns.Contains(columnName) && row[columnName] != DBNull.Value)
                {
                    try
                    {
                        newRow[columnName] = row[columnName];
                    }
                    catch
                    {
                        try
                        {
                            newRow[columnName] = Convert.ChangeType(row[columnName], prop.PropertyType);
                        }
                        catch
                        {
                        }
                    }
                }
            }

            newTable.Rows.Add(newRow);
        }

        return newTable;
    }

    // ------------------------ STRICT VALIDATION EXPORT ------------------------

    static DataTable ExportDataTableValidateColumns<TModel>(this DataSet dataSet) where TModel : new()
    {
        DataTable table = dataSet.Tables[0];
        DataTable newTable = new DataTable();

        PropertyInfo[] props = typeof(TModel).GetProperties();

        var expected = new List<string>();
        var requiredMap = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

        // Collect expected columns and required flags
        foreach (var prop in props)
        {
            var attr = prop.GetCustomAttribute<ExcelColumn>();
            if (attr != null && attr.Ignore)
                continue;

            string name = attr?.Name ?? prop.Name;

            expected.Add(name);

            // If no attribute → default Required = true? or false?
            // Based on your snippet => default is "return name, !nullable, false"
            bool required = attr?.Required ?? true;

            requiredMap[name] = required;

            Type type = attr?.Type ?? prop.PropertyType;
            newTable.Columns.Add(name, Nullable.GetUnderlyingType(type) ?? type);
        }

        var excelColumns = table.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToList();

        var missing = expected.Except(excelColumns, StringComparer.OrdinalIgnoreCase).ToList();
        var extra = excelColumns.Except(expected, StringComparer.OrdinalIgnoreCase).ToList();

        var errors = new List<string>();

        // ❗ Filter missing columns to only those that are required
        var missingRequired = missing
            .Where(name => requiredMap.TryGetValue(name, out var req) && req)
            .ToList();

        if (missingRequired.Count > 0)
            errors.Add($"Missing required columns: {string.Join(", ", missingRequired)}");

        if (extra.Count > 0)
            errors.Add($"Unknown columns: {string.Join(", ", extra)}");

        if (errors.Count > 0)
            ValidationException.Throw(errors);

        // fill rows
        foreach (DataRow row in table.Rows)
        {
            DataRow newRow = newTable.NewRow();

            foreach (var prop in props)
            {
                var attr = prop.GetCustomAttribute<ExcelColumn>();
                if (attr != null && attr.Ignore)
                    continue;

                string name = attr?.Name ?? prop.Name;

                // If missing but NOT required → skip silently
                if (!excelColumns.Contains(name, StringComparer.OrdinalIgnoreCase))
                    continue;

                if (row[name] == DBNull.Value)
                    continue;

                try
                {
                    newRow[name] = row[name];
                }
                catch
                {
                    try
                    {
                        newRow[name] = Convert.ChangeType(row[name], prop.PropertyType);
                    }
                    catch
                    {
                        throw new Exception($"Invalid value for column '{name}'");
                    }
                }
            }

            newTable.Rows.Add(newRow);
        }

        return newTable;
    }


    // ------------------------ ROW MAPPING ------------------------

    public static List<T> ReadFromDataTable<T>(this DataTable table, Func<T, bool>? validator) where T : new()
    {
        List<T> result = new();
        PropertyInfo[] props = typeof(T).GetProperties();

        foreach (DataRow row in table.Rows)
        {
            T? obj = row.ConvertRowToModel<T>(table.Columns, props);

            if (obj is not null && (validator == null || validator(obj)))
                result.Add(obj);
        }

        return result;
    }

    static T? ConvertRowToModel<T>(this DataRow row, DataColumnCollection dataColumn, PropertyInfo[] props)
        where T : new()
    {
        T? obj = new();

        foreach (DataColumn column in dataColumn)
        {
            try
            {
                string name = column.ColumnName;

                PropertyInfo? prop = props.FirstOrDefault(p =>
                {
                    var attr = p.GetCustomAttribute<ExcelColumn>();
                    if (attr == null)
                        return p.Name.Equals(name, StringComparison.OrdinalIgnoreCase);
                    if (attr.Ignore)
                        return false;
                    return name.Equals(attr.Name, StringComparison.OrdinalIgnoreCase);
                });

                if (prop is not null && row[column] != DBNull.Value)
                {
                    var attr = prop.GetCustomAttribute<ExcelColumn>();
                    Type type = attr?.Type ?? prop.PropertyType;

                    object value = Convert.ChangeType(row[column], type);
                    prop.SetValue(obj, value);
                }
            }
            catch
            {
                obj = default;
            }
        }

        return obj;
    }

    // ------------------------ MODEL VALIDATION ------------------------

    static List<T> ReadFromDataTable<T>(this DataTable table, bool validator = false) where T : new()
    {
        List<T> result = new();
        PropertyInfo[] props = typeof(T).GetProperties();

        if (validator && !table.ValidateColumnNames<T>(out List<string> validationErrors))
            ValidationException.Throw(validationErrors);

        bool skipFirst = true;

        foreach (DataRow row in table.Rows)
        {
            // Skip header row in Excel
            if (skipFirst)
            {
                skipFirst = false;
                continue;
            }

            T? obj = row.ConvertRowToModel<T>(table.Columns, props);
            if (obj is not null)
                result.Add(obj);
        }

        return result;
    }

    // ------------------------ CSV ------------------------

    public static IEnumerable<T> ReadCsv<T>(this IFormFile file) where T : new()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var stream = file.OpenReadStream();
        return stream.ReadCsv<T>();
    }

    public static IEnumerable<T> ReadCsv<T>(this Stream stream) where T : new()
    {
        DataTable table = new();
        using StreamReader reader = new(stream, Encoding.UTF8, true);

        string[] lines = reader.ReadToEnd().Split('\n');

        if (lines.Length > 0)
        {
            var headers = lines[0].Split(',');
            foreach (var h in headers)
                table.Columns.Add(h);

            for (int i = 1; i < lines.Length; i++)
                table.Rows.Add(lines[i].Split(','));
        }

        return table.ReadFromDataTable<T>();
    }

    // ------------------------ XML ------------------------

    public static IEnumerable<T> ReadXml<T>(this IFormFile file) where T : new()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using Stream stream = file.OpenReadStream();
        return stream.ReadXml<T>();
    }

    public static IEnumerable<T> ReadXml<T>(this Stream stream) where T : new()
    {
        using StreamReader reader = new(stream);

        DataSet ds = new();
        ds.ReadXml(stream);

        List<T> result = new();

        foreach (DataTable table in ds.Tables)
            result.AddRange(table.ReadFromDataTable<T>());

        return result.Distinct();
    }

    // ------------------------ GENERIC FILE ------------------------

    public static IEnumerable<T> ReadFromFile<T>(this IFormFile file) where T : new()
    {
        if (file == null)
            return default;

        string ext = Path.GetExtension(file.FileName).ToLower();

        return ext switch
        {
            ".csv" or ".txt" => file.ReadCsv<T>(),
            ".xlsx" or ".xls" or ".xlsm" => file.ReadExcel<T>(),
            ".xml" => file.ReadXml<T>(),
            _ => file.ReadExcel<T>()
        };
    }

    public static IEnumerable<T> ReadFromFile<T>(this string dataUrl) where T : new()
    {
        if (string.IsNullOrEmpty(dataUrl))
            return default;

        var file = dataUrl.ToStream();
        var (_, ext) = dataUrl.GetContentType();

        return ext switch
        {
            ".csv" or ".txt" => file.ReadCsv<T>(),
            ".xlsx" or ".xls" or ".xlsm" => file.ReadExcel<T>(),
            ".xml" => file.ReadXml<T>(),
            _ => file.ReadExcel<T>()
        };
    }
}