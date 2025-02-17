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
            DataTable dataTable = dataSet.ExportDataTable<T>();
            result = dataTable.ReadFromDataTable<T>(validator);
        }
        return result;
    }

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

    static DataTable ExportDataTable<TModel>(this DataSet dataSet) where TModel : new()
    {
        // Get DataTable from data set 
        DataTable table = dataSet.Tables[0];

        // Create a new DataTable
        DataTable newTable = new DataTable();

        // Get the properties of the model
        PropertyInfo[] props = typeof(TModel).GetProperties();

        // Add columns to the new DataTable based on the model properties
        foreach (var prop in props)
        {
            var attr = prop.GetCustomAttribute<ExcelColumn>();
            if (attr != null && attr.Ignore)
                continue;

            string columnName = attr?.Name ?? prop.Name;
            Type columnType = attr?.Type ?? prop.PropertyType;

            newTable.Columns.Add(columnName, Nullable.GetUnderlyingType(columnType) ?? columnType);
        }

        // Populate the new DataTable with rows from the original DataTable
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

    public static List<T> ReadFromDataTable<T>(this DataTable table, Func<T, bool>? validator) where T : new()
    {
        List<T> result = new();
        PropertyInfo[] props = typeof(T).GetProperties();

        foreach (DataRow row in table.Rows)
        {
            T? obj = row.ConvertRowToModel<T>(table.Columns, props);
            if (obj is not null)
                // Apply the validator if provided
                if (validator == null || validator(obj))
                    result.Add(obj);
        }
        return result;
    }

    static T? ConvertRowToModel<T>(this DataRow row, DataColumnCollection dataColumn, PropertyInfo[] props) where T : new()
    {
        T? obj = new();

        // Map properties from DataRow to object properties
        foreach (DataColumn column in dataColumn)
        {
            try
            {
                string propertyName = column.ColumnName;
                PropertyInfo? prop = props.FirstOrDefault((p) =>
                {
                    var attr = p.GetCustomAttribute<ExcelColumn>();
                    if (attr is null)
                        return p.Name.Equals(propertyName, StringComparison.CurrentCultureIgnoreCase);
                    if (attr.Ignore)
                        return false;
                    else
                        return propertyName.Equals(attr.Name, StringComparison.CurrentCultureIgnoreCase);
                });

                if (prop is not null && row[column] != DBNull.Value)
                {
                    ExcelColumn? attr = prop.GetCustomAttribute<ExcelColumn>();
                    Type convertType = attr is null || attr.Type is null ? prop.PropertyType : attr.Type;
                    object convertedValue = Convert.ChangeType(row[column], convertType);
                    prop.SetValue(obj, convertedValue);
                }
            }
            catch
            {
                obj = default;
            }
        }
        return obj;
    }


    static List<T> ReadFromDataTable<T>(this DataTable table, bool validator = false) where T : new()
    {
        List<T> result = new();
        PropertyInfo[] props = typeof(T).GetProperties();

        if (validator && !table.ValidateColumnNames<T>(out List<string> validationErrors))
            ValidationException.Throw(validationErrors);

        foreach (DataRow row in table.Rows)
        {
            T? obj = row.ConvertRowToModel<T>(table.Columns, props);
            if (obj is not null)
                result.Add(obj);
        }
        return result;
    }

    public static IEnumerable<T> ReadCsv<T>(this IFormFile file) where T : new()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var stream = file.OpenReadStream();
        return stream.ReadCsv<T>();
    }

    public static IEnumerable<T> ReadCsv<T>(this Stream stream) where T : new()
    {
        DataTable dataTable = new();
        using StreamReader reader = new(stream, Encoding.UTF8, true);
        string[] lines = reader.ReadToEnd().Split('\n');

        if (lines.Length > 0)
        {
            var headers = lines[0].Split(','); // Assuming the first row contains column headers
            foreach (var header in headers)
                dataTable.Columns.Add(header);

            for (int i = 1; i < lines.Length; i++)
            {
                var values = lines[i].Split(',');
                dataTable.Rows.Add(values);
            }
        }

        return dataTable.ReadFromDataTable<T>();
    }


    public static IEnumerable<T> ReadXml<T>(this IFormFile file) where T : new()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using Stream stream = file.OpenReadStream();
        return stream.ReadXml<T>();
    }

    public static IEnumerable<T> ReadXml<T>(this Stream stream) where T : new()
    {
        using StreamReader reader = new(stream);
        DataSet dataSet = new();
        dataSet.ReadXml(stream);
        List<T> result = new();
        foreach (DataTable item in dataSet.Tables)
        {
            var res = item.ReadFromDataTable<T>();
            result.AddRange(res);
        }
        return result.Distinct();
    }

    public static IEnumerable<T> ReadFromFile<T>(this IFormFile file) where T : new()
    {
        if (file is null)
            return default;

        string fileName = file.FileName;
        string extension = Path.GetExtension(fileName).ToLower();

        if (extension == ".csv" || extension == ".txt")
            return file.ReadCsv<T>();
        else if (extension == ".xlsx" || extension == ".xls" || extension == ".xlsm")
            return file.ReadExcel<T>();
        else if (extension == ".xml")
            return file.ReadXml<T>();

        return file.ReadExcel<T>();
    }

    public static IEnumerable<T> ReadFromFile<T>(this string dataUrl) where T : new()
    {
        if (string.IsNullOrEmpty(dataUrl))
            return default;

        var file = dataUrl.ToStream();
        var (contentType, extension) = dataUrl.GetContentType();

        if (extension == ".csv" || extension == ".txt")
            return file.ReadCsv<T>();
        else if (extension == ".xlsx" || extension == ".xls" || extension == ".xlsm")
            return file.ReadExcel<T>();
        else if (extension == ".xml")
            return file.ReadXml<T>();

        return file.ReadExcel<T>();
    }
}
