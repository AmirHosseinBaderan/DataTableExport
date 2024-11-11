using ExcelDataReader;
using ExcelExport.Import;
using FTeam.Excel.Export;
using Microsoft.AspNetCore.Http;
using System.Data;
using System.Reflection;
using System.Text;

namespace FTeam.Excel.Import;

public static class ExcelExtension
{

    public static IEnumerable<T> ReadExcel<T>(this IFormFile file) where T : new()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using Stream stream = file.OpenReadStream();
        return ReadExcel<T>(stream);
    }

    public static IEnumerable<T> ReadExcel<T>(this Stream stream) where T : new()
    {
        List<T> result = new();

        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration
        {
            FallbackEncoding = Encoding.UTF8,
            LeaveOpen = true,
        }))
        {
            DataSet dataSet = reader.AsDataSet();
            DataTable dataTable = dataSet.ExportDataTable();
            result = dataTable.ReadFromDataTable<T>();
        }
        return result;
    }

    static DataTable ExportDataTable(this DataSet dataSet)
    {
        // Get DataTable from data set 
        DataTable table = dataSet.Tables[0];

        // Get first row of table
        var firstRow = table.Rows[0];
        // Declare columns names 
        List<string> columnNames = new();
        foreach (var cell in firstRow.ItemArray)
            columnNames.Add(cell.ToString());

        table.Rows.Remove(firstRow);
        for (int i = 0; i < table.Columns.Count; i++)
            table.Columns[i].ColumnName = columnNames[i];

        return table;
    }

    static List<T> ReadFromDataTable<T>(this DataTable table) where T : new()
    {
        List<T> result = new();
        PropertyInfo[] props = typeof(T).GetProperties();
        foreach (DataRow row in table.Rows)
        {

            T obj = new();

            // Map properties from DataRow to object properties
            foreach (DataColumn column in table.Columns)
            {
                try
                {
                    string propertyName = column.ColumnName;
                    PropertyInfo prop = props.FirstOrDefault((p) =>
                    {
                        var attr = p.GetCustomAttribute<ExcelColumn>();
                        if (attr is null)
                            return p.Name.Equals(propertyName, StringComparison.CurrentCultureIgnoreCase);
                        if (attr.Ignore)
                            return false;
                        else
                            return propertyName.Equals(attr.Name, StringComparison.CurrentCultureIgnoreCase);

                    });

                    if (prop != null && row[column] != DBNull.Value)
                        prop.SetValue(obj, row[column]);
                }
                catch
                {
                }
            }

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
