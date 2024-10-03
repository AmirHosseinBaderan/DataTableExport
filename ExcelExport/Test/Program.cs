using ClosedXML.Excel;
using ExcelExport.Import;
using System.Reflection;

var data = Data.base64.ReadDataUrl();
var workBook = LoadExcelFromBytes(data.buffer);
var items = ConvertXLWorkbookToModel<CentralWarehouseExcel>(workBook);

Console.WriteLine(items.Count);

XLWorkbook LoadExcelFromBytes(byte[] excelBytes)
{
    using MemoryStream ms = new(excelBytes);
    return new XLWorkbook(ms);
}


List<T> ConvertXLWorkbookToModel<T>(XLWorkbook workbook) where T : new()
{
    var modelList = new List<T>();

    // Assuming you want to read the first worksheet (index 0)
    var worksheet = workbook.Worksheets.FirstOrDefault();
    if (worksheet is null)
        return [];

    var range = worksheet.RangeUsed();

    // Start reading from row 2 (assuming row 1 contains headers)
    for (int row = 2; row <= range.RowCount(); row++)
    {
        var model = new T();

        // Get the properties of the model using reflection
        PropertyInfo[] properties = typeof(T).GetProperties();

        // Assuming columns in the Excel sheet match property names
        for (int col = 1; col <= properties.Length; col++)
        {
            var cellValue = worksheet.Cell(row, col).Value.ToString();
            properties[col - 1].SetValue(model, Convert.ChangeType(cellValue, properties[col - 1].PropertyType));
        }

        modelList.Add(model);
    }

    return modelList;
}

public record CentralWarehouseExcel
{
    public string ProductName { get; set; }

    public string Barcode { get; set; }

    public string Group { get; set; }

    public string SubGroup { get; set; }
}
