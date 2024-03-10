using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;

namespace FTeam.Excel.Export;

public static class ExcelExtension
{
    public static IWorkbook ExportFromDataTable(this DataTable dataTable, string sheetName, string extension = "xlsx")
    {
        IWorkbook workbook = extension switch
        {
            "xlsx" => new XSSFWorkbook(),
            "xls" => new HSSFWorkbook(),
            _ => throw new Exception("The format '" + extension + "' is not supported.")
        };

        ISheet sheet = workbook.CreateSheet(sheetName);

        IRow columns = sheet.CreateRow(0);
        for (int i = 0; i < dataTable.Columns.Count; i++)
        {
            var cell = columns.CreateCell(i);
            var name = dataTable.Columns[i].ColumnName;
            cell.SetCellValue(name);
        }

        // set data 
        for (var i = 0; i < dataTable.Rows.Count; i++)
        {
            var row = sheet.CreateRow(i + 1);
            for (var j = 0; j < dataTable.Columns.Count; j++)
            {
                var cell = row.CreateCell(j);
                var columnName = dataTable.Columns[j].ToString();
                cell.SetCellValue(dataTable.Rows[i][columnName].ToString());
                cell.CellStyle.WrapText = true; // NOT WORKING
            }
        }

        return workbook;
    }

    public async static Task<MemoryStream> ExportToSteamAsync(this IWorkbook workbook)
    {
        MemoryStream tempStream = null;
        MemoryStream stream = null;
        try
        {
            // 1. Write the workbook to a temporary stream
            tempStream = new MemoryStream();
            workbook.Write(tempStream, false);
            // 2. Convert the tempStream to byteArray and copy to another stream
            var byteArray = tempStream.ToArray();
            stream = new MemoryStream();
            await stream.WriteAsync(byteArray);
            stream.Seek(0, SeekOrigin.Begin);

            return stream;
        }
        finally
        {
            if (tempStream != null) await tempStream.DisposeAsync();
            if (stream != null) await stream.DisposeAsync();
        }
    }
}
