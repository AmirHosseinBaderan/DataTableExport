using System;
using System.Data;

namespace FTeam.Excel.Export;

public class DataTableExporter
{
    private readonly DataTable _dataTable;

    public DataTableExporter(DataTable dataTable) => _dataTable = dataTable;

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
            object def = item[columnName];
            object value = callBack(def);
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