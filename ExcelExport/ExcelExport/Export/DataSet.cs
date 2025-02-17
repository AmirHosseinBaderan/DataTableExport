using ExcelDataReader;
using System.Data;

namespace ExcelExport.Export;

public static class DataSetExtensions
{
    public static DataSet ExportAsDataSet(this IExcelDataReader self, ExcelDataSetConfiguration configuration = null)
    {
        configuration ??= new ExcelDataSetConfiguration();

        self.Reset();
        int sheetIndex = -1;
        DataSet dataSet = new();
        do
        {
            sheetIndex++;
            if (configuration.FilterSheet == null || configuration.FilterSheet(self, sheetIndex))
            {
                var tableConfig = configuration.ConfigureDataTable?.Invoke(self) ?? new ExcelDataTableConfiguration();
                DataTable table = AsDataTable(self, tableConfig);
                dataSet.Tables.Add(table);
            }
        } while (self.NextResult());

        dataSet.AcceptChanges();
        if (configuration.UseColumnDataType)
        {
            FixDataTypes(dataSet);
        }

        self.Reset();
        return dataSet;
    }

    private static string GetUniqueColumnName(DataTable table, string name)
    {
        string uniqueName = name;
        int counter = 1;
        while (table.Columns[uniqueName] != null)
        {
            uniqueName = $"{name}_{counter}";
            counter++;
        }
        return uniqueName;
    }

    private static DataTable AsDataTable(IExcelDataReader self, ExcelDataTableConfiguration configuration)
    {
        DataTable dataTable = new DataTable { TableName = self.Name };
        dataTable.ExtendedProperties.Add("visiblestate", self.VisibleState);

        bool isFirstRow = true;
        List<int> columnIndexes = new List<int>();

        while (self.Read())
        {
            if (isFirstRow)
            {
                if (configuration.UseHeaderRow && configuration.ReadHeaderRow != null)
                    configuration.ReadHeaderRow(self);


                for (int i = 0; i < self.FieldCount; i++)
                    if (configuration.FilterColumn == null || configuration.FilterColumn(self, i))
                    {
                        string columnName = Convert.ToString(self.GetValue(i));
                        if (string.IsNullOrEmpty(columnName))
                            columnName = configuration.EmptyColumnNamePrefix + i;

                        DataColumn column = new DataColumn(GetUniqueColumnName(dataTable, columnName), typeof(object)) { Caption = columnName };
                        dataTable.Columns.Add(column);
                        columnIndexes.Add(i);
                    }


                dataTable.BeginLoadData();
                isFirstRow = false;
                if (configuration.UseHeaderRow)
                    continue;
            }

            if (configuration.FilterRow != null && !configuration.FilterRow(self))
                continue;

            if (IsEmptyRow(self))
                continue;

            DataRow dataRow = dataTable.NewRow();
            for (int j = 0; j < columnIndexes.Count; j++)
            {
                int columnIndex = columnIndexes[j];
                dataRow[j] = self.GetValue(columnIndex);
            }
            dataTable.Rows.Add(dataRow);
        }

        dataTable.EndLoadData();
        return dataTable;
    }

    private static bool IsEmptyRow(IExcelDataReader reader)
    {
        for (int i = 0; i < reader.FieldCount; i++)
        {
            if (reader.GetValue(i) != null)
            {
                return false;
            }
        }
        return true;
    }

    private static void FixDataTypes(DataSet dataset)
    {
        var updatedTables = new List<DataTable>(dataset.Tables.Count);
        bool hasFixedDataTypes = false;

        foreach (DataTable table in dataset.Tables)
        {
            if (table.Rows.Count == 0)
            {
                updatedTables.Add(table);
                continue;
            }

            DataTable newTable = null;
            for (int i = 0; i < table.Columns.Count; i++)
            {
                Type columnType = null;
                foreach (DataRow row in table.Rows)
                {
                    if (!row.IsNull(i))
                    {
                        Type currentType = row[i].GetType();
                        if (currentType != columnType)
                        {
                            if (columnType == null)
                            {
                                columnType = currentType;
                            }
                            else
                            {
                                columnType = null;
                                break;
                            }
                        }
                    }
                }

                if (columnType != null)
                {
                    hasFixedDataTypes = true;
                    newTable ??= table.Clone();
                    newTable.Columns[i].DataType = columnType;
                }
            }

            if (newTable != null)
            {
                newTable.BeginLoadData();
                foreach (DataRow row in table.Rows)
                {
                    newTable.ImportRow(row);
                }
                newTable.EndLoadData();
                updatedTables.Add(newTable);
            }
            else
            {
                updatedTables.Add(table);
            }
        }

        if (hasFixedDataTypes)
        {
            dataset.Tables.Clear();
            dataset.Tables.AddRange(updatedTables.ToArray());
        }
    }
}
