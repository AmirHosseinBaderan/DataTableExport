# DataTableExport
an exporter for data table to create excel from it 

# How use this

``` Csharp
public class User{

  // if set ignore to true property not map in excel file
  [ExcelColumn(Ignore = true)]
  public Guid Id {get;set;}

  // if set name for property this name will set in excel column
  [ExcelColumn(Name = "User Name")]
  public string Name {get;set;}

  // if set the type for property type of column will be
  [ExcelColumn(Name = "Registre Date",Type = typeof(string))]
  public DateTime CreateDate {get;set;}
}

```

# How convert list to data table 

``` Csharp

List<User> users = [];
DataTable dt = users.ExportAsTable();

```

# How change value of one column 

``` Csharp

// this way change value of all items for one property

List<User> user = [];
DataTable dt = new DataTableExporter(result.ExportAsTable()).SetCellsValue("Register Date",(value)=>{
  var res = DateTime.TryParse(value.ToString(), out DateTime date);
  return date.ToString("yyyy:MM:dd");
}).Export();

```

# Other functions


## SetCellsValue(string columnName,object value)
set just value to column with out any action

## SetCellsValue(string columnName, Func<object, object> callBack)
set value with callback to column

## SetCellValue(int row, string columnName, object value)
set value with row index and column name 

## SetCellValue(int row, int column, object value)
set value with row index and column index 

## Export()
create data table export at final

