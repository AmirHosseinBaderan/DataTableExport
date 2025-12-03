using FTeam.Excel.Export;
using FTeam.Excel.Import;
using System.ComponentModel;
using System.Text;
using static System.Console;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

string? path = null;

while (path is null)
{
    Write("Enter Excel Path : ");
    path = ReadLine();
}

FileStream fs = new(path, FileMode.Open, FileAccess.Read);
var items = fs.ReadExcel<ProductExcelModel>(validator: true);
foreach (var item in items)
    WriteLine(item.Name);

//IEnumerable<UserCompanyViewModel> items2 = fs.ReadExcel<UserCompanyViewModel>();
//foreach (var item2 in items2)
//    WriteLine(item2.UserCompanyCompanyName);

public record ProductExcelModel
{
    [ExcelColumn(Name = "عنوان",Required = true)] public string Name { get; set; }

    [ExcelColumn(Name = "بارکد",Required = true)] public string Barcode { get; set; }

    [ExcelColumn(Name = "شناسه یکتا",Required = true)] public string Identifire { get; set; }

    [ExcelColumn(Name = "واحد",Required = true)] public string Unit { get; set; }

    [ExcelColumn(Name = "مالیات",Required = true)] public string Taxes { get; set; }

    [ExcelColumn(Name = "قیمت",Required = true)] public string Price { get; set; }

    [ExcelColumn(Name = "دسته بندی",Required = true)] public string Category { get; set; }

    [ExcelColumn(Name = "زیردسته بندی")] public string SubCategory { get; set; }
}