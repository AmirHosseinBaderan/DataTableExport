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
var items = fs.ReadExcel<UserCompanyViewModel>();
foreach (var item in items)
    WriteLine(item.UserCompanyCompanyName);

//IEnumerable<UserCompanyViewModel> items2 = fs.ReadExcel<UserCompanyViewModel>();
//foreach (var item2 in items2)
//    WriteLine(item2.UserCompanyCompanyName);

public record ProductExcelModel
{
    [ExcelColumn(Name = "Name")]
    public string Name { get; set; }

    [ExcelColumn(Name = "Barcode")]
    public string Barcode { get; set; }

    [ExcelColumn(Name = "Identifier")]
    public string Identifire { get; set; }

    [ExcelColumn(Name = "Unit")]
    public string Unit { get; set; }

    [ExcelColumn(Name = "Taxes")]
    public string Taxes { get; set; }

    [ExcelColumn(Name = "Price")]
    public string Price { get; set; }

    [ExcelColumn(Name = "Category")]
    public string Category { get; set; }

    [ExcelColumn(Name = "زیر دسته بندی")]
    public string SubCategory { get; set; }
}

public class UserCompanyViewModel
{
    [ExcelColumn(Name = "نام شرکت")]
    public string? UserCompanyCompanyName { get; set; }

    [ExcelColumn(Name = "نام فروشگاه")]
    [DisplayName("نام فروشگاه")]
    public string? UserCompanyStoreName { get; set; }

    [ExcelColumn(Name = "نام صاحب شرکت")]
    public string? UserCompanyCompanyOwnerName { get; set; }

    [ExcelColumn(Name = "نام صاحب فروشگاه")]
    public string? UserCompanyStoreOwnerName { get; set; }

    [ExcelColumn(Name = "شناسه ملی")]
    public string? UserCompanyNationalId { get; set; }

    [ExcelColumn(Name = "شماره ملی")]
    public string? UserCompanyNationalCode { get; set; }

    [ExcelColumn(Name = "شناسه اقتصادی")]
    public string? UserCompanyEconomicId { get; set; }

    [ExcelColumn(Name = "شماره موبایل")]
    public string? UserCompanyMobileNumber { get; set; }

    [ExcelColumn(Name = "آدرس")]
    public string? UserCompanyAddress { get; set; }

    [ExcelColumn(Name = "کد پستی")]
    public string? UserCompanyPostalCode { get; set; }

    [ExcelColumn(Name = "تلفن")]
    public string? UserCompanyTelephone { get; set; }

    [ExcelColumn(Name = "نوع شرکت")]
    public string? CompanyTypeId { get; set; }

    [ExcelColumn(Name = "شماره ثبت")]
    public string? UserCompanyRegisterNumber { get; set; }

    [ExcelColumn(Name = "استان")]
    public string? UserCompanyProvince { get; set; }

    [ExcelColumn(Name = "شهر")]
    public string? UserCompanyCity { get; set; }

    [ExcelColumn(Name = "وضعیت")]
    public string? UserCompanyStatus { get; set; }

    [ExcelColumn(Name = "نوع مودی")]
    public string? UserCompanyType { get; set; }
}
