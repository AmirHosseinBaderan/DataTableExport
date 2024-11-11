﻿using FTeam.Excel.Export;
using FTeam.Excel.Import;
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
var items = fs.ReadExcel<ProductExcelModel>();
foreach (var item in items)
    WriteLine(item);


public record ProductExcelModel
{
    [ExcelColumn(Name = "عنوان")]
    public string Name { get; set; }

    [ExcelColumn(Name = "بارکد")]
    public string Barcode { get; set; }

    [ExcelColumn(Name = "شناسه یکتا")]
    public string Identifire { get; set; }

    [ExcelColumn(Name = "واحد")]
    public string Unit { get; set; }

    [ExcelColumn(Name = "مالیات")]
    public string Taxes { get; set; }

    [ExcelColumn(Name = "قیمت")]
    public string Price { get; set; }

    [ExcelColumn(Name = "دسته بندی")]
    public string Category { get; set; }
}