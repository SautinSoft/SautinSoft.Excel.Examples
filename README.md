# SautinSoft.Excel

![Nuget](https://img.shields.io/nuget/v/SautinSoft.Excel) ![Nuget](https://img.shields.io/nuget/dt/SautinSoft.Excel) 
![Passed](https://img.shields.io/badge/build-%20%E2%9C%93%202505%20tests%20passed%20(0%20failed)-logo=visualstudio) 
![windows](https://img.shields.io/badge/%20-%20%E2%9C%93?logo=windows)
![macOS](https://img.shields.io/badge/%20-%20%E2%9C%93?logo=apple)
![linux](https://img.shields.io/badge/%20-%20%E2%9C%93?logo=linux&logoColor=white)
![docker](https://img.shields.io/badge/%20-%20%E2%9C%93?logo=docker&logoColor=white)
![aws](https://img.shields.io/badge/%20-%20%E2%9C%93?logo=amazonaws)
![microsoftazure](https://img.shields.io/badge/%20-%20%E2%9C%93?logo=microsoftazure)
# Excel .Net is a standalone C# assembly which gives you full set of API to manipulate (read, write, edit, convert) with documents in XLSX, XLS, XLSB, CSV and others formats.


![Excel](https://user-images.githubusercontent.com/79837963/229030126-091cb2c1-5b13-4295-8f44-ed2b3e34aab1.png)



[SautinSoft.Excel](https://sautinsoft.com/products/excel/) is a standalone C# assembly which gives you full set of API to manipulate (read, write, edit, convert) with documents in XLSX, XLS, XLSB, CSV and others formats.

## Quick links

+ [Developer Guide](https://sautinsoft.com/products/excel/help/net/getting-started/overview.php)
+ [API Reference](https://sautinsoft.com/products/excel/help/net/api-reference/html/N_SautinSoft.htm)

## Top Features

+ [Convert Excel file to PDF file.](https://github.com/SautinSoft/SautinSoft.Excel.Examples/tree/main/CSharp)
+ [Create Excel document](https://github.com/SautinSoft/SautinSoft.Excel.Examples/tree/main/CSharp)
+ [Load Excel file](https://github.com/SautinSoft/SautinSoft.Excel.Examples/tree/main/CSharp)
+ [Modify XLSX/XLS documents](https://github.com/SautinSoft/SautinSoft.Excel.Examples/tree/main/CSharp)

## System Requirement

* .NET Framework 4.6.1 - 4.8.1
* .NET 5, 6, 7, 8, 9
* Windows, Linux, macOS, Android, iOS.

## Getting Started with Excel .Net

Are you ready to give Excel .NET a try? Simply execute `Install-Package sautinsoft.excel` from Package Manager Console in Visual Studio to fetch the NuGet package. If you already have Excel .NET and want to upgrade the version, please execute `Update-Package sautinsoft.excel` to get the latest version.

## Convert XLSX/XLS to PDF

```csharp
string inpFile = @"..\..\..\Example.xlsx";
string inpFile = @"..\..\..\Example.xls";
string outFile = @"..\..\..\Result.pdf";
ExcelDocument excelDocument = ExcelDocument.Load(inpFile);
excelDocument.Save(outFile, new PdfSaveOptions());
```
## Create Excel document

```csharp
ExcelDocument excelDocument = new ExcelDocument();
excelDocument.Worksheets.Add("The main worksheet");
excelDocument.Worksheets.Add("Second worksheet");

// Create a variable to address
var worksheet = excelDocument.Worksheets["The main worksheet"];

// Add plain text
worksheet.Cells["A1"].Value = "This is common string";
worksheet.Cells["B1"].Value = "Hello, World! 12345";

// Add the result of  the expression
worksheet.Cells["A2"].Value = "This is the result of a mathematical expression in C#";
worksheet.Cells["B2"].Value = 5 + 5;

excelDocument.Save(outFile, new XlsxSaveOptions());
```
## Load Excel file

```csharp
string filePath = @"..\..\..\example.xlsx";
// The file format is detected automatically from the file extension: ".xlsx".
ExcelDocument excel = ExcelDocument.Load(filePath);
    if (excel != null)
Console.WriteLine("Loaded successfully!");
```
## Modify XLSX/XLS documents

```csharp
string image = @"..\..\..\cup.jpg";
string outFile = @"..\..\..\Result.xlsx";

ExcelDocument excelDocument = new ExcelDocument();
excelDocument.Worksheets.Add("Page 1");
var worksheet = excelDocument.Worksheets["Page 1"];

// Insert an image
worksheet.Pictures.Add(image, SKRect.Create(1080, 960));

excelDocument.Save(outFile);
```

## Resources

+ **Website:** [www.sautinsoft.com](https://www.sautinsoft.com)
+ **Product Home:** [Excel .Net](https://sautinsoft.com/products/excel/)
+ [Download SautinSoft.Excel](https://sautinsoft.com/products/excel/download.php)
+ [Developer Guide](https://sautinsoft.com/products/excel/help/net/getting-started/overview.php)
+ [API Reference](https://sautinsoft.com/products/excel/help/net/api-reference/html/N_SautinSoft.htm)
+ [Support Team](https://sautinsoft.com/support.php)
+ [License](https://sautinsoft.com/products/excel/help/net/getting-started/agreement.php)


