using SautinSoft.Excel;
using SkiaSharp;
using System;
using System.Text;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free key here:   
            // https://sautinsoft.com/start-for-free/

            CreateExcelDocument();
        }

        /// <summary>
        /// Creates a new XLS document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-xls-document-net-csharp-vb.php
        /// </remarks>

        static void CreateExcelDocument()
        {
            // Set a path to our Document
            string outFile = @"..\..\..\Result.xls";

            // Create a new document
            ExcelDocument excelDocument = new ExcelDocument();

            // Add several worksheets
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

            // Add the formula
            worksheet.Cells["A3"].Value = "This is the formula";
            worksheet.Cells["B3"].Formula = "=RAND()";

            // Add external and internal links
            worksheet.Cells["A4"].Value = "These are hyperlinks";
            worksheet.Cells["B4"].Value = "External link";
            worksheet.Cells["B4"].Hyperlink = new ExcelHyperlink { Location = "https://sautinsoft.com" };
            worksheet.Cells["C4"].Value = "Internal link";
            worksheet.Cells["C4"].Hyperlink = new ExcelHyperlink { Location = "worksheet2!A1" };

            // Add the current time
            worksheet.Cells["A5"].Value = "This is DateTime";
            worksheet.Cells["B5"].Value = DateTime.Now;

            // Add a large composite text with formatting
            // Create a container of strings
            RichText text = new RichText();
            var part = new RichTextString("This is a very long string... ", new RichTextFormat { Italic = true, Bold = true, FontColor = SKColors.Blue });
            var part2 = new RichTextString("Which have several styles ",
                new RichTextFormat
                {
                    Italic = true,
                    Bold = true,
                    FontColor = SKColors.Green,
                    FontName = "Century",
                    FontSize = 20.2,
                });
            var part3 = new RichTextString("This is superscript text", new RichTextFormat { Strikethrough = true, Superscript = true, FontSize = 18 });
            var part4 = new RichTextString("This is subscript text", new RichTextFormat { Subscript = true, FontSize = 18 });
            
            // Add the following lines to the container
            text.Add(part);
            text.Add(part2);
            text.Add(part3);
            text.Add(part4);
            
            // Add the container to the cell
            worksheet.Cells["A6"].Value = text;

            // Print the properties of the document in a line and color it in a beautiful color
            worksheet.Cells["A8"].Value = $"This worksheet has name \"{worksheet.Name}\", uses {worksheet.Rows.Count} rows and {worksheet.CalculateMaxUsedColumns()} columns";
            worksheet.Cells["A8"].Style.Borders.SetBorders(MultipleBorders.Outside, SKColors.Cyan, LineStyle.Medium);
            worksheet.Cells["A8"].Style.Fill.SetSolid(SKColors.PaleTurquoise);

            // Add a string with numeric formatting
            worksheet.Cells["A9"].Value = .23451;
            worksheet.Cells["A9"].Style.NumberFormat = "#.##%";

            // Expand the columns to make them look attractive
            worksheet.Columns["A"].AutoFit();
            worksheet.Columns["B"].AutoFit();
            worksheet.Columns["C"].AutoFit();

            // Create a copy of the main page
            worksheet.InsertCopy("Just a copy worksheet", worksheet);

            // Save the document
            excelDocument.Save(outFile, new XlsSaveOptions());

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}