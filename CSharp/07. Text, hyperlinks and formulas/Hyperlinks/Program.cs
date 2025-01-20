using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using SautinSoft.Excel;
using SkiaSharp;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free key here:   
            // https://sautinsoft.com/start-for-free/

            AllTypesOfHyperlinks();
        }

        /// <summary>
        /// Inserting 2 types of hyperlinks into the cells.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/using-hyperlinks-xlsx-net-csharp-vb.php
        /// </remarks>
        static void AllTypesOfHyperlinks()
        {
            string outFile = @"..\..\..\Result.xlsx";
            // The file format is detected automatically from the file extension: ".xlsx".
            ExcelDocument excel = new ExcelDocument();

            // Add an empty worksheet to the file.
            excel.Worksheets.Add("Page 1");
            excel.Worksheets.Add("Page 2");
            var worksheet = excel.Worksheets["Page 1"];


            // Add hyperlinks into document.
            worksheet.Cells["A1"].Value = "External link";
            worksheet.Cells["A1"].Hyperlink = new ExcelHyperlink { Location = "https://sautinsoft.com", ToolTip = "SautinSoft" };

            worksheet.Cells["A2"].Value = "Internal link";
            worksheet.Cells["A2"].Hyperlink = new ExcelHyperlink { Location = "\'Page 2\'!A1", ToolTip = "A1 cell on another page" };
            

            // Expand the column to make it look attractive
            worksheet.Columns["A"].AutoFit();

            // Saving the excel document.
            excel.Save(outFile);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}