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

            AllTypesOfText();
        }

        /// <summary>
        /// Using various methods of inserting text into cells.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/using-text-xlsx-net-csharp-vb.php
        /// </remarks>
        static void AllTypesOfText()
        {
            string outFile = @"..\..\..\Result.xlsx";
            // The file format is detected automatically from the file extension: ".xlsx".
            ExcelDocument excel = new ExcelDocument();
            
            // Add an empty worksheet to the file.
            excel.Worksheets.Add("Page 1");
            var worksheet = excel.Worksheets["Page 1"];


            // This is a regular string.
            worksheet.Cells["A1"].Value = "Hello, World!";
            

            // This is a string with a calculation in C#.
            worksheet.Cells["A2"].Value = $"2+2*2 equals {2 + 2 * 2}";
            worksheet.Cells["B2"].Value = 2+2*2;

            
            // This is a string created using StringBuilder.
            var stringBuilder = new StringBuilder("Hello");
            stringBuilder.Append(" World");
            stringBuilder.Insert(5, ",");
            worksheet.Cells["A3"].Value = stringBuilder;


            // This is a RichText string with varied formatting.
            var wholeString = new RichText();
            var part1 = new RichTextString("Hello", new RichTextFormat() { Italic = true, FontColor = SKColors.Blue });
            var part2 = new RichTextString(", ", new RichTextFormat() { FontColor = SKColors.Red });
            var part3 = new RichTextString("World!", new RichTextFormat() { Bold = true, FontSize = 18, FontColor = SKColors.Green });

            wholeString.Add(part1); wholeString.Add(part2); wholeString.Add(part3);
            worksheet.Cells["A4"].Value = wholeString;

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