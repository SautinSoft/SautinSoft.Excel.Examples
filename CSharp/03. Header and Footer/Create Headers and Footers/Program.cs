using System;
using System.Collections.Generic;
using System.IO;
using SautinSoft.Excel;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free key here:   
            // https://sautinsoft.com/start-for-free/

            CreateHeadersFooters();
        }

        /// <summary>
        /// Create Headers and Footers in Excel Document.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/create-headers-footers-xlsx-document-net-csharp-vb.php
        /// </remarks>
        static void CreateHeadersFooters()
        {
            string outFile = @"..\..\..\example.xlsx";
            // The file format is detected automatically from the file extension: ".xlsx".
            ExcelDocument excel = new ExcelDocument();

            // Add an empty worksheet to the file
            excel.Worksheets.Add("Page 1");
            var worksheet = excel.Worksheets["Page 1"];

            // Add different headers
            worksheet.HeadersFooters.Header = "Header";
            worksheet.HeadersFooters.FirstHeader = "FirstHeader";
            worksheet.HeadersFooters.EvenHeader = "EvenHeader";

            // Add different footers
            worksheet.HeadersFooters.Footer = "Footer";
            worksheet.HeadersFooters.FirstFooter = "FirstFooter";
            worksheet.HeadersFooters.EvenFooter = "EvenFooter";

            // Set the settings for the first or even headers and footers
            worksheet.HeadersFooters.DifferentFirst = true;
            worksheet.HeadersFooters.DifferentOddEven= true;

            // Sample data
            List<List<object>> data = new List<List<object>>() {
                new List<object> { "Date", "Product", "Category", "Quantity", "Unit Price", "Total Cost" },
                new List<object> { new DateOnly(2024, 12, 1).ToString(), "Apples", "Fruits", 15, 1.2, "=D2*E2" },
                new List<object> { new DateOnly(2024, 12, 1).ToString(), "Bread", "Bakery", 10, 0.8, "=D3*E3" },
                new List<object> { new DateOnly(2024, 12, 2).ToString(), "Milk", "Dairy", 20, 1.5, "=D4*E4" },
                new List<object> { new DateOnly(2024, 12, 2).ToString(), "Oranges", "Fruits", 10, 1.8, "=D5*E5" },
                new List<object> { new DateOnly(2024, 12, 3).ToString(), "Chocolates", "Sweets", 5, 2.5, "=D6*E6" },
                new List<object> { new DateOnly(2024, 12, 3).ToString(), "Potatoes", "Vegetables", 25, 0.5, "=D7*E7" },
            };

            // Inserting data
            int i = 1;
            foreach (var row in data)
            {
                int j = 0;
                foreach (var item in row)
                {
                    worksheet.Cells["ABCDEFGHIJKLMNOPQRSTUVWXYZ"[j] + i.ToString()].Value = item;
                    j++;
                }
                i++;
            }

            // Saving the excel document
            excel.Save(outFile);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}