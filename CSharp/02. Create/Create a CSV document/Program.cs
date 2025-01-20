using SautinSoft.Excel;
using System;
using System.Collections.Generic;

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
        /// Creates a new CSV document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-csv-document-net-csharp-vb.php
        /// </remarks>

        static void CreateExcelDocument()
        {
            // Set a path to our Document
            string outFile = @"..\..\..\Result.csv";

            // Create a new document
            ExcelDocument excelDocument = new ExcelDocument();

            // Add a worksheet
            excelDocument.Worksheets.Add("The main worksheet");

            // Create a variable to address
            var worksheet = excelDocument.Worksheets["The main worksheet"];

            // First option for entering data
            List<List<object>> data = new List<List<object>>() {
                new List<object> { "Date", "Product", "Category", "Quantity", "Unit Price", "Total Cost" },
                new List<object> { new DateOnly(2024, 12, 1).ToString("yyyy-MM-dd"), "Apples", "Fruits", 15, 1.2, "=D2*E2" },
                new List<object> { new DateOnly(2024, 12, 1).ToString("yyyy-MM-dd"), "Bread", "Bakery", 10, 0.8, "=D3*E3" },
                new List<object> { new DateOnly(2024, 12, 2).ToString("yyyy-MM-dd"), "Milk", "Dairy", 20, 1.5, "=D4*E4" },
                new List<object> { new DateOnly(2024, 12, 2).ToString("yyyy-MM-dd"), "Oranges", "Fruits", 10, 1.8, "=D5*E5" },
                new List<object> { new DateOnly(2024, 12, 3).ToString("yyyy-MM-dd"), "Chocolates", "Sweets", 5, 2.5, "=D6*E6" },
                new List<object> { new DateOnly(2024, 12, 3).ToString("yyyy-MM-dd"), "Potatoes", "Vegetables", 25, 0.5, "=D7*E7" },
            };

            int i = 1;
            foreach (var row in data)
            {
                int j = 0;
                foreach (var item in row)
                {
                    worksheet.Cells["ABCDEF"[j] + i.ToString()].Value = item;
                    j++;
                }
                i++;
            }

            // Second option for entering data
            //worksheet.Cells["A1"].Value = "Date";
            //worksheet.Cells["B1"].Value = "Product";
            //worksheet.Cells["C1"].Value = "Category";
            //worksheet.Cells["D1"].Value = "Quantity";
            //worksheet.Cells["E1"].Value = "Unit Price";
            //worksheet.Cells["F1"].Value = "Total Cost";

            //worksheet.Cells["A2"].Value = new DateOnly(2024, 12, 1).ToString("yyyy-MM-dd");
            //worksheet.Cells["B2"].Value = "Apples";
            //worksheet.Cells["C2"].Value = "Fruits";
            //worksheet.Cells["D2"].Value = 15;
            //worksheet.Cells["E2"].Value = 1.2;
            //worksheet.Cells["F2"].Formula = "=D2*E2";

            //worksheet.Cells["A3"].Value = new DateOnly(2024, 12, 1).ToString("yyyy-MM-dd");
            //worksheet.Cells["B3"].Value = "Bread";
            //worksheet.Cells["C3"].Value = "Bakery";
            //worksheet.Cells["D3"].Value = 10;
            //worksheet.Cells["E3"].Value = 0.8;
            //worksheet.Cells["F3"].Formula = "=D3*E3";

            //worksheet.Cells["A4"].Value = new DateOnly(2024, 12, 2).ToString("yyyy-MM-dd");
            //worksheet.Cells["B4"].Value = "Milk";
            //worksheet.Cells["C4"].Value = "Dairy";
            //worksheet.Cells["D4"].Value = 20;
            //worksheet.Cells["E4"].Value = 1.5;
            //worksheet.Cells["F4"].Formula = "=D4*E4";

            //worksheet.Cells["A5"].Value = new DateOnly(2024, 12, 2).ToString("yyyy-MM-dd");
            //worksheet.Cells["B5"].Value = "Oranges";
            //worksheet.Cells["C5"].Value = "Fruits";
            //worksheet.Cells["D5"].Value = 10;
            //worksheet.Cells["E5"].Value = 1.8;
            //worksheet.Cells["F5"].Formula = "=D5*E5";

            //worksheet.Cells["A6"].Value = new DateOnly(2024, 12, 3).ToString("yyyy-MM-dd");
            //worksheet.Cells["B6"].Value = "Chocolate";
            //worksheet.Cells["C6"].Value = "Sweets";
            //worksheet.Cells["D6"].Value = 5;
            //worksheet.Cells["E6"].Value = 2.5;
            //worksheet.Cells["F6"].Formula = "=D6*E6"; 

            //worksheet.Cells["A7"].Value = new DateOnly(2024, 12, 3).ToString("yyyy-MM-dd");
            //worksheet.Cells["B7"].Value = "Potatoes";
            //worksheet.Cells["C7"].Value = "Vegetables";
            //worksheet.Cells["D7"].Value = 25;
            //worksheet.Cells["E7"].Value = 0.5;
            //worksheet.Cells["F7"].Formula = "=D7*E7";

            // Save the document
            excelDocument.Save(outFile, new CsvSaveOptions() { Separator = ',' });

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}