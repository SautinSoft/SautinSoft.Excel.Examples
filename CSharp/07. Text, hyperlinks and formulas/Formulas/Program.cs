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

            VariousFormulas();
        }

        /// <summary>
        /// Using various methods of inserting formulas into cells.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/using-formulas-xlsx-net-csharp-vb.php
        /// </remarks>
        static void VariousFormulas()
        {
            string outFile = @"..\..\..\Result.xlsx";
            // The file format is detected automatically from the file extension: ".xlsx".
            ExcelDocument excel = new ExcelDocument();

            // Sample data
            List<List<object>> data = new List<List<object>>()
            {
                new List<object> { "ID", "Value1", "Value2", "Category", "Date", "Factor1", "Factor2", "Status" },
                new List<object> { 1, 25, 100, "A", "2024-12-01", 1.5, 2.0, "Completed" },
                new List<object> { 2, 40, 200, "B", "2024-12-02", 0.8, 1.1, "Pending" },
                new List<object> { 3, 15, 300, "A", "2024-12-03", 1.2, 1.5, "Completed" },
                new List<object> { 4, 55, 400, "C", "2024-12-04", 2.0, 1.8, "In Progress" },
                new List<object> { 5, 30, 500, "B", "2024-12-05", 1.1, 1.3, "Completed" },
                new List<object> { 6, 45, 600, "C", "2024-12-06", 1.3, 1.7, "Pending" },
                new List<object> { 7, 50, 700, "A", "2024-12-07", 2.5, 1.9, "In Progress" },
                new List<object> { 8, 20, 800, "B", "2024-12-08", 0.7, 2.1, "Completed" },
                new List<object> { 9, 35, 900, "C", "2024-12-09", 1.4, 1.6, "Pending" },
                new List<object> { 10, 60, 1000, "A", "2024-12-10", 3.0, 2.2, "Completed" }
            };


            // Add an empty worksheet to the file.
            excel.Worksheets.Add("Page 1");
            var worksheet = excel.Worksheets["Page 1"];

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


            // Various formulas.
            worksheet.Cells["A14"].Value = "FORMULAS";
            worksheet.Cells["A15"].Formula = "=B2 + C2";
            worksheet.Cells["B15"].Formula = "=AVERAGE(B2:B11)";
            worksheet.Cells["C15"].Formula = "=IF(D2=\"A\", \"Category A\", \"Other\")";
            worksheet.Cells["D15"].Formula = "=COUNTIF(H2:H11, \"Completed\")";
            worksheet.Cells["E15"].Formula = "=SUMIF(D2:D11, \"A\", B2:B11)";
            worksheet.Cells["F15"].Formula = "=COUNTIF(E2:E11, \">2024-12-05\")";
            worksheet.Cells["G15"].Formula = "=AVERAGEIFS(B2:B11, H2:H11, \"Completed\", D2:D11, \"B\")";
            worksheet.Cells["H15"].Formula = "=SUMPRODUCT(F2:F11, G2:G11)";
            worksheet.Cells["I15"].Formula = "=COUNTA(UNIQUE(D2:D11))";
            worksheet.Cells["J15"].Formula = "=SUMIFS(C2:C11, H2:H11, \"Completed\", E2:E11, \">2024-12-05\")";
            worksheet.Cells["K15"].Formula = "=SUMIFS(B2:B11, H2:H11, \"Completed\", D2:D11, \"A\") / SUMIFS(F2:F11, H2:H11, \"Completed\", D2:D11, \"A\")\r\n";

            // Expand the columns to make them look attractive
            worksheet.Columns["E"].AutoFit();
            worksheet.Columns["H"].AutoFit();

            // Saving the excel document.
            excel.Save(outFile);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}