﻿using SautinSoft.Excel;
using System.IO;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free key here:   
            // https://sautinsoft.com/start-for-free/

            ConvertFromFile();
            ConvertFromStream();
        }

        /// <summary>
        /// Convert CSV to XLSX (file to file).
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/excel/help/net/developer-guide/convert-csv-to-xlsx-in-csharp-vb.php
        /// </remarks>
        static void ConvertFromFile()
        {
            string inpFile = @"..\..\..\Example.csv";
            string outFile = @"..\..\..\Result.xlsx";

            ExcelDocument excelDocument = ExcelDocument.Load(inpFile, new LoadOptions { CsvTryParseNumbers = true });
            excelDocument.Save(outFile, new XlsxSaveOptions());

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }

        /// <summary>
        /// Convert CSV to XLSX (using Stream).
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/excel/help/net/developer-guide/convert-csv-to-xlsx-in-csharp-vb.php
        /// </remarks>
        static void ConvertFromStream()
        {

            // We need files only for demonstration purposes.
            // The conversion process will be done completely in memory.
            string inpFile = @"..\..\..\Example.csv";
            string outFile = @"..\..\..\ResultStream.xlsx";
            byte[] inpData = File.ReadAllBytes(inpFile);
            byte[] outData = null;

            using (MemoryStream msInp = new MemoryStream(inpData))
            {

                // Load a document.
                ExcelDocument excelDocument = ExcelDocument.Load(inpFile, new LoadOptions { CsvTryParseNumbers = true });

                // Save the excel document to XLSX format.
                using (MemoryStream outMs = new MemoryStream())
                {
                    excelDocument.Save(outMs, new XlsxSaveOptions());
                    outData = outMs.ToArray();
                }
                // Show the result for demonstration purposes.
                if (outData != null)
                {
                    File.WriteAllBytes(outFile, outData);

                    // Important for Linux: Install MS Fonts
                    // sudo apt install ttf-mscorefonts-installer -y

                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
                }
            }
        }
    }
}