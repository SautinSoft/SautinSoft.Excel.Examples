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
        /// Convert XLSX to XLS (file to file).
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/excel/help/net/developer-guide/convert-xlsx-to-xls-in-csharp-vb.php
        /// </remarks>
        static void ConvertFromFile()
        {
            string inpFile = @"..\..\..\Example.xlsx";
            string outFile = @"..\..\..\Result.xls";

            ExcelDocument excelDocument = ExcelDocument.Load(inpFile);
            excelDocument.Save(outFile, new XlsSaveOptions());

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }

        /// <summary>
        /// Convert XLSX to XLS (using Stream).
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/excel/help/net/developer-guide/convert-xlsx-to-xls-in-csharp-vb.php
        /// </remarks>
        static void ConvertFromStream()
        {

            // We need files only for demonstration purposes.
            // The conversion process will be done completely in memory.
            string inpFile = @"..\..\..\Example.xlsx";
            string outFile = @"..\..\..\ResultStream.xls";
            byte[] inpData = File.ReadAllBytes(inpFile);
            byte[] outData = null;

            using (MemoryStream msInp = new MemoryStream(inpData))
            {

                // Load a document.
                ExcelDocument excelDocument = ExcelDocument.Load(inpFile);

                // Save the excel document to XLS format.
                using (MemoryStream outMs = new MemoryStream())
                {
                    excelDocument.Save(outMs, new XlsSaveOptions());
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