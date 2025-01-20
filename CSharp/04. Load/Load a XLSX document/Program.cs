using System;
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

            LoadXlsxFromFile();
            //LoadXlsxFromStream();
        }

        /// <summary>
        /// Loads a XLSX document into ExcelDocument from a file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/load-xlsx-document-net-csharp-vb.php
        /// </remarks>
        static void LoadXlsxFromFile()
        {
            string filePath = @"..\..\..\example.xlsx";
            // The file format is detected automatically from the file extension: ".xlsx".
            ExcelDocument excel = ExcelDocument.Load(filePath);

            if (excel != null)
                Console.WriteLine("Loaded successfully!");

            Console.ReadKey();
        }

        /// <summary>
        /// Loads a XLSX document into ExcelDocument from a MemoryStream.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/load-xlsx-document-net-csharp-vb.php
        /// </remarks>
        static void LoadXlsxFromStream()
        {
            // Assume that we already have a XLSX document as bytes array.
            byte[] fileBytes = File.ReadAllBytes(@"..\..\..\example.xlsx");

            ExcelDocument dc = null;

            // Create a MemoryStream
            using (MemoryStream ms = new MemoryStream(fileBytes))
            {
                // Load a document from the MemoryStream.
                // Specifying LoadOptions we explicitly set that a loadable document is .xlsx.
                dc = ExcelDocument.Load(ms, new LoadOptions() { Format = FileFormat.Xlsx});
            }
            if (dc != null)
                Console.WriteLine("Loaded successfully!");

            Console.ReadKey();
        }
    }
}