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

            LoadXlsFromFile();
            //LoadXlsFromStream();
        }

        /// <summary>
        /// Loads a XLS document into ExcelDocument from a file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/load-xls-document-net-csharp-vb.php
        /// </remarks>
        static void LoadXlsFromFile()
        {
            string filePath = @"..\..\..\example.xls";
            // The file format is detected automatically from the file extension: ".xls".
            ExcelDocument excel = ExcelDocument.Load(filePath, new LoadOptions() { Format = FileFormat.Xls});

            if (excel != null)
                Console.WriteLine("Loaded successfully!");

            Console.ReadKey();
        }

        /// <summary>
        /// Loads a XLS document into ExcelDocument from a MemoryStream.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/load-xls-document-net-csharp-vb.php
        /// </remarks>
        static void LoadXlsFromStream()
        {
            // Assume that we already have a XLS document as bytes array.
            byte[] fileBytes = File.ReadAllBytes(@"..\..\..\example.xls");

            ExcelDocument dc = null;

            // Create a MemoryStream
            using (MemoryStream ms = new MemoryStream(fileBytes))
            {
                // Load a document from the MemoryStream.
                // Specifying LoadOptions we explicitly set that a loadable document is .xls.
                dc = ExcelDocument.Load(ms, new LoadOptions() { Format = FileFormat.Xls });
            }
            if (dc != null)
                Console.WriteLine("Loaded successfully!");

            Console.ReadKey();
        }
    }
}