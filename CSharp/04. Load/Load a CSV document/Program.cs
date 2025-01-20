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

            LoadCsvFromFile();
            //LoadCsvFromStream();
        }

        /// <summary>
        /// Loads a CSV document into ExcelDocument from a file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/load-csv-document-net-csharp-vb.php
        /// </remarks>
        static void LoadCsvFromFile()
        {
            string filePath = @"..\..\..\example.csv";
            // The file format is detected automatically from the file extension: ".csv".
            ExcelDocument excel = ExcelDocument.Load(filePath);

            if (excel != null)
                Console.WriteLine("Loaded successfully!");

            Console.ReadKey();
        }

        /// <summary>
        /// Loads a CSV document into ExcelDocument from a MemoryStream.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/load-csv-document-net-csharp-vb.php
        /// </remarks>
        static void LoadCsvFromStream()
        {
            // Assume that we already have a CSV document as bytes array.
            byte[] fileBytes = File.ReadAllBytes(@"..\..\..\example.csv");

            ExcelDocument dc = null;

            // Create a MemoryStream
            using (MemoryStream ms = new MemoryStream(fileBytes))
            {
                // Load a document from the MemoryStream.
                // Specifying LoadOptions we explicitly set that a loadable document is .csv.
                dc = ExcelDocument.Load(ms, new LoadOptions() { Format = FileFormat.Csv });
            }
            if (dc != null)
                Console.WriteLine("Loaded successfully!");

            Console.ReadKey();
        }
    }
}