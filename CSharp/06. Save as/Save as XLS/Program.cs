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

            SaveToXlsFile();
            SaveToXlsStream();
        }

        /// <summary>
        /// Creates a new document and saves it as XLS file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/create-save-xls-document.php
        /// </remarks>
        static void SaveToXlsFile()
        {
            // Assume we already have a document.
            ExcelDocument excelDocument = new ExcelDocument();
            excelDocument.Worksheets.Add("New worksheet");

            // Add some text.
            excelDocument.Worksheets["New worksheet"].Cells["A1"].Value = "This is sample ExcelDocument";
            // Format the text
            excelDocument.Worksheets["New worksheet"].Columns["A"].AutoFit();

            string filePath = @"..\..\..\Result.xls";

            // The file format will be detected automatically from the file extension: ".xls".
            excelDocument.Save(filePath);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }

        /// <summary>
        /// Creates a new document and saves it as XLS using MemoryStream.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/create-save-xls-document.php
        /// </remarks>
        static void SaveToXlsStream()
        {
            // There variables are necessary only for demonstration purposes.
            byte[] fileData = null;
            string filePath = @"Result-stream.xls";

            // Assume we already have a document.
            ExcelDocument excelDocument = new ExcelDocument();
            excelDocument.Worksheets.Add("New worksheet");

            // Add some text
            excelDocument.Worksheets["New worksheet"].Cells["A1"].Value = "This is sample ExcelDocument";
            // Format the text
            excelDocument.Worksheets["New worksheet"].Columns["A"].AutoFit();

            // Let's save our document to a MemoryStream.
            using (MemoryStream ms = new MemoryStream())
            {
                // 2nd parameter: we've explicitly set to save our document in XLS format.
                excelDocument.Save(ms, new XlsSaveOptions());
                // Important for Linux: Install MS Fonts
                // sudo apt install ttf-mscorefonts-installer -y

                fileData = ms.ToArray();
            }

            File.WriteAllBytes(filePath, fileData);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}