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

            SaveToDocxFile();
            SaveToDocxStream();
        }

        /// <summary>
        /// Creates a new document and saves it as DOCX file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/create-save-docx-document.php
        /// </remarks>
        static void SaveToDocxFile()
        {
            // Assume we already have a document.
            ExcelDocument excelDocument = new ExcelDocument();
            excelDocument.Worksheets.Add("New worksheet");

            // Add some text.
            excelDocument.Worksheets["New worksheet"].Cells["A1"].Value = "This is sample ExcelDocument";
            // Format the text
            excelDocument.Worksheets["New worksheet"].Columns["A"].AutoFit();

            string filePath = @"..\..\..\Result.docx";

            // The file format will be detected automatically from the file extension: ".docx".
            excelDocument.Save(filePath);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }

        /// <summary>
        /// Creates a new document and saves it as DOCX using MemoryStream.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/create-save-docx-document.php
        /// </remarks>
        static void SaveToDocxStream()
        {
            // There variables are necessary only for demonstration purposes.
            byte[] fileData = null;
            string filePath = @"Result-stream.docx";

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
                // 2nd parameter: we've explicitly set to save our document in DOCX format.
                excelDocument.Save(ms, new DocxSaveOptions());
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