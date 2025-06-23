using SautinSoft.Excel;
using System.IO;
using SkiaSharp;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free key here:   
            // https://sautinsoft.com/start-for-free/

            InsertImage();
            //InsertImageFromStream();
            //InsertImageWithAnchorCells();
            //InsertImageFromStreamWithAnchorCells();
        }

        /// <summary>
        /// Create xlsx file with an image inside.
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/excel/help/net/developer-guide/drawings/insert-images-in-excel-csharp-vb.php
        /// </remarks>
        static void InsertImage()
        {
            string image = @"..\..\..\cup.jpg";
            string outFile = @"..\..\..\Result.xlsx";

            ExcelDocument excelDocument = new ExcelDocument();

            excelDocument.Worksheets.Add("Page 1");
            var worksheet = excelDocument.Worksheets["Page 1"];

            // Insert an image
            worksheet.Drawings.Add(image, new SautinSoft.Excel.Drawing.Rectangle(0, 0, 1080, 960));

            excelDocument.Save(outFile);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }

        static void InsertImageFromStream()
        {
            string image = @"..\..\..\cup.jpg";
            string outFile = @"..\..\..\Excel Image.xlsx";

            ExcelDocument excelDocument = new ExcelDocument();

            excelDocument.Worksheets.Add("Page 1");
            var worksheet = excelDocument.Worksheets["Page 1"];

            // Insert an image from a stream
            byte[] imageInBytes = File.ReadAllBytes(image);
            using (var streamImage = new MemoryStream(imageInBytes))
            {
                worksheet.Drawings.Add(streamImage, new SautinSoft.Excel.Drawing.Rectangle(0, 0, 1080, 960), ExcelPictureFormat.Jpeg);
                excelDocument.Save(outFile);
            }

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }

        static void InsertImageWithAnchorCells()
        {
            string image = @"..\..\..\cup.jpg";
            string outFile = @"..\..\..\Excel Image.xlsx";

            ExcelDocument excelDocument = new ExcelDocument();

            excelDocument.Worksheets.Add("Page 1");
            var worksheet = excelDocument.Worksheets["Page 1"];

            // Insert an image anchored to cells
            worksheet.Drawings.Add(image, PositionOption.FreeFloating, new AnchorCell(worksheet.Columns[6], worksheet.Rows[6], true),
                new AnchorCell(worksheet.Columns[20], worksheet.Rows[40], true));
            excelDocument.Save(outFile);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }

        static void InsertImageFromStreamWithAnchorCells()
        {
            string image = @"..\..\..\cup.jpg";
            string outFile = @"..\..\..\Excel Image.xlsx";

            ExcelDocument excelDocument = new ExcelDocument();

            excelDocument.Worksheets.Add("Page 1");
            var worksheet = excelDocument.Worksheets["Page 1"];

            // Insert an image from a stream, anchored to cells
            byte[] imageInBytes = File.ReadAllBytes(image);
            using (var streamImage = new MemoryStream(imageInBytes))
            {
                worksheet.Drawings.Add(streamImage, PositionOption.MoveAndSize, new AnchorCell(worksheet.Columns[6], worksheet.Rows[6], true),
                    new AnchorCell(worksheet.Columns[20], worksheet.Rows[40], true), ExcelPictureFormat.Jpeg);
                excelDocument.Save(outFile);
            }


            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}