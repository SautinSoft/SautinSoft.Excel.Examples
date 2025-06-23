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

            InsertShape();
        }

        /// <summary>
        /// Create xlsx file with a shape inside.
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/excel/help/net/developer-guide/drawings/insert-shape-in-excel-csharp-vb.php
        /// </remarks>
        static void InsertShape()
        {
            string outFile = @"..\..\..\Result.xlsx";

            ExcelDocument excelDocument = new ExcelDocument();

            excelDocument.Worksheets.Add("Page 1");
            var worksheet = excelDocument.Worksheets["Page 1"];

            // Insert a shape
            ShapeProperty property = new ShapeProperty();
            ExcelShape shape = new ExcelShape(property);
            property.Fill.SetSolid(SKColors.Blue);
            property.Geometry.SetPreset(SautinSoft.Excel.Drawing.Figure.Ellipse);
            property.Outline.Fill.SetSolid(SKColors.Black);

            worksheet.Drawings.Add(shape);
            shape.BoundingRectangle = new SautinSoft.Excel.Drawing.Rectangle(0, 0, 200, 300);

            excelDocument.Save(outFile);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}