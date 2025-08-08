using SautinSoft.Excel;
using SkiaSharp;
using System.IO;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free key here:   
            // https://sautinsoft.com/start-for-free/

            Sample();
        }

        /// <summary>
        /// Change style properties in the cell range.
        /// </summary>
		/// <remarks>
        /// Details: 
        /// </remarks>
        static void Sample()
        {
            string inpFile = @"..\..\..\Example.xlsx";
            string outFile = @"..\..\..\Result.xlsx";

            ExcelDocument excelDocument = ExcelDocument.Load(inpFile);
            foreach (ExcelCell cell in excelDocument.Worksheets[0].Cells.GetSubrange("E6", "F13"))
            {
                cell.Style.Fill.SetSolid(SKColors.Yellow);
            }
            excelDocument.Save(outFile);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}