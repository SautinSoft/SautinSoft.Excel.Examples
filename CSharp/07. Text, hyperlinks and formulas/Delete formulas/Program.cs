using SautinSoft.Excel;
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
        /// Delete formulasin the cell range.
        /// </summary>
		/// <remarks>
        /// Details: 
        /// </remarks>
        static void Sample()
        {
            string inpFile = @"..\..\..\Example.xlsx";
            string outFile = @"..\..\..\Result.xlsx";

            ExcelDocument excelDocument = ExcelDocument.Load(inpFile);
            CellRange range = excelDocument.Worksheets[0].Cells.GetSubrange("A7", "B12");
            foreach (ExcelCell cell in range)
            {
                cell.Formula = null;
            }
            excelDocument.Save(outFile);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}