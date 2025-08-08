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
        /// Find and Replace Text.
        /// </summary>
		/// <remarks>
        /// Details: 
        /// </remarks>
        static void Sample()
        {
            string inpFile = @"..\..\..\Example.xlsx";
            string outFile = @"..\..\..\Result.xlsx";

            ExcelDocument excelDocument = ExcelDocument.Load(inpFile);
            CellRange range = excelDocument.Worksheets[0].Cells.GetSubrange("A2", "C9");
            range.FindText("Random", false, true, out int row, out int col);
            if (row > -1 && col > -1) excelDocument.Worksheets[0].Cells[row, col].Value = "Replace";
            excelDocument.Save(outFile);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}