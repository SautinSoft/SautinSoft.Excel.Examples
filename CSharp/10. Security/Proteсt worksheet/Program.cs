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

            ProtectWorksheet();
        }

        /// <summary>
        /// Protect worksheet in the file.
        /// </summary>
		/// <remarks>
        /// Details: 
        /// </remarks>
        static void ProtectWorksheet()
        {
            string inpFile = @"..\..\..\Example.xlsx";
            string outFile = @"..\..\..\Result.xlsx";

            ExcelDocument excelDocument = ExcelDocument.Load(inpFile);
            excelDocument.Worksheets[0].Protected = true;
            excelDocument.Save(outFile);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}