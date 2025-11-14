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
        /// Protect worksheet in the file without passwords.
        /// </summary>
		/// <remarks>
        /// Details: 
        /// </remarks>
        static void ProtectWorksheet()
        {
            string inpFile = @"..\..\..\Example.xlsx";
            string outFile = @"..\..\..\Result.xlsx";

            ExcelDocument excelDocument = ExcelDocument.Load(inpFile);
            // To prevent other users from accidentally or deliberately changing, moving, or deleting data in a worksheet, you can lock the cells on your Excel worksheet and then protect the sheet with a password.
            // Say you own the team status report worksheet, where you want team members to add data in specific cells only and not be able to modify anything else.
            // With worksheet protection, you can make only certain parts of the sheet editable and users will not be able to modify data in any other region in the sheet. 

            excelDocument.Worksheets[0].Protected = true;

            // Using MS Excel just click on File-> Info-> Unprotect.
            excelDocument.Save(outFile);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}