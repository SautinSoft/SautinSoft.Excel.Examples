using SautinSoft.Excel;
using System.IO;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
			// Before starting, we recommend to get a free key:
            // https://sautinsoft.com/start-for-free/
            
            // Apply the key here:
			// SautinSoft.Excel.ExcelDocument.SetLicense("1234567890");
                  
            // Place your serial(s) number.
            // You will get own serial number(s) after purchasing the license.
            // If you will have any questions, email us to sales@sautinsoft.com or ask at online chat https://www.sautinsoft.com.

            
            string inpFile = @"..\..\..\test.xlsx";
            string outFile = @"..\..\..\Result.docx";

            ExcelDocument excelDocument = ExcelDocument.Load(inpFile);
            excelDocument.Save(outFile, new DocxSaveOptions());

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });

        }
    }
}
