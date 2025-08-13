// Get your free key here:   
// https://sautinsoft.com/start-for-free/


using SautinSoft.Excel;

// Convert Excel to PDF in memory
ExcelDocument x = new ExcelDocument();


string excelFile = @"test.xlsx";
string pdfFile = @"test.pdf";
byte[] pdfBytes;

try
{
    // Let us say, we have a memory stream with Excel data.
    using (MemoryStream ms = new MemoryStream(File.ReadAllBytes(excelFile)))
    {
        x.Save(ms, new PdfSaveOptions());
        pdfBytes = ms.ToArray();
    }
    // Save pdfBytes to a file for demonstration purposes.
    File.WriteAllBytes(pdfFile, pdfBytes);
  //  System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(pdfFile) { UseShellExecute = true });

}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
    Console.ReadLine();
}



