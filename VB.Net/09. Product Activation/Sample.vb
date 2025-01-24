Imports SautinSoft.Excel
Imports System.IO

Namespace Sample
    Class Sample
        Shared Sub Main(ByVal args As String())
            ' Before starting, we recommend to get a free key:
            ' https://sautinsoft.com/start-for-free/
            
            ' Apply the key here:
            ' SautinSoft.Excel.ExcelDocument.SetLicense("1234567890")
                  
            ' Place your serial(s) number.
            ' You will get own serial number(s) after purchasing the license.
            ' If you have any questions, email us to sales@sautinsoft.com or ask at online chat https://www.sautinsoft.com.

            Dim inpFile As String = "..\..\..\test.xlsx"
            Dim outFile As String = "..\..\..\Result.docx"

            Dim excelDocument As ExcelDocument = ExcelDocument.Load(inpFile)
            excelDocument.Save(outFile, New DocxSaveOptions())

            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})

        End Sub
    End Class
End Namespace