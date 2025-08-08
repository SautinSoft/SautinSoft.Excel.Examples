Imports SautinSoft.Excel
Imports System.IO

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' Get your free key here:   
			' https://sautinsoft.com/start-for-free/

			ProtectWorksheet()
		End Sub

		''' <summary>
		''' Protect worksheet in the file.
		''' </summary>
		''' <remarks>
		''' Details: 
		''' </remarks>
		Private Shared Sub ProtectWorksheet()
			Dim inpFile As String = "..\..\..\Example.xlsx"
			Dim outFile As String = "..\..\..\Result.xlsx"

			Dim excelDocument As ExcelDocument = ExcelDocument.Load(inpFile)
			excelDocument.Worksheets(0).Protected = True
			excelDocument.Save(outFile)

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
