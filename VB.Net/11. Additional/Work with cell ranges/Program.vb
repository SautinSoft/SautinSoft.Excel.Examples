Imports SautinSoft.Excel
Imports SkiaSharp
Imports System.IO

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' Get your free key here:   
			' https://sautinsoft.com/start-for-free/

			Sample()
		End Sub

		''' <summary>
		''' Change style properties in the cell range.
		''' </summary>
		''' <remarks>
		''' Details: 
		''' </remarks>
		Private Shared Sub Sample()
			Dim inpFile As String = "..\..\..\Example.xlsx"
			Dim outFile As String = "..\..\..\Result.xlsx"

			Dim excelDocument As ExcelDocument = ExcelDocument.Load(inpFile)
			For Each cell As ExcelCell In excelDocument.Worksheets(0).Cells.GetSubrange("E6", "F13")
				cell.Style.Fill.SetSolid(SKColors.Yellow)
			Next cell
			excelDocument.Save(outFile)

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
