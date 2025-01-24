Imports SautinSoft.Excel
Imports System.IO

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' Get your free key here:   
			' https://sautinsoft.com/start-for-free/

			ConvertFromFile()
			ConvertFromStream()
		End Sub

		''' <summary>
		''' Convert XLSX to RTF (file to file).
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/excel/help/net/developer-guide/convert-xlsx-to-rtf-in-csharp-vb.php
		''' </remarks>
		Private Shared Sub ConvertFromFile()
			Dim inpFile As String = "..\..\..\Example.xlsx"
			Dim outFile As String = "..\..\..\Result.rtf"

			Dim excelDocument As ExcelDocument = ExcelDocument.Load(inpFile)
			excelDocument.Save(outFile, New RtfSaveOptions())

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub

		''' <summary>
		''' Convert XLSX to RTF (using Stream).
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/excel/help/net/developer-guide/convert-xlsx-to-rtf-in-csharp-vb.php
		''' </remarks>
		Private Shared Sub ConvertFromStream()

			' We need files only for demonstration purposes.
			' The conversion process will be done completely in memory.
			Dim inpFile As String = "..\..\..\Example.xlsx"
			Dim outFile As String = "..\..\..\ResultStream.rtf"
			Dim inpData() As Byte = File.ReadAllBytes(inpFile)
			Dim outData() As Byte = Nothing

			Using msInp As New MemoryStream(inpData)

				' Load a document.
				Dim excelDocument As ExcelDocument = ExcelDocument.Load(inpFile)

				' Save the excel document to RTF format.
				Using outMs As New MemoryStream()
					excelDocument.Save(outMs, New RtfSaveOptions())
					outData = outMs.ToArray()
				End Using
				' Show the result for demonstration purposes.
				If outData IsNot Nothing Then
					File.WriteAllBytes(outFile, outData)

					' Important for Linux: Install MS Fonts
					' sudo apt install ttf-mscorefonts-installer -y

					System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
				End If
			End Using
		End Sub
	End Class
End Namespace
