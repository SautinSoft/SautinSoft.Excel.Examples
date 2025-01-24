Imports System.IO
Imports SautinSoft.Excel

Namespace Example
	Friend Class Program

		Shared Sub Main(ByVal args() As String)
			' Get your free key here:   
			' https://sautinsoft.com/start-for-free/

			SaveToRtfFile()
			SaveToRtfStream()
		End Sub

		''' <summary>
		''' Creates a new document and saves it as RTF file.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/create-save-rtf-document.php
		''' </remarks>
		Private Shared Sub SaveToRtfFile()
			' Assume we already have a document.
			Dim excelDocument As New ExcelDocument()
			excelDocument.Worksheets.Add("New worksheet")

			' Add some text.
			excelDocument.Worksheets("New worksheet").Cells("A1").Value = "This is sample ExcelDocument"
			' Format the text
			excelDocument.Worksheets("New worksheet").Columns("A").AutoFit()

			Dim filePath As String = "..\..\..\Result.rtf"

			' The file format will be detected automatically from the file extension: ".rtf".
			excelDocument.Save(filePath)

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
		End Sub

		''' <summary>
		''' Creates a new document and saves it as RTF using MemoryStream.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/create-save-rtf-document.php
		''' </remarks>
		Private Shared Sub SaveToRtfStream()
			' There variables are necessary only for demonstration purposes.
			Dim fileData() As Byte = Nothing
			Dim filePath As String = "Result-stream.rtf"

			' Assume we already have a document.
			Dim excelDocument As New ExcelDocument()
			excelDocument.Worksheets.Add("New worksheet")

			' Add some text
			excelDocument.Worksheets("New worksheet").Cells("A1").Value = "This is sample ExcelDocument"
			' Format the text
			excelDocument.Worksheets("New worksheet").Columns("A").AutoFit()

			' Let's save our document to a MemoryStream.
			Using ms As New MemoryStream()
				' 2nd parameter: we've explicitly set to save our document in RTF format.
				excelDocument.Save(ms, New RtfSaveOptions())
				' Important for Linux: Install MS Fonts
				' sudo apt install ttf-mscorefonts-installer -y

				fileData = ms.ToArray()
			End Using

			File.WriteAllBytes(filePath, fileData)

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
