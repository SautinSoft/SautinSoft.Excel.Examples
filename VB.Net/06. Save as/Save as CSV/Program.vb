Imports System.IO
Imports SautinSoft.Excel

Namespace Example
	Friend Class Program

		Shared Sub Main(ByVal args() As String)
			' Get your free key here:   
			' https://sautinsoft.com/start-for-free/

			SaveToCsvFile()
			SaveToCsvStream()
		End Sub

		''' <summary>
		''' Creates a new document and saves it as CSV file.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/create-save-csv-document.php
		''' </remarks>
		Private Shared Sub SaveToCsvFile()
			' Assume we already have a document.
			Dim excelDocument As New ExcelDocument()
			excelDocument.Worksheets.Add("New worksheet")

			' Add some text.
			excelDocument.Worksheets("New worksheet").Cells("A1").Value = "This is sample ExcelDocument"
			' Format the text
			excelDocument.Worksheets("New worksheet").Columns("A").AutoFit()

			Dim filePath As String = "..\..\..\Result.csv"

			' The file format will be detected automatically from the file extension: ".csv".
			excelDocument.Save(filePath)

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
		End Sub

		''' <summary>
		''' Creates a new document and saves it as CSV using MemoryStream.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/create-save-csv-document.php
		''' </remarks>
		Private Shared Sub SaveToCsvStream()
			' There variables are necessary only for demonstration purposes.
			Dim fileData() As Byte = Nothing
			Dim filePath As String = "Result-stream.csv"

			' Assume we already have a document.
			Dim excelDocument As New ExcelDocument()
			excelDocument.Worksheets.Add("New worksheet")

			' Add some text
			excelDocument.Worksheets("New worksheet").Cells("A1").Value = "This is sample ExcelDocument"
			' Format the text
			excelDocument.Worksheets("New worksheet").Columns("A").AutoFit()

			' Let's save our document to a MemoryStream.
			Using ms As New MemoryStream()
				' 2nd parameter: we've explicitly set to save our document in CSV format.
				excelDocument.Save(ms, New CsvSaveOptions())
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
