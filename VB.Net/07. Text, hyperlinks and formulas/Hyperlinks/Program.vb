Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports SautinSoft.Excel
Imports SkiaSharp

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' Get your free key here:   
			' https://sautinsoft.com/start-for-free/

			AllTypesOfHyperlinks()
		End Sub

		''' <summary>
		''' Inserting 2 types of hyperlinks into the cells.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/using-hyperlinks-xlsx-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub AllTypesOfHyperlinks()
			Dim outFile As String = "..\..\..\Result.xlsx"
			' The file format is detected automatically from the file extension: ".xlsx".
			Dim excel As New ExcelDocument()

			' Add an empty worksheet to the file.
			excel.Worksheets.Add("Page 1")
			excel.Worksheets.Add("Page 2")
			Dim worksheet = excel.Worksheets("Page 1")


			' Add hyperlinks into document.
			worksheet.Cells("A1").Value = "External link"
			worksheet.Cells("A1").Hyperlink = New ExcelHyperlink With {
				.Location = "https://sautinsoft.com",
				.ToolTip = "SautinSoft"
			}

			worksheet.Cells("A2").Value = "Internal link"
			worksheet.Cells("A2").Hyperlink = New ExcelHyperlink With {
				.Location = "'Page 2'!A1",
				.ToolTip = "A1 cell on another page"
			}


			' Expand the column to make it look attractive
			worksheet.Columns("A").AutoFit()

			' Saving the excel document.
			excel.Save(outFile)

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
