Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports SautinSoft.Excel

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' Get your free key here:   
			' https://sautinsoft.com/start-for-free/

			CreateExcelFontsSizeFromFile()

		End Sub

		''' <summary>
		''' Create a XLSX document with different fonts and size into ExcelDocument .
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/create-xlsx-font-size-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub CreateExcelFontsSizeFromFile()
			Dim outFile As String = "..\..\..\Result.xlsx"
			' The file format is detected automatically from the file extension: ".xlsx".
			Dim excel As New ExcelDocument()

			' Set default font name and size
			excel.DefaultFontName = "Segoe UI"
			excel.DefaultFontSize = 20

			' Add an empty worksheet to the file
			excel.Worksheets.Add("Page 1")
			Dim worksheet = excel.Worksheets("Page 1")

			' Sample data
			Dim data As New List(Of List(Of Object))() _
				From {
					New List(Of Object) From {"Date", "Product", "Category", "Quantity", "Unit Price", "Total Cost"},
					New List(Of Object) From {(New DateOnly(2024, 12, 1)).ToString(), "Apples", "Fruits", 15, 1.2, "=D2*E2"},
					New List(Of Object) From {(New DateOnly(2024, 12, 1)).ToString(), "Bread", "Bakery", 10, 0.8, "=D3*E3"},
					New List(Of Object) From {(New DateOnly(2024, 12, 2)).ToString(), "Milk", "Dairy", 20, 1.5, "=D4*E4"},
					New List(Of Object) From {(New DateOnly(2024, 12, 2)).ToString(), "Oranges", "Fruits", 10, 1.8, "=D5*E5"},
					New List(Of Object) From {(New DateOnly(2024, 12, 3)).ToString(), "Chocolates", "Sweets", 5, 2.5, "=D6*E6"},
					New List(Of Object) From {(New DateOnly(2024, 12, 3)).ToString(), "Potatoes", "Vegetables", 25, 0.5, "=D7*E7"}
				}

			' Inserting data
			Dim i As Integer = 1
			For Each row In data
				Dim j As Integer = 0
				For Each item In row
					worksheet.Cells("ABCDEFGHIJKLMNOPQRSTUVWXYZ".Chars(j) + i.ToString()).Value = item
					j += 1
				Next item
				i += 1
			Next row

			' Saving the excel document
			excel.Save(outFile)

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
