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

			CreateHeadersFooters()
		End Sub

		''' <summary>
		''' Create Headers and Footers in Excel Document.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/create-headers-footers-xlsx-document-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub CreateHeadersFooters()
			Dim outFile As String = "..\..\..\example.xlsx"
			' The file format is detected automatically from the file extension: ".xlsx".
			Dim excel As New ExcelDocument()

			' Add an empty worksheet to the file
			excel.Worksheets.Add("Page 1")
			Dim worksheet = excel.Worksheets("Page 1")

			' Add different headers
			worksheet.HeadersFooters.Header = "Header"
			worksheet.HeadersFooters.FirstHeader = "FirstHeader"
			worksheet.HeadersFooters.EvenHeader = "EvenHeader"

			' Add different footers
			worksheet.HeadersFooters.Footer = "Footer"
			worksheet.HeadersFooters.FirstFooter = "FirstFooter"
			worksheet.HeadersFooters.EvenFooter = "EvenFooter"

			' Set the settings for the first or even headers and footers
			worksheet.HeadersFooters.DifferentFirst = True
			worksheet.HeadersFooters.DifferentOddEven= True

			' Sample data
			Dim data As New List(Of List(Of Object))() _
				From {
					New List(Of Object) From {"Date", "Product", "Category", "Quantity", "Unit Price", "Total Cost"},
					New List(Of Object) From {(New DateTime(2024, 12, 1)).ToString("yyyy-MM-dd"), "Apples", "Fruits", 15, 1.2, "=D2*E2"},
					New List(Of Object) From {(New DateTime(2024, 12, 1)).ToString("yyyy-MM-dd"), "Bread", "Bakery", 10, 0.8, "=D3*E3"},
					New List(Of Object) From {(New DateTime(2024, 12, 2)).ToString("yyyy-MM-dd"), "Milk", "Dairy", 20, 1.5, "=D4*E4"},
					New List(Of Object) From {(New DateTime(2024, 12, 2)).ToString("yyyy-MM-dd"), "Oranges", "Fruits", 10, 1.8, "=D5*E5"},
					New List(Of Object) From {(New DateTime(2024, 12, 3)).ToString("yyyy-MM-dd"), "Chocolates", "Sweets", 5, 2.5, "=D6*E6"},
					New List(Of Object) From {(New DateTime(2024, 12, 3)).ToString("yyyy-MM-dd"), "Potatoes", "Vegetables", 25, 0.5, "=D7*E7"}
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
