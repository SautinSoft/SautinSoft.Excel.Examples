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

			VariousFormulas()
		End Sub

		''' <summary>
		''' Using various methods of inserting formulas into cells.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/using-formulas-xlsx-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub VariousFormulas()
			Dim outFile As String = "..\..\..\Result.xlsx"
			' The file format is detected automatically from the file extension: ".xlsx".
			Dim excel As New ExcelDocument()

			' Sample data
			Dim data As New List(Of List(Of Object))() _
				From {
					New List(Of Object) From {"ID", "Value1", "Value2", "Category", "Date", "Factor1", "Factor2", "Status"},
					New List(Of Object) From {1, 25, 100, "A", "2024-12-01", 1.5, 2.0, "Completed"},
					New List(Of Object) From {2, 40, 200, "B", "2024-12-02", 0.8, 1.1, "Pending"},
					New List(Of Object) From {3, 15, 300, "A", "2024-12-03", 1.2, 1.5, "Completed"},
					New List(Of Object) From {4, 55, 400, "C", "2024-12-04", 2.0, 1.8, "In Progress"},
					New List(Of Object) From {5, 30, 500, "B", "2024-12-05", 1.1, 1.3, "Completed"},
					New List(Of Object) From {6, 45, 600, "C", "2024-12-06", 1.3, 1.7, "Pending"},
					New List(Of Object) From {7, 50, 700, "A", "2024-12-07", 2.5, 1.9, "In Progress"},
					New List(Of Object) From {8, 20, 800, "B", "2024-12-08", 0.7, 2.1, "Completed"},
					New List(Of Object) From {9, 35, 900, "C", "2024-12-09", 1.4, 1.6, "Pending"},
					New List(Of Object) From {10, 60, 1000, "A", "2024-12-10", 3.0, 2.2, "Completed"}
				}


			' Add an empty worksheet to the file.
			excel.Worksheets.Add("Page 1")
			Dim worksheet = excel.Worksheets("Page 1")

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


			' Various formulas.
			worksheet.Cells("A14").Value = "FORMULAS"
			worksheet.Cells("A15").Formula = "=B2 + C2"
			worksheet.Cells("B15").Formula = "=AVERAGE(B2:B11)"
			worksheet.Cells("C15").Formula = "=IF(D2=""A"", ""Category A"", ""Other"")"
			worksheet.Cells("D15").Formula = "=COUNTIF(H2:H11, ""Completed"")"
			worksheet.Cells("E15").Formula = "=SUMIF(D2:D11, ""A"", B2:B11)"
			worksheet.Cells("F15").Formula = "=COUNTIF(E2:E11, "">2024-12-05"")"
			worksheet.Cells("G15").Formula = "=AVERAGEIFS(B2:B11, H2:H11, ""Completed"", D2:D11, ""B"")"
			worksheet.Cells("H15").Formula = "=SUMPRODUCT(F2:F11, G2:G11)"
			worksheet.Cells("I15").Formula = "=COUNTA(UNIQUE(D2:D11))"
			worksheet.Cells("J15").Formula = "=SUMIFS(C2:C11, H2:H11, ""Completed"", E2:E11, "">2024-12-05"")"
			worksheet.Cells("K15").Formula = "=SUMIFS(B2:B11, H2:H11, ""Completed"", D2:D11, ""A"") / SUMIFS(F2:F11, H2:H11, ""Completed"", D2:D11, ""A"")" & vbCrLf

			' Expand the columns to make them look attractive
			worksheet.Columns("E").AutoFit()
			worksheet.Columns("H").AutoFit()

			' Saving the excel document.
			excel.Save(outFile)

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
