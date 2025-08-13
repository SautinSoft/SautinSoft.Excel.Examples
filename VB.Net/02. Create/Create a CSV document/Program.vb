Option Infer On

Imports SautinSoft.Excel
Imports System
Imports System.Collections.Generic

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' Get your free key here:   
			' https://sautinsoft.com/start-for-free/

			CreateExcelDocument()
		End Sub

		''' <summary>
		''' Creates a new CSV document.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-csv-document-net-csharp-vb.php
		''' </remarks>

		Private Shared Sub CreateExcelDocument()
			' Set a path to our Document
			Dim outFile As String = "..\..\..\Result.csv"

			' Create a new document
			Dim excelDocument As New ExcelDocument()

			' Add a worksheet
			excelDocument.Worksheets.Add("The main worksheet")

			' Create a variable to address
			Dim worksheet = excelDocument.Worksheets("The main worksheet")

			' First option for entering data
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

			Dim i As Integer = 1
			For Each row In data
				Dim j As Integer = 0
				For Each item In row
					worksheet.Cells("ABCDEF".Chars(j) + i.ToString()).Value = item
					j += 1
				Next item
				i += 1
			Next row

			' Second option for entering data
			'worksheet.Cells["A1"].Value = "Date";
			'worksheet.Cells["B1"].Value = "Product";
			'worksheet.Cells["C1"].Value = "Category";
			'worksheet.Cells["D1"].Value = "Quantity";
			'worksheet.Cells["E1"].Value = "Unit Price";
			'worksheet.Cells["F1"].Value = "Total Cost";

			'worksheet.Cells["A2"].Value = new DateOnly(2024, 12, 1).ToString("yyyy-MM-dd");
			'worksheet.Cells["B2"].Value = "Apples";
			'worksheet.Cells["C2"].Value = "Fruits";
			'worksheet.Cells["D2"].Value = 15;
			'worksheet.Cells["E2"].Value = 1.2;
			'worksheet.Cells["F2"].Formula = "=D2*E2";

			'worksheet.Cells["A3"].Value = new DateOnly(2024, 12, 1).ToString("yyyy-MM-dd");
			'worksheet.Cells["B3"].Value = "Bread";
			'worksheet.Cells["C3"].Value = "Bakery";
			'worksheet.Cells["D3"].Value = 10;
			'worksheet.Cells["E3"].Value = 0.8;
			'worksheet.Cells["F3"].Formula = "=D3*E3";

			'worksheet.Cells["A4"].Value = new DateOnly(2024, 12, 2).ToString("yyyy-MM-dd");
			'worksheet.Cells["B4"].Value = "Milk";
			'worksheet.Cells["C4"].Value = "Dairy";
			'worksheet.Cells["D4"].Value = 20;
			'worksheet.Cells["E4"].Value = 1.5;
			'worksheet.Cells["F4"].Formula = "=D4*E4";

			'worksheet.Cells["A5"].Value = new DateOnly(2024, 12, 2).ToString("yyyy-MM-dd");
			'worksheet.Cells["B5"].Value = "Oranges";
			'worksheet.Cells["C5"].Value = "Fruits";
			'worksheet.Cells["D5"].Value = 10;
			'worksheet.Cells["E5"].Value = 1.8;
			'worksheet.Cells["F5"].Formula = "=D5*E5";

			'worksheet.Cells["A6"].Value = new DateOnly(2024, 12, 3).ToString("yyyy-MM-dd");
			'worksheet.Cells["B6"].Value = "Chocolate";
			'worksheet.Cells["C6"].Value = "Sweets";
			'worksheet.Cells["D6"].Value = 5;
			'worksheet.Cells["E6"].Value = 2.5;
			'worksheet.Cells["F6"].Formula = "=D6*E6"; 

			'worksheet.Cells["A7"].Value = new DateOnly(2024, 12, 3).ToString("yyyy-MM-dd");
			'worksheet.Cells["B7"].Value = "Potatoes";
			'worksheet.Cells["C7"].Value = "Vegetables";
			'worksheet.Cells["D7"].Value = 25;
			'worksheet.Cells["E7"].Value = 0.5;
			'worksheet.Cells["F7"].Formula = "=D7*E7";

			' Save the document
			excelDocument.Save(outFile, New CsvSaveOptions() With {.Separator = ","c})

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
