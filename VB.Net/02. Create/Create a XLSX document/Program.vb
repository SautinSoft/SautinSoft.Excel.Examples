Option Infer On

Imports SautinSoft.Excel
Imports SkiaSharp
Imports System
Imports System.Text

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' Get your free key here:   
			' https://sautinsoft.com/start-for-free/

			CreateExcelDocument()
		End Sub

		''' <summary>
		''' Creates a new XLSX document.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-xlsx-document-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub CreateExcelDocument()
			' Set a path to our Document
			Dim outFile As String = "..\..\..\Result.xlsx"

			' Create a new document
			Dim excelDocument As New ExcelDocument()

			' Add several worksheets
			excelDocument.Worksheets.Add("The main worksheet")
			excelDocument.Worksheets.Add("Second worksheet")

			' Create a variable to address
			Dim worksheet = excelDocument.Worksheets("The main worksheet")

			' Add plain text
			worksheet.Cells("A1").Value = "This is common string"
			worksheet.Cells("B1").Value = "Hello, World! 12345"

			' Add the result of  the expression
			worksheet.Cells("A2").Value = "This is the result of a mathematical expression in C#"
			worksheet.Cells("B2").Value = 5 + 5

			' Add the formula
			worksheet.Cells("A3").Value = "This is the formula"
			worksheet.Cells("B3").Formula = "=RAND()"

			' Add external and internal links
			worksheet.Cells("A4").Value = "These are hyperlinks"
			worksheet.Cells("B4").Value = "External link"
			worksheet.Cells("B4").Hyperlink = New ExcelHyperlink With {.Location = "https://sautinsoft.com"}
			worksheet.Cells("C4").Value = "Internal link"
			worksheet.Cells("C4").Hyperlink = New ExcelHyperlink With {.Location = "worksheet2!A1"}

			' Add the current time
			worksheet.Cells("A5").Value = "This is DateTime"
			worksheet.Cells("B5").Value = DateTime.Now

			' Add a large composite text with formatting
			' Create a container of strings
			Dim text As New RichText()
			Dim part = New RichTextString("This is a very long string... ", New RichTextFormat With {
				.Italic = True,
				.Bold = True,
				.FontColor = SKColors.Blue
			})
			Dim part2 = New RichTextString("Which have several styles ", New RichTextFormat With {
				.Italic = True,
				.Bold = True,
				.FontColor = SKColors.Green,
				.FontName = "Century",
				.FontSize = 20.2
			})
			Dim part3 = New RichTextString("This is superscript text", New RichTextFormat With {
				.Strikethrough = True,
				.Superscript = True,
				.FontSize = 18
			})
			Dim part4 = New RichTextString("This is subscript text", New RichTextFormat With {
				.Subscript = True,
				.FontSize = 18
			})

			' Add the following lines to the container
			text.Add(part)
			text.Add(part2)
			text.Add(part3)
			text.Add(part4)

			' Add the container to the cell
			worksheet.Cells("A6").Value = text

			' Create a string with StringBuilder
			Dim stringBuilder As New StringBuilder("Hello, World!")
			stringBuilder.Append(" From StringBuilder!")
			worksheet.Cells("A7").Value = stringBuilder

			' Print the properties of the document in a line and color it in a beautiful color
			worksheet.Cells("A8").Value = $"This worksheet has name ""{worksheet.Name}"", uses {worksheet.Rows.Count} rows and {worksheet.CalculateMaxUsedColumns()} columns"
			worksheet.Cells("A8").Style.Borders.SetBorders(MultipleBorders.Outside, SKColors.Cyan, LineStyle.Medium)
			worksheet.Cells("A8").Style.Fill.SetSolid(SKColors.PaleTurquoise)

			' Add a string with numeric formatting
			worksheet.Cells("A9").Value = .23451
			worksheet.Cells("A9").Style.NumberFormat = "#.##%"

			' Create a comment for it
			worksheet.Cells("A9").Comment.Text = "This is formatted string"
			worksheet.Cells("A9").Comment.Author = "Alex"

			' Expand the columns to make them look attractive
			worksheet.Columns("A").AutoFit()
			worksheet.Columns("B").AutoFit()
			worksheet.Columns("C").AutoFit()

			' Create a copy of the main page
			worksheet.InsertCopy("Just a copy worksheet", worksheet)

			' Save the document
			excelDocument.Save(outFile, New XlsxSaveOptions())

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
