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

			AllTypesOfText()
		End Sub

		''' <summary>
		''' Using various methods of inserting text into cells.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/using-text-xlsx-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub AllTypesOfText()
			Dim outFile As String = "..\..\..\Result.xlsx"
			' The file format is detected automatically from the file extension: ".xlsx".
			Dim excel As New ExcelDocument()

			' Add an empty worksheet to the file.
			excel.Worksheets.Add("Page 1")
			Dim worksheet = excel.Worksheets("Page 1")


			' This is a regular string.
			worksheet.Cells("A1").Value = "Hello, World!"


			' This is a string with a calculation in C#.
			worksheet.Cells("A2").Value = $"2+2*2 equals {2 + 2 * 2}"
			worksheet.Cells("B2").Value = 2+2*2


			' This is a string created using StringBuilder.
			Dim stringBuilder As New StringBuilder("Hello")
			stringBuilder.Append(" World")
			stringBuilder.Insert(5, ",")
			worksheet.Cells("A3").Value = stringBuilder


			' This is a RichText string with varied formatting.
			Dim wholeString = New RichText()
			Dim part1 = New RichTextString("Hello", New RichTextFormat() With {
				.Italic = True,
				.FontColor = SKColors.Blue
			})
			Dim part2 = New RichTextString(", ", New RichTextFormat() With {.FontColor = SKColors.Red})
			Dim part3 = New RichTextString("World!", New RichTextFormat() With {
				.Bold = True,
				.FontSize = 18,
				.FontColor = SKColors.Green
			})

			wholeString.Add(part1)
			wholeString.Add(part2)
			wholeString.Add(part3)
			worksheet.Cells("A4").Value = wholeString

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
