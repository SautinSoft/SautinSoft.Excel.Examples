Option Infer On

Imports SautinSoft.Excel
Imports System.IO
Imports SkiaSharp

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' Get your free key here:   
			' https://sautinsoft.com/start-for-free/

			InsertImage()
			'InsertImageFromStream();
			'InsertImageWithAnchorCells();
			'InsertImageFromStreamWithAnchorCells();
		End Sub

		''' <summary>
		''' Create xlsx file with an image inside.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/excel/help/net/developer-guide/insert-images-in-excel-csharp-vb.php
		''' </remarks>
		Private Shared Sub InsertImage()
			Dim image As String = "..\..\..\cup.jpg"
			Dim outFile As String = "..\..\..\Result.xlsx"

			Dim excelDocument As New ExcelDocument()

			excelDocument.Worksheets.Add("Page 1")
			Dim worksheet = excelDocument.Worksheets("Page 1")

			' Insert an image
			worksheet.Drawings.Add(image, New SautinSoft.Excel.Drawing.Rectangle(0, 0, 1080, 960))

			excelDocument.Save(outFile)

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub

		Private Shared Sub InsertImageFromStream()
			Dim image As String = "..\..\..\cup.jpg"
			Dim outFile As String = "..\..\..\Excel Image.xlsx"

			Dim excelDocument As New ExcelDocument()

			excelDocument.Worksheets.Add("Page 1")
			Dim worksheet = excelDocument.Worksheets("Page 1")

			' Insert an image from a stream
			Dim imageInBytes() As Byte = File.ReadAllBytes(image)
			Using streamImage = New MemoryStream(imageInBytes)
				worksheet.Drawings.Add(streamImage, New SautinSoft.Excel.Drawing.Rectangle(0, 0, 1080, 960), ExcelPictureFormat.Jpeg)
				excelDocument.Save(outFile)
			End Using

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub

		Private Shared Sub InsertImageWithAnchorCells()
			Dim image As String = "..\..\..\cup.jpg"
			Dim outFile As String = "..\..\..\Excel Image.xlsx"

			Dim excelDocument As New ExcelDocument()

			excelDocument.Worksheets.Add("Page 1")
			Dim worksheet = excelDocument.Worksheets("Page 1")

			' Insert an image anchored to cells
			worksheet.Drawings.Add(image, PositionOption.FreeFloating, New AnchorCell(worksheet.Columns(6), worksheet.Rows(6), True), New AnchorCell(worksheet.Columns(20), worksheet.Rows(40), True))
			excelDocument.Save(outFile)

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub

		Private Shared Sub InsertImageFromStreamWithAnchorCells()
			Dim image As String = "..\..\..\cup.jpg"
			Dim outFile As String = "..\..\..\Excel Image.xlsx"

			Dim excelDocument As New ExcelDocument()

			excelDocument.Worksheets.Add("Page 1")
			Dim worksheet = excelDocument.Worksheets("Page 1")

			' Insert an image from a stream, anchored to cells
			Dim imageInBytes() As Byte = File.ReadAllBytes(image)
			Using streamImage = New MemoryStream(imageInBytes)
				worksheet.Drawings.Add(streamImage, PositionOption.MoveAndSize, New AnchorCell(worksheet.Columns(6), worksheet.Rows(6), True), New AnchorCell(worksheet.Columns(20), worksheet.Rows(40), True), ExcelPictureFormat.Jpeg)
				excelDocument.Save(outFile)
			End Using


			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
