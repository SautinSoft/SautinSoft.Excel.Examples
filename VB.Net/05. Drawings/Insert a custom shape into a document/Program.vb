Option Infer On

Imports SautinSoft.Excel
Imports System.IO
Imports SkiaSharp

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' Get your free key here:   
			' https://sautinsoft.com/start-for-free/

			InsertShape()
		End Sub

		''' <summary>
		''' Create xlsx file with a shape inside.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/excel/help/net/developer-guide/insert-images-in-excel-csharp-vb.php
		''' </remarks>
		Private Shared Sub InsertShape()
			Dim outFile As String = "..\..\..\Result.xlsx"

			Dim excelDocument As New ExcelDocument()

			excelDocument.Worksheets.Add("Page 1")
			Dim worksheet = excelDocument.Worksheets("Page 1")

			' Insert a shape
			Dim [property] As New ShapeProperty()
			Dim shape As New ExcelShape([property])
			[property].Fill.SetSolid(SKColors.Red)
			[property].Outline.Fill.SetSolid(SKColors.Black)

			Dim custom = [property].Geometry.SetCustom()
			Dim path = custom.AddPath(New SautinSoft.Excel.Drawing.Size(200, 200))
			path.MoveTo(New SautinSoft.Excel.Drawing.Point(100, 50))
			path.AddCubicBezier(New SautinSoft.Excel.Drawing.Point(50, 0), New SautinSoft.Excel.Drawing.Point(0, 50), New SautinSoft.Excel.Drawing.Point(100, 200))
			path.AddCubicBezier(New SautinSoft.Excel.Drawing.Point(200, 50), New SautinSoft.Excel.Drawing.Point(150, 0), New SautinSoft.Excel.Drawing.Point(100, 50))
			path.ClosePath()

			worksheet.Drawings.Add(shape)
			shape.BoundingRectangle = New SautinSoft.Excel.Drawing.Rectangle(0, 0, 200, 200)

			excelDocument.Save(outFile)

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
