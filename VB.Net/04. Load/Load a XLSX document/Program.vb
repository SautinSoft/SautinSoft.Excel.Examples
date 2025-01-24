Imports System
Imports System.IO
Imports SautinSoft.Excel

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' Get your free key here:   
			' https://sautinsoft.com/start-for-free/

			LoadXlsxFromFile()
			'LoadXlsxFromStream();
		End Sub

		''' <summary>
		''' Loads a XLSX document into ExcelDocument from a file.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/load-xlsx-document-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub LoadXlsxFromFile()
			Dim filePath As String = "..\..\..\example.xlsx"
			' The file format is detected automatically from the file extension: ".xlsx".
			Dim excel As ExcelDocument = ExcelDocument.Load(filePath)

			If excel IsNot Nothing Then
				Console.WriteLine("Loaded successfully!")
			End If

			Console.ReadKey()
		End Sub

		''' <summary>
		''' Loads a XLSX document into ExcelDocument from a MemoryStream.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/load-xlsx-document-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub LoadXlsxFromStream()
			' Assume that we already have a XLSX document as bytes array.
			Dim fileBytes() As Byte = File.ReadAllBytes("..\..\..\example.xlsx")

			Dim dc As ExcelDocument = Nothing

			' Create a MemoryStream
			Using ms As New MemoryStream(fileBytes)
				' Load a document from the MemoryStream.
				' Specifying LoadOptions we explicitly set that a loadable document is .xlsx.
				dc = ExcelDocument.Load(ms, New LoadOptions() With {.Format = FileFormat.Xlsx})
			End Using
			If dc IsNot Nothing Then
				Console.WriteLine("Loaded successfully!")
			End If

			Console.ReadKey()
		End Sub
	End Class
End Namespace
