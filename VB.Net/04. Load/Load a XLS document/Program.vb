Imports System
Imports System.IO
Imports SautinSoft.Excel

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' Get your free key here:   
			' https://sautinsoft.com/start-for-free/

			LoadXlsFromFile()
			'LoadXlsFromStream();
		End Sub

		''' <summary>
		''' Loads a XLS document into ExcelDocument from a file.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/load-xls-document-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub LoadXlsFromFile()
			Dim filePath As String = "..\..\..\example.xls"
			' The file format is detected automatically from the file extension: ".xls".
			Dim excel As ExcelDocument = ExcelDocument.Load(filePath, New LoadOptions() With {.Format = FileFormat.Xls})

			If excel IsNot Nothing Then
				Console.WriteLine("Loaded successfully!")
			End If

			Console.ReadKey()
		End Sub

		''' <summary>
		''' Loads a XLS document into ExcelDocument from a MemoryStream.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/excel/help/net/developer-guide/load-xls-document-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub LoadXlsFromStream()
			' Assume that we already have a XLS document as bytes array.
			Dim fileBytes() As Byte = File.ReadAllBytes("..\..\..\example.xls")

			Dim dc As ExcelDocument = Nothing

			' Create a MemoryStream
			Using ms As New MemoryStream(fileBytes)
				' Load a document from the MemoryStream.
				' Specifying LoadOptions we explicitly set that a loadable document is .xls.
				dc = ExcelDocument.Load(ms, New LoadOptions() With {.Format = FileFormat.Xls})
			End Using
			If dc IsNot Nothing Then
				Console.WriteLine("Loaded successfully!")
			End If

			Console.ReadKey()
		End Sub
	End Class
End Namespace
