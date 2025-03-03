Imports DevExpress.Spreadsheet
Imports DevExpress.XtraSpreadsheet.Export
Imports System
Imports System.IO

Namespace SpreadsheetDocServerAPIPart2
	Public Module ExportActions
		Public ExportDocToHTMLAction As Action(Of Workbook) = AddressOf ExportDocToHTML
		Private Sub ExportDocToHTML(ByVal workbook As Workbook)
'			#Region "#ExportToHTML"
			Dim worksheet As Worksheet = workbook.Worksheets("Grouping")
			workbook.Worksheets.ActiveWorksheet = worksheet

			Dim options As New HtmlDocumentExporterOptions()

			' Specify the cell range you want to save as HTML.
			options.SheetIndex = worksheet.Index
			options.Range = "B2:G7"

			' Export data to HTML format.
			Using htmlStream As New FileStream("OutputWorksheet.html", FileMode.Create)
				workbook.ExportToHtml(htmlStream, options)
			End Using

			System.Diagnostics.Process.Start("OutputWorksheet.html")

'			#End Region ' #ExportToHTML
		End Sub
	End Module
End Namespace
