Imports DevExpress.Spreadsheet
Imports DevExpress.XtraSpreadsheet.Export
Imports System.IO

Namespace SpreadsheetDocServerAPIPart2

    Public Module ExportActions

        Private Sub ExportDocToHTML(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#ExportToHTML"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Grouping")
            workbook.Worksheets.ActiveWorksheet = worksheet
            Dim options As DevExpress.XtraSpreadsheet.Export.HtmlDocumentExporterOptions = New DevExpress.XtraSpreadsheet.Export.HtmlDocumentExporterOptions()
            ' Specify the cell range you want to save as HTML.
            options.SheetIndex = worksheet.Index
            options.Range = "B2:G7"
            ' Export data to HTML format.
            Using htmlStream As System.IO.FileStream = New System.IO.FileStream("OutputWorksheet.html", System.IO.FileMode.Create)
                workbook.ExportToHtml(htmlStream, options)
            End Using

            System.Diagnostics.Process.Start("OutputWorksheet.html")
#End Region  ' #ExportToHTML
        End Sub
    End Module
End Namespace
