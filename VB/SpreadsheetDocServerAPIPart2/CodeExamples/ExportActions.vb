Imports DevExpress.Spreadsheet
Imports DevExpress.XtraSpreadsheet.Export
Imports System.IO

Namespace SpreadsheetDocServerAPIPart2
    Public NotInheritable Class ExportActions

        Private Sub New()
        End Sub

        Private Shared Sub ExportDocToHTML(ByVal workbook As Workbook)
'            #Region "#ExportToHTML"
            Dim worksheet As Worksheet = workbook.Worksheets("Grouping")
            workbook.Worksheets.ActiveWorksheet = worksheet

            Dim options As New HtmlDocumentExporterOptions()

            ' Specify the part of the document to be exported to HTML.
            options.SheetIndex = worksheet.Index
            options.Range = "B2:G7"

            ' Export the active worksheet to a stream as HTML with the specified options.
            Using htmlStream As New FileStream("OutputWorksheet.html", FileMode.Create)
                workbook.ExportToHtml(htmlStream, options)
            End Using

            System.Diagnostics.Process.Start("OutputWorksheet.html")

'            #End Region ' #ExportToHTML
        End Sub
    End Class
End Namespace
