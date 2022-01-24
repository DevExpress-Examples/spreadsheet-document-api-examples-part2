Imports DevExpress.Spreadsheet
Imports System
Imports System.Collections.Generic
Imports System.Drawing

Namespace SpreadsheetDocServerAPIPart2

    Public Module SearchActions

        Private Sub SimpleSearchValue(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#SimpleSearch"
            workbook.Calculate()
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("ExpenseReport")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Find and highlight cells that contain the word "holiday".
            Dim searchResult As System.Collections.Generic.IEnumerable(Of DevExpress.Spreadsheet.Cell) = worksheet.Search("holiday")
            For Each cell As DevExpress.Spreadsheet.Cell In searchResult
                cell.Fill.BackgroundColor = System.Drawing.Color.LightGreen
            Next
#End Region  ' #SimpleSearch
        End Sub

        Private Sub AdvancedSearchValue(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#AdvancedSearch"
            workbook.Calculate()
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("ExpenseReport")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Specify the search term.
            Dim searchString As String = System.DateTime.Today.ToString("d")
            ' Specify search options.
            Dim options As DevExpress.Spreadsheet.SearchOptions = New DevExpress.Spreadsheet.SearchOptions()
            options.SearchBy = DevExpress.Spreadsheet.SearchBy.Columns
            options.SearchIn = DevExpress.Spreadsheet.SearchIn.Values
            options.MatchEntireCellContents = True
            ' Find and highlight all cells that contain today's date.
            Dim searchResult As System.Collections.Generic.IEnumerable(Of DevExpress.Spreadsheet.Cell) = worksheet.Search(searchString, options)
            For Each cell As DevExpress.Spreadsheet.Cell In searchResult
                cell.Fill.BackgroundColor = System.Drawing.Color.LightGreen
            Next
#End Region  ' #AdvancedSearch
        End Sub
    End Module
End Namespace
