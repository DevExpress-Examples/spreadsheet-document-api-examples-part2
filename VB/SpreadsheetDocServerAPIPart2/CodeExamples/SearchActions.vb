Imports DevExpress.Spreadsheet
Imports System
Imports System.Collections.Generic
Imports System.Drawing

Namespace SpreadsheetDocServerAPIPart2
    Public NotInheritable Class SearchActions

        Private Sub New()
        End Sub

        Private Shared Sub SimpleSearchValue(ByVal workbook As Workbook)
'            #Region "#SimpleSearch"
            workbook.Calculate()
            Dim worksheet As Worksheet = workbook.Worksheets("ExpenseReport")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Find and highlight cells containing the word "holiday".
            Dim searchResult As IEnumerable(Of Cell) = worksheet.Search("holiday")
            For Each cell As Cell In searchResult
                cell.Fill.BackgroundColor = Color.LightGreen
            Next cell
'            #End Region ' #SimpleSearch
        End Sub

        Private Shared Sub AdvancedSearchValue(ByVal workbook As Workbook)
'            #Region "#AdvancedSearch"
            workbook.Calculate()
            Dim worksheet As Worksheet = workbook.Worksheets("ExpenseReport")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Specify the search term.
            Dim searchString As String = Date.Today.ToString("d")

            ' Specify search options.
            Dim options As New SearchOptions()
            options.SearchBy = SearchBy.Columns
            options.SearchIn = SearchIn.Values
            options.MatchEntireCellContents = True

            ' Find all cells containing today's date and paint them light-green.
            Dim searchResult As IEnumerable(Of Cell) = worksheet.Search(searchString, options)
            For Each cell As Cell In searchResult
                cell.Fill.BackgroundColor = Color.LightGreen
            Next cell
'            #End Region ' #AdvancedSearch
        End Sub
    End Class
End Namespace
