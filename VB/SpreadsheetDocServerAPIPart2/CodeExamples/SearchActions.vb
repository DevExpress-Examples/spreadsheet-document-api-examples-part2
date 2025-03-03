Imports DevExpress.Spreadsheet
Imports System
Imports System.Collections.Generic
Imports System.Drawing

Namespace SpreadsheetDocServerAPIPart2
	Public Module SearchActions
		Public SimpleSearchValueAction As Action(Of Workbook) = AddressOf SimpleSearchValue
		Public AdvancedSearchValueAction As Action(Of Workbook) = AddressOf AdvancedSearchValue

		Private Sub SimpleSearchValue(ByVal workbook As Workbook)
'			#Region "#SimpleSearch"
			workbook.Calculate()
			Dim worksheet As Worksheet = workbook.Worksheets("ExpenseReport")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Find and highlight cells that contain the word "holiday".
			Dim searchResult As IEnumerable(Of Cell) = worksheet.Search("holiday")
			For Each cell As Cell In searchResult
				cell.Fill.BackgroundColor = Color.LightGreen
			Next cell
'			#End Region ' #SimpleSearch
		End Sub

		Private Sub AdvancedSearchValue(ByVal workbook As Workbook)
'			#Region "#AdvancedSearch"
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

			' Find and highlight all cells that contain today's date.
			Dim searchResult As IEnumerable(Of Cell) = worksheet.Search(searchString, options)
			For Each cell As Cell In searchResult
				cell.Fill.BackgroundColor = Color.LightGreen
			Next cell
'			#End Region ' #AdvancedSearch
		End Sub
	End Module
End Namespace
