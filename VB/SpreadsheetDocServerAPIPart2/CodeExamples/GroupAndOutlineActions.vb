Imports DevExpress.Spreadsheet
Imports System
Imports System.Collections.Generic

Namespace SpreadsheetDocServerAPIPart2
	Public Module GroupAndOutlineActions
		Public GroupRowsAction As Action(Of Workbook) = AddressOf GroupRows
		Public GroupColumnsAction As Action(Of Workbook) = AddressOf GroupColumns
		Public UngroupRowsAction As Action(Of Workbook) = AddressOf UngroupRows
		Public UngroupColumnsAction As Action(Of Workbook) = AddressOf UngroupColumns
		Public AutoOutlineAction As Action(Of Workbook) = AddressOf AutoOutline
		Public SubtotalAction As Action(Of Workbook) = AddressOf Subtotal

		Private Sub GroupRows(ByVal workbook As IWorkbook)
'			#Region "#GroupRows    "
			Dim worksheet As Worksheet = workbook.Worksheets("Grouping")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Group rows 3 through 6 and collapse the group.
			worksheet.Rows.Group(2, 5, True)

			' Group rows 9 through 12 and expand the group.
			worksheet.Rows.Group(8, 11, False)

			' Group rows 2 through 13 to create the outer group. 
			worksheet.Rows.Group(1, 12, False)
'			#End Region ' #GroupRows
		End Sub

		Private Sub GroupColumns(ByVal workbook As IWorkbook)
'			#Region "#GroupColumns"
			Dim worksheet As Worksheet = workbook.Worksheets("Grouping")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Group columns "C" through "F" and expand the group.
			worksheet.Columns.Group(2, 5, False)
'			#End Region ' #GroupColumns
		End Sub

		Private Sub UngroupRows(ByVal workbook As IWorkbook)
'			#Region "#UngroupRows    "
			Dim worksheet As Worksheet = workbook.Worksheets("Grouping and Outline")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Ungroup rows 3 through 6 and display collapsed data.
			worksheet.Rows.UnGroup(2, 5, True)

			' Ungroup rows 9 through 12.
			worksheet.Rows.UnGroup(8, 11, False)

			' Remove the outer row group.
			worksheet.Rows.UnGroup(1, 12, False)
'			#End Region ' #UngroupRows
		End Sub

		Private Sub UngroupColumns(ByVal workbook As IWorkbook)
'			#Region "#UngroupColumns"
			Dim worksheet As Worksheet = workbook.Worksheets("Grouping and Outline")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Ungroup columns "C" through "F".
			worksheet.Columns.UnGroup(2, 5, False)
'			#End Region ' #UngroupColumns
		End Sub

		Private Sub AutoOutline(ByVal workbook As IWorkbook)
'			#Region "#AutoOutline"
			Dim worksheet As Worksheet = workbook.Worksheets("Grouping")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Outline data automatically based on the summary formulas.
			worksheet.AutoOutline()
'			#End Region ' #AutoOutline
		End Sub

		Private Sub Subtotal(ByVal workbook As IWorkbook)
'			#Region "#Subtotal"
			Dim worksheet As Worksheet = workbook.Worksheets("Regional Sales")
			workbook.Worksheets.ActiveWorksheet = worksheet
			' Obtain the target cell range.
			Dim dataRange As CellRange = worksheet("B3:E23")
			' Calculate subtotals for column "D".
			Dim subtotalColumnsList As New List(Of Integer)()
			subtotalColumnsList.Add(3)
			' Insert subtotals by each change in column "B"
			' and calculate the SUM fuction for the related rows in column "D".
			worksheet.Subtotal(dataRange, 1, subtotalColumnsList, 9, "Total")
'			#End Region ' #Subtotal
		End Sub
	End Module
End Namespace
