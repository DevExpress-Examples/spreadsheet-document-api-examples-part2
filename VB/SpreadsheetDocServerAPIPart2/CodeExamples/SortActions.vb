Imports DevExpress.Spreadsheet
Imports System
Imports System.Collections.Generic

Namespace SpreadsheetDocServerAPIPart2
	Public Module SortActions
		Public SimpleSortAction As Action(Of Workbook) = AddressOf SimpleSort
		Public DescendingOrderAction As Action(Of Workbook) = AddressOf DescendingOrder
		Public SortBySpecifiedColumnAction As Action(Of Workbook) = AddressOf SortBySpecifiedColumn
		Public SortByMultipleColumnsAction As Action(Of Workbook) = AddressOf SortByMultipleColumns
		Public SortByFillColorAction As Action(Of Workbook) = AddressOf SortByFillColor
		Public SortByFontColorAction As Action(Of Workbook) = AddressOf SortByFontColor
		Private Sub SimpleSort(ByVal workbook As Workbook)
'			#Region "#SimpleSort"
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Fill in the range.
			worksheet.Cells("A2").Value = "Donald Dozier Bradley"
			worksheet.Cells("A3").Value = "Tony Charles Mccallum-Geteer"
			worksheet.Cells("A4").Value = "Calvin Liu"
			worksheet.Cells("A5").Value = "Anita A Boyd"
			worksheet.Cells("A6").Value = "Angela R. Scott"
			worksheet.Cells("A7").Value = "D Fox"

			' Sort the "A2:A7" range in ascending order.
			Dim range As CellRange = worksheet.Range("A2:A7")
			worksheet.Sort(range)

			' Create a heading.
			Dim header As CellRange = worksheet.Range("A1")
			header(0).Value = "Ascending order"
			header.ColumnWidthInCharacters = 30
			header.Style = workbook.Styles("Heading 1")
'			#End Region ' #SimpleSort
		End Sub

		Private Sub DescendingOrder(ByVal workbook As Workbook)
'			#Region "#DescendingOrder"
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Fill in the range.
			worksheet.Cells("A2").Value = "Donald Dozier Bradley"
			worksheet.Cells("A3").Value = "Tony Charles Mccallum-Geteer"
			worksheet.Cells("A4").Value = "Calvin Liu"
			worksheet.Cells("A5").Value = "Anita A Boyd"
			worksheet.Cells("A6").Value = "Angela R. Scott"
			worksheet.Cells("A7").Value = "D Fox"

			' Sort the "A2:A7" range in descending order.
			Dim range As CellRange = worksheet.Range("A2:A7")
			worksheet.Sort(range, False)

			' Create a heading.
			Dim header As CellRange = worksheet.Range("A1")
			header(0).Value = "Descending order"
			header.ColumnWidthInCharacters = 30
			header.Style = workbook.Styles("Heading 1")
'			#End Region ' #DescendingOrder
		End Sub

		Private Sub SortBySpecifiedColumn(ByVal workbook As Workbook)
'			#Region "#SortBySpecifiedColumn"
			Dim worksheet As Worksheet = workbook.Worksheets("SortSample")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Sort the "A3:F22" range by column "D" in ascending order.
			Dim range As CellRange = worksheet.Range("A3:F22")
			worksheet.Sort(range, 3, True)

'			#End Region ' #SortBySpecifiedColumn
		End Sub

		Private Sub SortByMultipleColumns(ByVal workbook As Workbook)
'			#Region "#SortByMultipleColumns"
			Dim worksheet As Worksheet = workbook.Worksheets("SortSample")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Create sorting fields.
			Dim fields As New List(Of SortFieldBase)()

			' Create the first sorting field.
			Dim sortField1 As New SortField()
			' Sort a cell range by column "A" in ascending order.
			sortField1.ColumnOffset = 0
			sortField1.Comparer = worksheet.Comparers.Ascending
			fields.Add(sortField1)

			' Create the second sorting field.
			Dim sortField2 As New SortField()
			' Sort a cell range by column "B" in ascending order.
			sortField2.ColumnOffset = 1
			sortField2.Comparer = worksheet.Comparers.Ascending
			fields.Add(sortField2)

			' Sort the "A3:F22" cell range by sorting fields.
			Dim range As CellRange = worksheet.Range("A3:F22")
			worksheet.Sort(range, fields)

'			#End Region ' #SortByMultipleColumns
		End Sub

		Private Sub SortByFillColor(ByVal workbook As Workbook)
'			#Region "#SortByFillColor"
			Dim worksheet As Worksheet = workbook.Worksheets("SortSample")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Sort the "A3:F22" range by column "A" in ascending order.
			Dim range As CellRange = worksheet.Range("A3:F22")
			worksheet.Sort(range, 0, worksheet("A3").Fill)

'			#End Region ' #SortByFillColor
		End Sub

		Private Sub SortByFontColor(ByVal workbook As Workbook)
'			#Region "#SortByFontColor"
			Dim worksheet As Worksheet = workbook.Worksheets("SortSample")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Sort the "A3:F22" range by column "F" in ascending order.
			Dim range As CellRange = worksheet.Range("A3:F22")
			worksheet.Sort(range,5, worksheet("F12").Font.Color)

'			#End Region ' #SortByFontColor
		End Sub


	End Module
End Namespace
