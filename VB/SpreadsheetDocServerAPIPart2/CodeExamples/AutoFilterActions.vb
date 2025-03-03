Imports DevExpress.Spreadsheet
Imports System
Imports System.Collections.Generic
Imports System.Drawing

Namespace SpreadsheetDocServerAPIPart2
	Public Module AutoFilterActions

		Public ApplyFilterAction As Action(Of Workbook) = AddressOf ApplyFilter
		Public FilterAndSortBySingleColumnAction As Action(Of Workbook) = AddressOf FilterAndSortBySingleColumn
		Public FilterAndSortByMultipleColumnsAction As Action(Of Workbook) = AddressOf FilterAndSortByMultipleColumns
		Public FilterNumericByConditionAction As Action(Of Workbook) = AddressOf FilterNumericByCondition
		Public FilterTextByConditionAction As Action(Of Workbook) = AddressOf FilterTextByCondition
		Public FilterByValueAction As Action(Of Workbook) = AddressOf FilterByValue
		Public FilterByMultipleValuesAction As Action(Of Workbook) = AddressOf FilterByMultipleValues
		Public FilterDatesByConditionAction As Action(Of Workbook) = AddressOf FilterDatesByCondition
		Public FilterMixedDataTypesByValuesAction As Action(Of Workbook) = AddressOf FilterMixedDataTypesByValues
		Public Top10FilterValueAction As Action(Of Workbook) = AddressOf Top10FilterValue
		Public DynamicFilterValueAction As Action(Of Workbook) = AddressOf DynamicFilterValue
		Public FilterAndSortByColorAction As Action(Of Workbook) = AddressOf FilterAndSortByColor
		Public FilterByBackgroundColorAction As Action(Of Workbook) = AddressOf FilterByBackgroundColor
		Public FilterByFillColorAction As Action(Of Workbook) = AddressOf FilterByFillColor
		Public FilterByFontColorAction As Action(Of Workbook) = AddressOf FilterByFontColor
		Public ReapplyFilterValueAction As Action(Of Workbook) = AddressOf ReapplyFilterValue
		Public ClearFilterAction As Action(Of Workbook) = AddressOf ClearFilter
		Public DisableFilterAction As Action(Of Workbook) = AddressOf DisableFilter

		Private Sub ApplyFilter(ByVal workbook As Workbook)
'			#Region "#ApplyFilter"
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)
'			#End Region ' #ApplyFilter
		End Sub

		Private Sub FilterAndSortBySingleColumn(ByVal workbook As Workbook)
'			#Region "#FilterAndSortBySingleColumn"
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Sort data in the "B2:E23" range
			' in descending order by column "A".
			worksheet.AutoFilter.SortState.Sort(0, True)
'			#End Region ' #FilterAndSortBySingleColumn
		End Sub

		Private Sub FilterAndSortByMultipleColumns(ByVal workbook As Workbook)
'			#Region "#FilterAndSortByMultipleColumns"
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Sort data in the "B2:E23" range
			' in descending order by columns "A" and "C".
			Dim sortConditions As New List(Of SortCondition)()
			Dim color As Color = worksheet("D12").Font.Color

			sortConditions.Add(New SortCondition(0, True))
			sortConditions.Add(New SortCondition(2, color, False))
			worksheet.AutoFilter.SortState.Sort(sortConditions)
'			#End Region ' #FilterAndSortByMultipleColumns
		End Sub

		Private Sub FilterNumericByCondition(ByVal workbook As Workbook)
'			#Region "#FilterNumbersByCondition"
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Filter values in the "Sales" column that are in a range from 5000$ to 8000$.
			Dim sales As AutoFilterColumn = worksheet.AutoFilter.Columns(2)
			sales.ApplyCustomFilter(5000, FilterComparisonOperator.GreaterThanOrEqual, 8000, FilterComparisonOperator.LessThanOrEqual, True)
'			#End Region ' #FilterNumbersByCondition
		End Sub

		Private Sub FilterTextByCondition(ByVal workbook As Workbook)
'			#Region "#FilterTextByCondition"
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Filter values in the "Product" column that contain "Gi" and include empty cells.
			Dim products As AutoFilterColumn = worksheet.AutoFilter.Columns(1)
			products.ApplyCustomFilter("*Gi*", FilterComparisonOperator.Equal, FilterValue.FilterByBlank, FilterComparisonOperator.Equal, False)
'			#End Region ' #FilterTextByCondition
		End Sub

		Private Sub FilterByValue(ByVal workbook As Workbook)
'			#Region "#FilterBySingleValue"
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Filter data in the "Product" column by a specific value.
			worksheet.AutoFilter.Columns(1).ApplyFilterCriteria("Mozzarella di Giovanni")
'			#End Region ' #FilterBySingleValue
		End Sub

		Private Sub FilterByMultipleValues(ByVal workbook As Workbook)
'			#Region "#FilterByMultipleValues"
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Filter data in the "Product" column by an array of values.
			worksheet.AutoFilter.Columns(1).ApplyFilterCriteria(New CellValue() { "Mozzarella di Giovanni", "Gorgonzola Telino" })
'			#End Region ' #FilterByMultipleValues
		End Sub

		Private Sub FilterDatesByCondition(ByVal workbook As Workbook)
'			#Region "#FilterDatesByCondition    "
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Filter values in the "Reported Date" column
			' to display dates that are between June 1, 2014 and February 1, 2015.
			worksheet.AutoFilter.Columns(3).ApplyCustomFilter(New Date(2014, 6, 1), FilterComparisonOperator.GreaterThanOrEqual, New Date(2015, 2, 1), FilterComparisonOperator.LessThanOrEqual, True)
'			#End Region ' #FilterDatesByCondition
		End Sub

		Private Sub FilterMixedDataTypesByValues(ByVal workbook As Workbook)
'			#Region "#FilterMixedDataByValues    "
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Create date grouping item to filter January 2015 dates.
			Dim groupings As IList(Of DateGrouping) = New List(Of DateGrouping)()
			Dim dateGroupingJan2015 As New DateGrouping(New Date(2015, 1, 1), DateTimeGroupingType.Month)
			groupings.Add(dateGroupingJan2015)

			' Filter data in the "Reported Date" column
			' to display values reported in January 2015.
			worksheet.AutoFilter.Columns(3).ApplyFilterCriteria("gennaio 2015", groupings)
'			#End Region ' #FilterMixedDataByValues
		End Sub

		Private Sub Top10FilterValue(ByVal workbook As Workbook)
'			#Region "#TopTenFilter    "
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Apply a filter to the "Sales" column to display the top ten values.
			worksheet.AutoFilter.Columns(2).ApplyTop10Filter(Top10Type.Top10Items, 10)
'			#End Region ' #TopTenFilter
		End Sub

		Private Sub DynamicFilterValue(ByVal workbook As Workbook)
'			#Region "#DynamicFilter"
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Apply a dynamic filter to the "Sales" column
			' to display only values that are above the average.
			worksheet.AutoFilter.Columns(2).ApplyDynamicFilter(DynamicFilterType.AboveAverage)
			' Apply a dynamic filter to the "Reported Date" column
			' to display values reported this year.
			worksheet.AutoFilter.Columns(3).ApplyDynamicFilter(DynamicFilterType.ThisYear)
'			#End Region ' #DynamicFilter
		End Sub


		Private Sub FilterAndSortByColor(ByVal workbook As Workbook)
'			#Region "#FilterAndSortByColor"
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Sort data in the "B2:E23" range
			' in descending order by column "D".
			Dim color As Color = worksheet("D12").Font.Color
			worksheet.AutoFilter.SortState.Sort(2, color, False)
'			#End Region ' #FilterAndSortByColor
		End Sub

		Private Sub FilterByBackgroundColor(ByVal workbook As Workbook)
'			#Region "#FilterByBackgroundColor"
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Filter values in the "Products" column by background color.
			Dim products As AutoFilterColumn = worksheet.AutoFilter.Columns(1)
			products.ApplyFillColorFilter(worksheet("C12").FillColor)
'			#End Region ' #FilterByBackgroundColor
		End Sub

		Private Sub FilterByFillColor(ByVal workbook As Workbook)
'			#Region "#FilterByFillColor"
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Filter values in the "Products" column by fill color.
			Dim products As AutoFilterColumn = worksheet.AutoFilter.Columns(1)
			products.ApplyFillFilter(worksheet("C10").Fill)
'			#End Region ' #FilterByFillColor
		End Sub

		Private Sub FilterByFontColor(ByVal workbook As Workbook)
'			#Region "#FilterByFontColor"
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Filter values in the "Sales" column by font color.
			Dim products As AutoFilterColumn = worksheet.AutoFilter.Columns(2)
			products.ApplyFontColorFilter(worksheet("D10").Font.Color)
'			#End Region ' #FilterByFontColor
		End Sub

		Private Sub ReapplyFilterValue(ByVal workbook As Workbook)
'			#Region "#ReapplyFilter    "
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Filter values in the "Sales" column that are greater than 5000$.
			worksheet.AutoFilter.Columns(2).ApplyCustomFilter(5000, FilterComparisonOperator.GreaterThan)

			' Change data and reapply the filter.
			worksheet("D3").Value = 5000
			worksheet.AutoFilter.ReApply()
'			#End Region ' #ReapplyFilter
		End Sub

		Private Sub ClearFilter(ByVal workbook As Workbook)
'			#Region "#ClearFilter"
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Filter values in the "Sales" column that are greater than 5000$.
			worksheet.AutoFilter.Columns(2).ApplyCustomFilter(5000, FilterComparisonOperator.GreaterThan)

			' Clear the filter.
			worksheet.AutoFilter.Clear()
'			#End Region ' #ClearFilter
		End Sub

		Private Sub DisableFilter(ByVal workbook As Workbook)
'			#Region "#DisableFilter"
			Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Enable filtering for the "B2:E23" cell range.
			Dim range As CellRange = worksheet("B2:E23")
			worksheet.AutoFilter.Apply(range)

			' Disable filtering for the entire worksheet.
			worksheet.AutoFilter.Disable()
'			#End Region ' #DisableFilter
		End Sub
	End Module
End Namespace
