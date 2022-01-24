Imports DevExpress.Spreadsheet
Imports System
Imports System.Collections.Generic

Namespace SpreadsheetDocServerAPIPart2

    Public Module AutoFilterActions

        Private Sub ApplyFilter(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#ApplyFilter"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Enable filtering for the "B2:E23" cell range.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)
#End Region  ' #ApplyFilter
        End Sub

        Private Sub FilterAndSortBySingleColumn(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#FilterAndSortBySingleColumn"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Enable filtering for the "B2:E23" cell range.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)
            ' Sort data in the "B2:E23" range
            ' in descending order by column "A".
            worksheet.AutoFilter.SortState.Sort(0, True)
#End Region  ' #FilterAndSortBySingleColumn
        End Sub

        Private Sub FilterAndSortByMultipleColumns(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#FilterAndSortByMultipleColumns"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Enable filtering for the "B2:E23" cell range.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)
            ' Sort data in the "B2:E23" range
            ' in descending order by columns "A" and "C".
            Dim sortConditions As System.Collections.Generic.List(Of DevExpress.Spreadsheet.SortCondition) = New System.Collections.Generic.List(Of DevExpress.Spreadsheet.SortCondition)()
            sortConditions.Add(New DevExpress.Spreadsheet.SortCondition(0, True))
            sortConditions.Add(New DevExpress.Spreadsheet.SortCondition(2, True))
            worksheet.AutoFilter.SortState.Sort(sortConditions)
#End Region  ' #FilterAndSortByMultipleColumns
        End Sub

        Private Sub FilterNumericByCondition(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#FilterNumbersByCondition"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Enable filtering for the "B2:E23" cell range.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)
            ' Filter values in the "Sales" column that are in a range from 5000$ to 8000$.
            Dim sales As DevExpress.Spreadsheet.AutoFilterColumn = worksheet.AutoFilter.Columns(2)
            sales.ApplyCustomFilter(5000, DevExpress.Spreadsheet.FilterComparisonOperator.GreaterThanOrEqual, 8000, DevExpress.Spreadsheet.FilterComparisonOperator.LessThanOrEqual, True)
#End Region  ' #FilterNumbersByCondition
        End Sub

        Private Sub FilterTextByCondition(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#FilterTextByCondition"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Enable filtering for the "B2:E23" cell range.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)
            ' Filter values in the "Product" column that contain "Gi" and include empty cells.
            Dim products As DevExpress.Spreadsheet.AutoFilterColumn = worksheet.AutoFilter.Columns(1)
            products.ApplyCustomFilter("*Gi*", DevExpress.Spreadsheet.FilterComparisonOperator.Equal, DevExpress.Spreadsheet.FilterValue.FilterByBlank, DevExpress.Spreadsheet.FilterComparisonOperator.Equal, False)
#End Region  ' #FilterTextByCondition
        End Sub

        Private Sub FilterByValue(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#FilterBySingleValue"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Enable filtering for the "B2:E23" cell range.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)
            ' Filter data in the "Product" column by a specific value.
            worksheet.AutoFilter.Columns(CInt((1))).ApplyFilterCriteria("Mozzarella di Giovanni")
#End Region  ' #FilterBySingleValue
        End Sub

        Private Sub FilterByMultipleValues(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#FilterByMultipleValues"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Enable filtering for the "B2:E23" cell range.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)
            ' Filter data in the "Product" column by an array of values.
            worksheet.AutoFilter.Columns(CInt((1))).ApplyFilterCriteria(New DevExpress.Spreadsheet.CellValue() {"Mozzarella di Giovanni", "Gorgonzola Telino"})
#End Region  ' #FilterByMultipleValues
        End Sub

        Private Sub FilterDatesByCondition(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#FilterDatesByCondition    "
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Enable filtering for the "B2:E23" cell range.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)
            ' Filter values in the "Reported Date" column
            ' to display dates that are between June 1, 2014 and February 1, 2015.
            worksheet.AutoFilter.Columns(CInt((3))).ApplyCustomFilter(New System.DateTime(2014, 6, 1), DevExpress.Spreadsheet.FilterComparisonOperator.GreaterThanOrEqual, New System.DateTime(2015, 2, 1), DevExpress.Spreadsheet.FilterComparisonOperator.LessThanOrEqual, True)
#End Region  ' #FilterDatesByCondition
        End Sub

        Private Sub FilterMixedDataTypesByValues(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#FilterMixedDataByValues    "
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Enable filtering for the "B2:E23" cell range.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)
            ' Create date grouping item to filter January 2015 dates.
            Dim groupings As System.Collections.Generic.IList(Of DevExpress.Spreadsheet.DateGrouping) = New System.Collections.Generic.List(Of DevExpress.Spreadsheet.DateGrouping)()
            Dim dateGroupingJan2015 As DevExpress.Spreadsheet.DateGrouping = New DevExpress.Spreadsheet.DateGrouping(New System.DateTime(2015, 1, 1), DevExpress.Spreadsheet.DateTimeGroupingType.Month)
            groupings.Add(dateGroupingJan2015)
            ' Filter data in the "Reported Date" column
            ' to display values reported in January 2015.
            worksheet.AutoFilter.Columns(CInt((3))).ApplyFilterCriteria("gennaio 2015", groupings)
#End Region  ' #FilterMixedDataByValues
        End Sub

        Private Sub Top10FilterValue(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#TopTenFilter    "
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Enable filtering for the "B2:E23" cell range.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)
            ' Apply a filter to the "Sales" column to display the top ten values.
            worksheet.AutoFilter.Columns(CInt((2))).ApplyTop10Filter(DevExpress.Spreadsheet.Top10Type.Top10Items, 10)
#End Region  ' #TopTenFilter
        End Sub

        Private Sub DynamicFilterValue(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#DynamicFilter"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Enable filtering for the "B2:E23" cell range.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)
            ' Apply a dynamic filter to the "Sales" column
            ' to display only values that are above the average.
            worksheet.AutoFilter.Columns(CInt((2))).ApplyDynamicFilter(DevExpress.Spreadsheet.DynamicFilterType.AboveAverage)
            ' Apply a dynamic filter to the "Reported Date" column
            ' to display values reported this year.
            worksheet.AutoFilter.Columns(CInt((3))).ApplyDynamicFilter(DevExpress.Spreadsheet.DynamicFilterType.ThisYear)
#End Region  ' #DynamicFilter
        End Sub

        Private Sub ReapplyFilterValue(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#ReapplyFilter    "
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Enable filtering for the "B2:E23" cell range.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)
            ' Filter values in the "Sales" column that are greater than 5000$.
            worksheet.AutoFilter.Columns(CInt((2))).ApplyCustomFilter(5000, DevExpress.Spreadsheet.FilterComparisonOperator.GreaterThan)
            ' Change data and reapply the filter.
            worksheet(CStr(("D3"))).Value = 5000
            worksheet.AutoFilter.ReApply()
#End Region  ' #ReapplyFilter
        End Sub

        Private Sub ClearFilter(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#ClearFilter"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Enable filtering for the "B2:E23" cell range.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)
            ' Filter values in the "Sales" column that are greater than 5000$.
            worksheet.AutoFilter.Columns(CInt((2))).ApplyCustomFilter(5000, DevExpress.Spreadsheet.FilterComparisonOperator.GreaterThan)
            ' Clear the filter.
            worksheet.AutoFilter.Clear()
#End Region  ' #ClearFilter
        End Sub

        Private Sub DisableFilter(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#DisableFilter"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Enable filtering for the "B2:E23" cell range.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)
            ' Disable filtering for the entire worksheet.
            worksheet.AutoFilter.Disable()
#End Region  ' #DisableFilter
        End Sub
    End Module
End Namespace
