Imports DevExpress.Spreadsheet
Imports System
Imports System.Collections.Generic

Namespace SpreadsheetDocServerAPIPart2
    Public NotInheritable Class AutoFilterActions

        Private Sub New()
        End Sub

        Private Shared Sub ApplyFilter(ByVal workbook As Workbook)
'            #Region "#ApplyFilter"
            Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Enable filtering for the specified cell range.
            Dim range As Range = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)
'            #End Region ' #ApplyFilter
        End Sub

        Private Shared Sub FilterAndSortBySingleColumn(ByVal workbook As Workbook)
'            #Region "#FilterAndSortBySingleColumn"
            Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Enable filtering for the specified cell range.
            Dim range As Range = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)

            ' Sort the data in descending order by the first column.
            worksheet.AutoFilter.SortState.Sort(0, True)
'            #End Region ' #FilterAndSortBySingleColumn
        End Sub

        Private Shared Sub FilterAndSortByMultipleColumns(ByVal workbook As Workbook)
'            #Region "#FilterAndSortByMultipleColumns"
            Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Enable filtering for the specified cell range.
            Dim range As Range = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)

            ' Sort the data in descending order by the first and third columns.
            Dim sortConditions As New List(Of SortCondition)()
            sortConditions.Add(New SortCondition(0, True))
            sortConditions.Add(New SortCondition(2, True))
            worksheet.AutoFilter.SortState.Sort(sortConditions)
'            #End Region ' #FilterAndSortByMultipleColumns
        End Sub

        Private Shared Sub FilterNumericByCondition(ByVal workbook As Workbook)
'            #Region "#FilterNumbersByCondition"
            Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Enable filtering for the specified cell range.
            Dim range As Range = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)

            ' Filter values in the "Sales" column that are in a range from 5000$ to 8000$.
            Dim sales As AutoFilterColumn = worksheet.AutoFilter.Columns(2)
            sales.ApplyCustomFilter(5000, FilterComparisonOperator.GreaterThanOrEqual, 8000, FilterComparisonOperator.LessThanOrEqual, True)
'            #End Region ' #FilterNumbersByCondition
        End Sub

        Private Shared Sub FilterTextByCondition(ByVal workbook As Workbook)
'            #Region "#FilterTextByCondition"
            Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Enable filtering for the specified cell range.
            Dim range As Range = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)

            ' Filter values in the "Product" column that contain "Gi" and include empty cells.
            Dim products As AutoFilterColumn = worksheet.AutoFilter.Columns(1)
            products.ApplyCustomFilter("*Gi*", FilterComparisonOperator.Equal, FilterValue.FilterByBlank, FilterComparisonOperator.Equal, False)
'            #End Region ' #FilterTextByCondition
        End Sub

        Private Shared Sub FilterByValue(ByVal workbook As Workbook)
'            #Region "#FilterBySingleValue"
            Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Enable filtering for the specified cell range.
            Dim range As Range = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)

            ' Filter the data in the "Product" column by a specific value.
            worksheet.AutoFilter.Columns(1).ApplyFilterCriteria("Mozzarella di Giovanni")
'            #End Region ' #FilterBySingleValue
        End Sub

        Private Shared Sub FilterByMultipleValues(ByVal workbook As Workbook)
'            #Region "#FilterByMultipleValues"
            Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Enable filtering for the specified cell range.
            Dim range As Range = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)

            ' Filter the data in the "Product" column by an array of values.
            worksheet.AutoFilter.Columns(1).ApplyFilterCriteria(New CellValue() { "Mozzarella di Giovanni", "Gorgonzola Telino" })
'            #End Region ' #FilterByMultipleValues
        End Sub

        Private Shared Sub FilterDatesByCondition(ByVal workbook As Workbook)
'            #Region "#FilterDatesByCondition    "
            Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Enable filtering for the specified cell range.
            Dim range As Range = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)

            ' Filter values in the "Reported Date" column to display dates that are between June 1, 2014 and February 1, 2015.
            worksheet.AutoFilter.Columns(3).ApplyCustomFilter(New Date(2014, 6, 1), FilterComparisonOperator.GreaterThanOrEqual, New Date(2015, 2, 1), FilterComparisonOperator.LessThanOrEqual, True)
'            #End Region ' #FilterDatesByCondition
        End Sub

        Private Shared Sub FilterMixedDataTypesByValues(ByVal workbook As Workbook)
'            #Region "#FilterMixedDataByValues    "
            Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Enable filtering for the specified cell range.
            Dim range As Range = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)

            ' Create date grouping item to filter January 2015 dates.
            Dim groupings As IList(Of DateGrouping) = New List(Of DateGrouping)()
            Dim dateGroupingJan2015 As New DateGrouping(New Date(2015, 1, 1), DateTimeGroupingType.Month)
            groupings.Add(dateGroupingJan2015)

            ' Filter the data in the "Reported Date" column to display values reported in January 2015.
            worksheet.AutoFilter.Columns(3).ApplyFilterCriteria("gennaio 2015", groupings)
'            #End Region ' #FilterMixedDataByValues
        End Sub

        Private Shared Sub Top10FilterValue(ByVal workbook As Workbook)
'            #Region "#TopTenFilter    "
            Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Enable filtering for the specified cell range.
            Dim range As Range = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)

            ' Apply a filter to the "Sales" column to display the top ten values.
            worksheet.AutoFilter.Columns(2).ApplyTop10Filter(Top10Type.Top10Items, 10)
'            #End Region ' #TopTenFilter
        End Sub

        Private Shared Sub DynamicFilterValue(ByVal workbook As Workbook)
'            #Region "#DynamicFilter"
            Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Enable filtering for the specified cell range.
            Dim range As Range = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)

            ' Apply a dynamic filter to the "Sales" column to display only values that are above the average.
            worksheet.AutoFilter.Columns(2).ApplyDynamicFilter(DynamicFilterType.AboveAverage)
            ' Apply a dynamic filter to the "Reported Date" column to display values reported this year.
            worksheet.AutoFilter.Columns(3).ApplyDynamicFilter(DynamicFilterType.ThisYear)
'            #End Region ' #DynamicFilter
        End Sub

        Private Shared Sub ReapplyFilterValue(ByVal workbook As Workbook)
'            #Region "#ReapplyFilter    "
            Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Enable filtering for the specified cell range.
            Dim range As Range = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)

            ' Filter values in the "Sales" column that are greater than 5000$.
            worksheet.AutoFilter.Columns(2).ApplyCustomFilter(5000, FilterComparisonOperator.GreaterThan)

            ' Change the data and reapply the filter.
            worksheet("D3").Value = 5000
            worksheet.AutoFilter.ReApply()
'            #End Region ' #ReapplyFilter
        End Sub

        Private Shared Sub ClearFilter(ByVal workbook As Workbook)
'            #Region "#ClearFilter"
            Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Enable filtering for the specified cell range.
            Dim range As Range = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)

            ' Filter values in the "Sales" column that are greater than 5000$.
            worksheet.AutoFilter.Columns(2).ApplyCustomFilter(5000, FilterComparisonOperator.GreaterThan)

            ' Clear the filter.
            worksheet.AutoFilter.Clear()
'            #End Region ' #ClearFilter
        End Sub

        Private Shared Sub DisableFilter(ByVal workbook As Workbook)
'            #Region "#DisableFilter"
            Dim worksheet As Worksheet = workbook.Worksheets("Regional sales")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Enable filtering for the specified cell range.
            Dim range As Range = worksheet("B2:E23")
            worksheet.AutoFilter.Apply(range)

            ' Disable filtering for the entire worksheet.
            worksheet.AutoFilter.Disable()
'            #End Region ' #DisableFilter
        End Sub
    End Class
End Namespace
