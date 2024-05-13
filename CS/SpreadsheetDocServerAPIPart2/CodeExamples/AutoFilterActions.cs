using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace SpreadsheetDocServerAPIPart2
{
    public static class AutoFilterActions
    {
        static void ApplyFilter(Workbook workbook)
        {
            #region #ApplyFilter
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);
            #endregion #ApplyFilter
        }

        static void FilterAndSortBySingleColumn(Workbook workbook)
        {
            #region #FilterAndSortBySingleColumn
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);

            // Sort data in the "B2:E23" range
            // in descending order by column "A".
            worksheet.AutoFilter.SortState.Sort(0, true);
            #endregion #FilterAndSortBySingleColumn
        }

        static void FilterAndSortByMultipleColumns(Workbook workbook)
        {
            #region #FilterAndSortByMultipleColumns
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);

            // Sort data in the "B2:E23" range
            // in descending order by columns "A" and "C".
            List<SortCondition> sortConditions = new List<SortCondition>();
            Color color = worksheet["D12"].Font.Color;

            sortConditions.Add(new SortCondition(0, true));
            sortConditions.Add(new SortCondition(2, color, false));
            worksheet.AutoFilter.SortState.Sort(sortConditions);
            #endregion #FilterAndSortByMultipleColumns
        }

        static void FilterNumericByCondition(Workbook workbook)
        {
            #region #FilterNumbersByCondition
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);

            // Filter values in the "Sales" column that are in a range from 5000$ to 8000$.
            AutoFilterColumn sales = worksheet.AutoFilter.Columns[2];
            sales.ApplyCustomFilter(5000, FilterComparisonOperator.GreaterThanOrEqual, 8000, FilterComparisonOperator.LessThanOrEqual, true);
            #endregion #FilterNumbersByCondition
        }

        static void FilterTextByCondition(Workbook workbook)
        {
            #region #FilterTextByCondition
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);

            // Filter values in the "Product" column that contain "Gi" and include empty cells.
            AutoFilterColumn products = worksheet.AutoFilter.Columns[1];
            products.ApplyCustomFilter("*Gi*", FilterComparisonOperator.Equal, FilterValue.FilterByBlank, FilterComparisonOperator.Equal, false);
            #endregion #FilterTextByCondition
        }

        static void FilterByValue(Workbook workbook)
        {
            #region #FilterBySingleValue
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);

            // Filter data in the "Product" column by a specific value.
            worksheet.AutoFilter.Columns[1].ApplyFilterCriteria("Mozzarella di Giovanni");
            #endregion #FilterBySingleValue
        }

        static void FilterByMultipleValues(Workbook workbook)
        {
            #region #FilterByMultipleValues
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);

            // Filter data in the "Product" column by an array of values.
            worksheet.AutoFilter.Columns[1].ApplyFilterCriteria(new CellValue[] { "Mozzarella di Giovanni", "Gorgonzola Telino" });
            #endregion #FilterByMultipleValues
        }

        static void FilterDatesByCondition(Workbook workbook)
        {
            #region #FilterDatesByCondition    
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);

            // Filter values in the "Reported Date" column
            // to display dates that are between June 1, 2014 and February 1, 2015.
            worksheet.AutoFilter.Columns[3].ApplyCustomFilter
                (new DateTime(2014, 6, 1), FilterComparisonOperator.GreaterThanOrEqual, 
                new DateTime(2015, 2, 1), FilterComparisonOperator.LessThanOrEqual, true);
            #endregion #FilterDatesByCondition
        }

        static void FilterMixedDataTypesByValues(Workbook workbook)
        {
            #region #FilterMixedDataByValues    
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);
            
            // Create date grouping item to filter January 2015 dates.
            IList<DateGrouping> groupings = new List<DateGrouping>();
            DateGrouping dateGroupingJan2015 = new DateGrouping(new DateTime(2015, 1, 1), DateTimeGroupingType.Month);
            groupings.Add(dateGroupingJan2015);

            // Filter data in the "Reported Date" column
            // to display values reported in January 2015.
            worksheet.AutoFilter.Columns[3].ApplyFilterCriteria("gennaio 2015", groupings);
            #endregion #FilterMixedDataByValues
        }

        static void Top10FilterValue(Workbook workbook)
        {
            #region #TopTenFilter    
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);

            // Apply a filter to the "Sales" column to display the top ten values.
            worksheet.AutoFilter.Columns[2].ApplyTop10Filter(Top10Type.Top10Items, 10);
            #endregion #TopTenFilter
        }

        static void DynamicFilterValue(Workbook workbook)
        {
            #region #DynamicFilter
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);

            // Apply a dynamic filter to the "Sales" column
            // to display only values that are above the average.
            worksheet.AutoFilter.Columns[2].ApplyDynamicFilter(DynamicFilterType.AboveAverage);
            // Apply a dynamic filter to the "Reported Date" column
            // to display values reported this year.
            worksheet.AutoFilter.Columns[3].ApplyDynamicFilter(DynamicFilterType.ThisYear);
            #endregion #DynamicFilter
        }


        static void FilterAndSortByColor(Workbook workbook)
        {
            #region #FilterAndSortByColor
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);

            // Sort data in the "B2:E23" range
            // in descending order by column "D".
            Color color = worksheet["D12"].Font.Color;
            worksheet.AutoFilter.SortState.Sort(2, color, false);
            #endregion #FilterAndSortByColor
        }

        static void FilterByBackgroundColor(Workbook workbook)
        {
            #region #FilterByBackgroundColor
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);

            // Filter values in the "Products" column by background color.
            AutoFilterColumn products = worksheet.AutoFilter.Columns[1];
            products.ApplyFillColorFilter(worksheet["C12"].FillColor);
            #endregion #FilterByBackgroundColor
        }

        static void FilterByFillColor(Workbook workbook)
        {
            #region #FilterByFillColor
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);

            // Filter values in the "Products" column by fill color.
            AutoFilterColumn products = worksheet.AutoFilter.Columns[1];
            products.ApplyFillFilter(worksheet["C10"].Fill);
            #endregion #FilterByFillColor
        }

        static void FilterByFontColor(Workbook workbook)
        {
            #region #FilterByFontColor
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);

            // Filter values in the "Sales" column by font color.
            AutoFilterColumn products = worksheet.AutoFilter.Columns[2];
            products.ApplyFontColorFilter(worksheet["D10"].Font.Color);
            #endregion #FilterByFontColor
        }

        static void ReapplyFilterValue(Workbook workbook)
        {
            #region #ReapplyFilter    
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);

            // Filter values in the "Sales" column that are greater than 5000$.
            worksheet.AutoFilter.Columns[2].ApplyCustomFilter(5000, FilterComparisonOperator.GreaterThan);

            // Change data and reapply the filter.
            worksheet["D3"].Value = 5000;
            worksheet.AutoFilter.ReApply();
            #endregion #ReapplyFilter
        }

        static void ClearFilter(Workbook workbook)
        {
            #region #ClearFilter
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);

            // Filter values in the "Sales" column that are greater than 5000$.
            worksheet.AutoFilter.Columns[2].ApplyCustomFilter(5000, FilterComparisonOperator.GreaterThan);

            // Clear the filter.
            worksheet.AutoFilter.Clear();
            #endregion #ClearFilter
        }

        static void DisableFilter(Workbook workbook)
        {
            #region #DisableFilter
            Worksheet worksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Enable filtering for the "B2:E23" cell range.
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);

            // Disable filtering for the entire worksheet.
            worksheet.AutoFilter.Disable();
            #endregion #DisableFilter
        }
    }
}
