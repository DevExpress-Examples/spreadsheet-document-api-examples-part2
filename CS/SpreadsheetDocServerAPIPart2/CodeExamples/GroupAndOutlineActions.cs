using DevExpress.Spreadsheet;
using System.Collections.Generic;

namespace SpreadsheetDocServerAPIPart2
{
    public static class GroupAndOutlineActions
    {
        static void GroupRows(Workbook workbook)
        {
            #region #GroupRows    
            Worksheet worksheet = workbook.Worksheets["Grouping"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Group rows 3 through 6 and collapse the group.
            worksheet.Rows.Group(2, 5, true);

            // Group rows 9 through 12 and expand the group.
            worksheet.Rows.Group(8, 11, false);

            // Group rows 2 through 13 to create the outer group. 
            worksheet.Rows.Group(1, 12, false);
            #endregion #GroupRows
        }

        static void GroupColumns(Workbook workbook)
        {
            #region #GroupColumns
            Worksheet worksheet = workbook.Worksheets["Grouping"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Group columns "C" through "F" and expand the group.
            worksheet.Columns.Group(2, 5, false);
            #endregion #GroupColumns
        }

        static void UngroupRows(Workbook workbook)
        {
            #region #UngroupRows    
            Worksheet worksheet = workbook.Worksheets["Grouping and Outline"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Ungroup rows 3 through 6 and display collapsed data.
            worksheet.Rows.UnGroup(2, 5, true);

            // Ungroup rows 9 through 12.
            worksheet.Rows.UnGroup(8, 11, false);

            // Remove the outer row group.
            worksheet.Rows.UnGroup(1, 12, false);
            #endregion #UngroupRows
        }

        static void UngroupColumns(Workbook workbook)
        {
            #region #UngroupColumns
            Worksheet worksheet = workbook.Worksheets["Grouping and Outline"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Ungroup columns "C" through "F".
            worksheet.Columns.UnGroup(2, 5, false);
            #endregion #UngroupColumns
        }

        static void AutoOutline(Workbook workbook)
        {
            #region #AutoOutline
            Worksheet worksheet = workbook.Worksheets["Grouping"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Outline data automatically based on the summary formulas.
            worksheet.AutoOutline();
            #endregion #AutoOutline
        }

        static void Subtotal(Workbook workbook)
        {
            #region #Subtotal
            Worksheet worksheet = workbook.Worksheets["Regional Sales"];
            workbook.Worksheets.ActiveWorksheet = worksheet;
            // Obtain the target cell range.
            CellRange dataRange = worksheet["B3:E23"];
            // Calculate subtotals for column "D".
            List<int> subtotalColumnsList = new List<int>();
            subtotalColumnsList.Add(3);
            // Insert subtotals by each change in column "B"
            // and calculate the SUM fuction for the related rows in column "D".
            worksheet.Subtotal(dataRange, 1, subtotalColumnsList, 9, "Total");
            #endregion #Subtotal
        }
    }
}
