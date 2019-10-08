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

            // Group four rows starting from the third row and collapse the group.
            worksheet.Rows.Group(2, 5, true);

            // Group four rows starting from the ninth row and expand the group.
            worksheet.Rows.Group(8, 11, false);

            // Create the outer group of rows by grouping rows 2 through 13. 
            worksheet.Rows.Group(1, 12, false);
            #endregion #GroupRows
        }

        static void GroupColumns(Workbook workbook)
        {
            #region #GroupColumns
            Worksheet worksheet = workbook.Worksheets["Grouping"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Group four columns starting from the third column "C" and expand the group.
            worksheet.Columns.Group(2, 5, false);
            #endregion #GroupColumns
        }

        static void UngroupRows(Workbook workbook)
        {
            #region #UngroupRows    
            Worksheet worksheet = workbook.Worksheets["Grouping and Outline"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Ungroup four rows (from the third row to the sixth row) and display collapsed data.
            worksheet.Rows.UnGroup(2, 5, true);

            // Ungroup four rows (from the ninth row to the twelfth row).
            worksheet.Rows.UnGroup(8, 11, false);

            // Remove the outer group of rows.
            worksheet.Rows.UnGroup(1, 12, false);
            #endregion #UngroupRows
        }

        static void UngroupColumns(Workbook workbook)
        {
            #region #UngroupColumns
            Worksheet worksheet = workbook.Worksheets["Grouping and Outline"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Ungroup four columns (from the column "C" to the column "F").
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

            CellRange dataRange = worksheet["B3:E23"];
            // Specify that subtotals should be calculated for the column "D". 
            List<int> subtotalColumnsList = new List<int>();
            subtotalColumnsList.Add(3);
            // Insert subtotals by each change in the column "B" and calculate the SUM fuction for the related rows in the column "D".
            worksheet.Subtotal(dataRange, 1, subtotalColumnsList, 9, "Total");
            #endregion #Subtotal
        }
    }
}
