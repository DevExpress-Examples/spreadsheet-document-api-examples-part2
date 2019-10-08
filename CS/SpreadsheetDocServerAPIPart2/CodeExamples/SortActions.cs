using DevExpress.Spreadsheet;
using System.Collections.Generic;

namespace SpreadsheetDocServerAPIPart2
{
    public static class SortActions
    {
        static void SimpleSort(Workbook workbook)
        {
            #region #SimpleSort
            Worksheet worksheet = workbook.Worksheets[0];

            // Fill in the range.
            worksheet.Cells["A2"].Value = "Donald Dozier Bradley";
            worksheet.Cells["A3"].Value = "Tony Charles Mccallum-Geteer";
            worksheet.Cells["A4"].Value = "Calvin Liu";
            worksheet.Cells["A5"].Value = "Anita A Boyd";
            worksheet.Cells["A6"].Value = "Angela R. Scott";
            worksheet.Cells["A7"].Value = "D Fox";

            // Sort the range in ascending order.
            CellRange range = worksheet.Range["A2:A7"];
            worksheet.Sort(range);

            // Create a heading.
            CellRange header = worksheet.Range["A1"];
            header[0].Value = "Ascending order";
            header.ColumnWidthInCharacters = 30;
            header.Style = workbook.Styles["Heading 1"];
            #endregion #SimpleSort
        }

        static void DescendingOrder(Workbook workbook)
        {
            #region #DescendingOrder
            Worksheet worksheet = workbook.Worksheets[0];

            // Fill in the range.
            worksheet.Cells["A2"].Value = "Donald Dozier Bradley";
            worksheet.Cells["A3"].Value = "Tony Charles Mccallum-Geteer";
            worksheet.Cells["A4"].Value = "Calvin Liu";
            worksheet.Cells["A5"].Value = "Anita A Boyd";
            worksheet.Cells["A6"].Value = "Angela R. Scott";
            worksheet.Cells["A7"].Value = "D Fox";

            // Sort the range in descending order.
            CellRange range = worksheet.Range["A2:A7"];
            worksheet.Sort(range, false);

            // Create a heading.
            CellRange header = worksheet.Range["A1"];
            header[0].Value = "Descending order";
            header.ColumnWidthInCharacters = 30;
            header.Style = workbook.Styles["Heading 1"];
            #endregion #DescendingOrder
        }

        static void SortBySpecifiedColumn(Workbook workbook)
        {
            #region #SortBySpecifiedColumn
            Worksheet worksheet = workbook.Worksheets["SortSample"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Sort by a column with the offset = 3 in the range being sorted.
            // Use ascending order.
            CellRange range = worksheet.Range["A3:F22"];
            worksheet.Sort(range, 3);

            #endregion #SortBySpecifiedColumn
        }

        static void SortByMultipleColumns(Workbook workbook)
        {
            #region #SortByMultipleColumns
            Worksheet worksheet = workbook.Worksheets["SortSample"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Create sorting fields.
            List<SortField> fields = new List<SortField>();

            // First sorting field. First column (offset = 0) will be sorted using ascending order.
            SortField sortField1 = new SortField();
            sortField1.ColumnOffset = 0;
            sortField1.Comparer = worksheet.Comparers.Ascending;
            fields.Add(sortField1);

            // Second sorting field. Second column (offset = 1) will be sorted using ascending order.
            SortField sortField2 = new SortField();
            sortField2.ColumnOffset = 1;
            sortField2.Comparer = worksheet.Comparers.Ascending;
            fields.Add(sortField2);

            // Sort the range by sorting fields.
            CellRange range = worksheet.Range["A3:F22"];
            worksheet.Sort(range, fields);

            #endregion #SortByMultipleColumns
        }
    }
}
