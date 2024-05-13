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

            // Sort the "A2:A7" range in ascending order.
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

            // Sort the "A2:A7" range in descending order.
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

            // Sort the "A3:F22" range by column "D" in ascending order.
            CellRange range = worksheet.Range["A3:F22"];
            worksheet.Sort(range, 3, true);

            #endregion #SortBySpecifiedColumn
        }

        static void SortByMultipleColumns(Workbook workbook)
        {
            #region #SortByMultipleColumns
            Worksheet worksheet = workbook.Worksheets["SortSample"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Create sorting fields.
            List<SortFieldBase> fields = new List<SortFieldBase>();

            // Create the first sorting field.
            SortField sortField1 = new SortField();
            // Sort a cell range by column "A" in ascending order.
            sortField1.ColumnOffset = 0;
            sortField1.Comparer = worksheet.Comparers.Ascending;
            fields.Add(sortField1);

            // Create the second sorting field.
            SortField sortField2 = new SortField();
            // Sort a cell range by column "B" in ascending order.
            sortField2.ColumnOffset = 1;
            sortField2.Comparer = worksheet.Comparers.Ascending;
            fields.Add(sortField2);

            // Sort the "A3:F22" cell range by sorting fields.
            CellRange range = worksheet.Range["A3:F22"];
            worksheet.Sort(range, fields);

            #endregion #SortByMultipleColumns
        }

        static void SortByFillColor(Workbook workbook)
        {
            #region #SortByFillColor
            Worksheet worksheet = workbook.Worksheets["SortSample"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Sort the "A3:F22" range by column "A" in ascending order.
            CellRange range = worksheet.Range["A3:F22"];
            worksheet.Sort(range, 0, worksheet["A3"].Fill);

            #endregion #SortByFillColor
        }

        static void SortByFontColor(Workbook workbook)
        {
            #region #SortByFontColor
            Worksheet worksheet = workbook.Worksheets["SortSample"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Sort the "A3:F22" range by column "F" in ascending order.
            CellRange range = worksheet.Range["A3:F22"];
            worksheet.Sort(range,5, worksheet["F12"].Font.Color);

            #endregion #SortByFontColor
        }


    }
}
