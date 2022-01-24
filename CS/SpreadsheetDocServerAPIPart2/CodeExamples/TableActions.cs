using DevExpress.Spreadsheet;
using System;
using System.Drawing;

namespace SpreadsheetDocServerAPIPart2
{
    public static class TableActions
    {
        static void CreateTable(Workbook workbook)
        {
            #region #CreateTable
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert a table in a worksheet.
            Table table = worksheet.Tables.Add(worksheet["A1:F12"], false);

            // Apply a built-in table style to the table.
            table.Style = workbook.TableStyles[BuiltInTableStyleId.TableStyleMedium20];
            #endregion #CreateTable
        }

        static void TableRanges(Workbook workbook)
        {
            #region #TableRanges
            Worksheet worksheet = workbook.Worksheets["TableRanges"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access a table.
            Table table = worksheet.Tables[0];

            // Obtain table columns.
            TableColumn productColumn = table.Columns[0];
            TableColumn priceColumn = table.Columns[1];
            TableColumn quantityColumn = table.Columns[2];
            TableColumn discountColumn = table.Columns[3];

            // Add a new column to the end of the table .
            TableColumn amountColumn = table.Columns.Add();

            // Specify the column name. 
            amountColumn.Name = "Amount";

            // Specify the formula to calculate the amount for each product 
            // and display the result in the "Amount" column.
            amountColumn.Formula = "=[Price]*[Quantity]*(1-[Discount])";

            // Display the total row for the table.
            table.ShowTotals = true;

            // Use the SUM function to calculate the total value for the "Amount" column.
            discountColumn.TotalRowLabel = "Total:";
            amountColumn.TotalRowFunction = TotalRowFunction.Sum;

            // Specify the number format for each column.
            priceColumn.DataRange.NumberFormat = "$#,##0.00";
            discountColumn.DataRange.NumberFormat = "0.0%";
            amountColumn.Range.NumberFormat = "$#,##0.00;$#,##0.00;\"\";@";

            // Specify horizontal alignment for the header and total rows.
            table.HeaderRowRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            table.TotalRowRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

            // Specify horizontal alignment 
            // for all columns except the first column.
            for (int i = 1; i < table.Columns.Count; i++)
            {
                table.Columns[i].DataRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            }

            // Set the width of table columns.
            table.Range.ColumnWidthInCharacters = 10;
            worksheet.Visible = true;
            #endregion #TableRanges
        }
        static void FormatTable(Workbook workbook)
        {
            #region #FormatTable
            Worksheet worksheet = workbook.Worksheets["FormatTable"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access a table.
            Table table = worksheet.Tables[0];

            // Access the workbook's collection of table styles.
            TableStyleCollection tableStyles = workbook.TableStyles;

            // Access the built-in table style by its name.
            TableStyle tableStyle = tableStyles[BuiltInTableStyleId.TableStyleMedium16];

            // Apply the style to the table.
            table.Style = tableStyle;

            // Show header and total rows.
            table.ShowHeaders = true;
            table.ShowTotals = true;

            // Enable banded column formatting for the table.
            table.ShowTableStyleRowStripes = false;
            table.ShowTableStyleColumnStripes = true;

            // Format the first column in the table. 
            table.ShowTableStyleFirstColumn = true;
            worksheet.Visible = true;
            #endregion #FormatTable
        }


        static void CustomTableStyle(Workbook workbook)
        {
            #region #CustomTableStyle
            Worksheet worksheet = workbook.Worksheets["Custom Table Style"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access a table.
            Table table = worksheet.Tables[0];

            String styleName = "testTableStyle";

            // If a style with the specified name exists in the collection,
            // apply this style to the table.
            if (workbook.TableStyles.Contains(styleName))
            {
                table.Style = workbook.TableStyles[styleName];
            }
            else
            {
                // Add a new table style under the "testTableStyle" name
                // to the table style collection.
                TableStyle customTableStyle = workbook.TableStyles.Add("testTableStyle");

                // Modify table style formatting. 
                // Specify format characteristics for different table elements.
                customTableStyle.BeginUpdate();
                try
                {
                    customTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Font.Color = Color.FromArgb(107, 107, 107);

                    // Format the header row. 
                    TableStyleElement headerRowStyle = customTableStyle.TableStyleElements[TableStyleElementType.HeaderRow];
                    headerRowStyle.Fill.BackgroundColor = Color.FromArgb(64, 66, 166);
                    headerRowStyle.Font.Color = Color.White;
                    headerRowStyle.Font.Bold = true;

                    // Format the total row. 
                    TableStyleElement totalRowStyle = customTableStyle.TableStyleElements[TableStyleElementType.TotalRow];
                    totalRowStyle.Fill.BackgroundColor = Color.FromArgb(115, 193, 211);
                    totalRowStyle.Font.Color = Color.White;
                    totalRowStyle.Font.Bold = true;

                    // Specify banded row formatting for the table.
                    TableStyleElement secondRowStripeStyle = customTableStyle.TableStyleElements[TableStyleElementType.SecondRowStripe];
                    secondRowStripeStyle.Fill.BackgroundColor = Color.FromArgb(234, 234, 234);
                    secondRowStripeStyle.StripeSize = 1;
                }
                finally
                {
                    customTableStyle.EndUpdate();
                }
                // Apply the custom style to the table.
                table.Style = customTableStyle;
            }

            worksheet.Visible = true;
            #endregion #CustomTableStyle
        }

        static void DuplicateTableStyle(Workbook workbook)
        {
            #region #DuplicateTableStyle
            Worksheet worksheet = workbook.Worksheets["Duplicate Table Style"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access table.
            Table table1 = worksheet.Tables[0];
            Table table2 = worksheet.Tables[1];

            // Obtain the built-in table style.
            TableStyle sourceTableStyle = workbook.TableStyles[BuiltInTableStyleId.TableStyleMedium17];

            // Duplicate the table style.
            TableStyle newTableStyle = sourceTableStyle.Duplicate();

            // Modify the duplicated table style's formatting.
            newTableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Fill.BackgroundColor = Color.FromArgb(0xA7, 0xEA, 0x52);

            // Apply styles to tables.
            table1.Style = sourceTableStyle;
            table2.Style = newTableStyle;

            worksheet.Visible = true;
            #endregion #DuplicateTableStyle
        }
    }
}
