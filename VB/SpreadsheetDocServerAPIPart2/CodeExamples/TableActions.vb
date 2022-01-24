Imports DevExpress.Spreadsheet
Imports System
Imports System.Drawing

Namespace SpreadsheetDocServerAPIPart2

    Public Module TableActions

        Private Sub CreateTable(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#CreateTable"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets(0)
            ' Insert a table in a worksheet.
            Dim table As DevExpress.Spreadsheet.Table = worksheet.Tables.Add(worksheet("A1:F12"), False)
            ' Apply a built-in table style to the table.
            table.Style = workbook.TableStyles(DevExpress.Spreadsheet.BuiltInTableStyleId.TableStyleMedium20)
#End Region  ' #CreateTable
        End Sub

        Private Sub TableRanges(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#TableRanges"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("TableRanges")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Access a table.
            Dim table As DevExpress.Spreadsheet.Table = worksheet.Tables(0)
            ' Obtain table columns.
            Dim productColumn As DevExpress.Spreadsheet.TableColumn = table.Columns(0)
            Dim priceColumn As DevExpress.Spreadsheet.TableColumn = table.Columns(1)
            Dim quantityColumn As DevExpress.Spreadsheet.TableColumn = table.Columns(2)
            Dim discountColumn As DevExpress.Spreadsheet.TableColumn = table.Columns(3)
            ' Add a new column to the end of the table .
            Dim amountColumn As DevExpress.Spreadsheet.TableColumn = table.Columns.Add()
            ' Specify the column name. 
            amountColumn.Name = "Amount"
            ' Specify the formula to calculate the amount for each product 
            ' and display the result in the "Amount" column.
            amountColumn.Formula = "=[Price]*[Quantity]*(1-[Discount])"
            ' Display the total row for the table.
            table.ShowTotals = True
            ' Use the SUM function to calculate the total value for the "Amount" column.
            discountColumn.TotalRowLabel = "Total:"
            amountColumn.TotalRowFunction = DevExpress.Spreadsheet.TotalRowFunction.Sum
            ' Specify the number format for each column.
            priceColumn.DataRange.NumberFormat = "$#,##0.00"
            discountColumn.DataRange.NumberFormat = "0.0%"
            amountColumn.Range.NumberFormat = "$#,##0.00;$#,##0.00;"""";@"
            ' Specify horizontal alignment for the header and total rows.
            table.HeaderRowRange.Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center
            table.TotalRowRange.Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center
            ' Specify horizontal alignment 
            ' for all columns except the first column.
            For i As Integer = 1 To table.Columns.Count - 1
                table.Columns(CInt((i))).DataRange.Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center
            Next

            ' Set the width of table columns.
            table.Range.ColumnWidthInCharacters = 10
            worksheet.Visible = True
#End Region  ' #TableRanges
        End Sub

        Private Sub FormatTable(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#FormatTable"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("FormatTable")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Access a table.
            Dim table As DevExpress.Spreadsheet.Table = worksheet.Tables(0)
            ' Access the workbook's collection of table styles.
            Dim tableStyles As DevExpress.Spreadsheet.TableStyleCollection = workbook.TableStyles
            ' Access the built-in table style by its name.
            Dim tableStyle As DevExpress.Spreadsheet.TableStyle = tableStyles(DevExpress.Spreadsheet.BuiltInTableStyleId.TableStyleMedium16)
            ' Apply the style to the table.
            table.Style = tableStyle
            ' Show header and total rows.
            table.ShowHeaders = True
            table.ShowTotals = True
            ' Enable banded column formatting for the table.
            table.ShowTableStyleRowStripes = False
            table.ShowTableStyleColumnStripes = True
            ' Format the first column in the table. 
            table.ShowTableStyleFirstColumn = True
            worksheet.Visible = True
#End Region  ' #FormatTable
        End Sub

        Private Sub CustomTableStyle(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#CustomTableStyle"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Custom Table Style")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Access a table.
            Dim table As DevExpress.Spreadsheet.Table = worksheet.Tables(0)
            Dim styleName As System.[String] = "testTableStyle"
            ' If a style with the specified name exists in the collection,
            ' apply this style to the table.
            If workbook.TableStyles.Contains(styleName) Then
                table.Style = workbook.TableStyles(styleName)
            Else
                ' Add a new table style under the "testTableStyle" name
                ' to the table style collection.
                Dim lCustomTableStyle As DevExpress.Spreadsheet.TableStyle = workbook.TableStyles.Add("testTableStyle")
                ' Modify table style formatting. 
                ' Specify format characteristics for different table elements.
                lCustomTableStyle.BeginUpdate()
                Try
                    lCustomTableStyle.TableStyleElements(CType((DevExpress.Spreadsheet.TableStyleElementType.WholeTable), DevExpress.Spreadsheet.TableStyleElementType)).Font.Color = System.Drawing.Color.FromArgb(107, 107, 107)
                    ' Format the header row. 
                    Dim headerRowStyle As DevExpress.Spreadsheet.TableStyleElement = lCustomTableStyle.TableStyleElements(DevExpress.Spreadsheet.TableStyleElementType.HeaderRow)
                    headerRowStyle.Fill.BackgroundColor = System.Drawing.Color.FromArgb(64, 66, 166)
                    headerRowStyle.Font.Color = System.Drawing.Color.White
                    headerRowStyle.Font.Bold = True
                    ' Format the total row. 
                    Dim totalRowStyle As DevExpress.Spreadsheet.TableStyleElement = lCustomTableStyle.TableStyleElements(DevExpress.Spreadsheet.TableStyleElementType.TotalRow)
                    totalRowStyle.Fill.BackgroundColor = System.Drawing.Color.FromArgb(115, 193, 211)
                    totalRowStyle.Font.Color = System.Drawing.Color.White
                    totalRowStyle.Font.Bold = True
                    ' Specify banded row formatting for the table.
                    Dim secondRowStripeStyle As DevExpress.Spreadsheet.TableStyleElement = lCustomTableStyle.TableStyleElements(DevExpress.Spreadsheet.TableStyleElementType.SecondRowStripe)
                    secondRowStripeStyle.Fill.BackgroundColor = System.Drawing.Color.FromArgb(234, 234, 234)
                    secondRowStripeStyle.StripeSize = 1
                Finally
                    lCustomTableStyle.EndUpdate()
                End Try

                ' Apply the custom style to the table.
                table.Style = lCustomTableStyle
            End If

            worksheet.Visible = True
#End Region  ' #CustomTableStyle
        End Sub

        Private Sub DuplicateTableStyle(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#DuplicateTableStyle"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Duplicate Table Style")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Access table.
            Dim table1 As DevExpress.Spreadsheet.Table = worksheet.Tables(0)
            Dim table2 As DevExpress.Spreadsheet.Table = worksheet.Tables(1)
            ' Obtain the built-in table style.
            Dim sourceTableStyle As DevExpress.Spreadsheet.TableStyle = workbook.TableStyles(DevExpress.Spreadsheet.BuiltInTableStyleId.TableStyleMedium17)
            ' Duplicate the table style.
            Dim newTableStyle As DevExpress.Spreadsheet.TableStyle = sourceTableStyle.Duplicate()
            ' Modify the duplicated table style's formatting.
            newTableStyle.TableStyleElements(CType((DevExpress.Spreadsheet.TableStyleElementType.HeaderRow), DevExpress.Spreadsheet.TableStyleElementType)).Fill.BackgroundColor = System.Drawing.Color.FromArgb(&HA7, &HEA, &H52)
            ' Apply styles to tables.
            table1.Style = sourceTableStyle
            table2.Style = newTableStyle
            worksheet.Visible = True
#End Region  ' #DuplicateTableStyle
        End Sub
    End Module
End Namespace
