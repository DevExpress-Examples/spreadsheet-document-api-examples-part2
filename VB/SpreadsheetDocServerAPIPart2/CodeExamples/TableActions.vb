Imports DevExpress.Spreadsheet
Imports System
Imports System.Drawing

Namespace SpreadsheetDocServerAPIPart2
    Public NotInheritable Class TableActions

        Private Sub New()
        End Sub

        Private Shared Sub CreateTable(ByVal workbook As Workbook)
'            #Region "#CreateTable"
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Insert a table in the worksheet.
            Dim table As Table = worksheet.Tables.Add(worksheet("A1:F12"), False)

            ' Format the table by applying a built-in table style.
            table.Style = workbook.TableStyles(BuiltInTableStyleId.TableStyleMedium20)
'            #End Region ' #CreateTable
        End Sub

        Private Shared Sub TableRanges(ByVal workbook As Workbook)
            '            #Region "#TableRanges"
            Dim worksheet As Worksheet = workbook.Worksheets("TableRanges")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access a table.
            Dim table As Table = worksheet.Tables(0)

            ' Access table columns.
            Dim productColumn As TableColumn = table.Columns(0)
            Dim priceColumn As TableColumn = table.Columns(1)
            Dim quantityColumn As TableColumn = table.Columns(2)
            Dim discountColumn As TableColumn = table.Columns(3)

            ' Add a new column to the end of the table .
            Dim amountColumn As TableColumn = table.Columns.Add()

            ' Set the name of the last column. 
            amountColumn.Name = "Amount"

            ' Set the formula to calculate the amount per product 
            ' and display results in the "Amount" column.
            amountColumn.Formula = "=[Price]*[Quantity]*(1-[Discount])"

            ' Display the total row in the table.
            table.ShowTotals = True

            ' Set the label and function to display the sum of the "Amount" column.
            discountColumn.TotalRowLabel = "Total:"
            amountColumn.TotalRowFunction = TotalRowFunction.Sum

            ' Specify the number format for each column.
            priceColumn.DataRange.NumberFormat = "$#,##0.00"
            discountColumn.DataRange.NumberFormat = "0.0%"
            amountColumn.Range.NumberFormat = "$#,##0.00;$#,##0.00;"""";@"

            ' Specify horizontal alignment for header and total rows of the table.
            table.HeaderRowRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
            table.TotalRowRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center

            ' Specify horizontal alignment to display data in all columns except the first one.
            For i As Integer = 1 To table.Columns.Count - 1
                table.Columns(i).DataRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
            Next i

            ' Set the width of table columns.
            table.Range.ColumnWidthInCharacters = 10
            worksheet.Visible = True
            '            #End Region ' #TableRanges
        End Sub
        Private Shared Sub FormatTable(ByVal workbook As Workbook)
'            #Region "#FormatTable"
            Dim worksheet As Worksheet = workbook.Worksheets("FormatTable")
            workbook.Worksheets.ActiveWorksheet = worksheet


            ' Access a table.
            Dim table As Table = worksheet.Tables(0)

            ' Access the workbook's collection of table styles.
            Dim tableStyles As TableStyleCollection = workbook.TableStyles

            ' Access the built-in table style from the collection by its name.
            Dim tableStyle As TableStyle = tableStyles(BuiltInTableStyleId.TableStyleMedium16)

            ' Apply the table style to the existing table.
            table.Style = tableStyle

            ' Show header and total rows and format them as specified by the applied table style.
            table.ShowHeaders = True
            table.ShowTotals = True

            ' Apply banded column formatting to the table.
            table.ShowTableStyleRowStripes = False
            table.ShowTableStyleColumnStripes = True

            ' Apply special formatting to the first column of the table. 
            table.ShowTableStyleFirstColumn = True
            worksheet.Visible = True
'            #End Region ' #FormatTable
        End Sub


        Private Shared Sub CustomTableStyle(ByVal workbook As Workbook)
'            #Region "#CustomTableStyle"
            Dim worksheet As Worksheet = workbook.Worksheets("Custom Table Style")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access a table.
            Dim table As Table = worksheet.Tables(0)

            Dim styleName As String = "testTableStyle"

            ' If the style under the specified name already exists in the collection,
            If workbook.TableStyles.Contains(styleName) Then
                ' apply this style to the table.
                table.Style = workbook.TableStyles(styleName)
            Else
                ' Add a new table style under the "testTableStyle" name to the TableStyles collection.

                Dim customTableStyle_Renamed As TableStyle = workbook.TableStyles.Add("testTableStyle")

                ' Modify the required formatting characteristics of the table style. 
                ' Specify the format for different table elements.
                customTableStyle_Renamed.BeginUpdate()
                Try
                    customTableStyle_Renamed.TableStyleElements(TableStyleElementType.WholeTable).Font.Color = Color.FromArgb(107, 107, 107)

                    ' Specify formatting characteristics for the table header row. 
                    Dim headerRowStyle As TableStyleElement = customTableStyle_Renamed.TableStyleElements(TableStyleElementType.HeaderRow)
                    headerRowStyle.Fill.BackgroundColor = Color.FromArgb(64, 66, 166)
                    headerRowStyle.Font.Color = Color.White
                    headerRowStyle.Font.Bold = True

                    ' Specify formatting characteristics for the table total row. 
                    Dim totalRowStyle As TableStyleElement = customTableStyle_Renamed.TableStyleElements(TableStyleElementType.TotalRow)
                    totalRowStyle.Fill.BackgroundColor = Color.FromArgb(115, 193, 211)
                    totalRowStyle.Font.Color = Color.White
                    totalRowStyle.Font.Bold = True

                    ' Specify banded row formatting for the table.
                    Dim secondRowStripeStyle As TableStyleElement = customTableStyle_Renamed.TableStyleElements(TableStyleElementType.SecondRowStripe)
                    secondRowStripeStyle.Fill.BackgroundColor = Color.FromArgb(234, 234, 234)
                    secondRowStripeStyle.StripeSize = 1
                Finally
                    customTableStyle_Renamed.EndUpdate()
                End Try
                ' Apply the created custom style to the table.
                table.Style = customTableStyle_Renamed
            End If

            worksheet.Visible = True
'            #End Region ' #CustomTableStyle
        End Sub

        Private Shared Sub DuplicateTableStyle(ByVal workbook As Workbook)
'            #Region "#DuplicateTableStyle"
            Dim worksheet As Worksheet = workbook.Worksheets("Duplicate Table Style")
            workbook.Worksheets.ActiveWorksheet = worksheet


            ' Access a table.
            Dim table1 As Table = worksheet.Tables(0)
            Dim table2 As Table = worksheet.Tables(1)

            ' Get the table style to be duplicated.
            Dim sourceTableStyle As TableStyle = workbook.TableStyles(BuiltInTableStyleId.TableStyleMedium17)

            ' Duplicate the table style.
            Dim newTableStyle As TableStyle = sourceTableStyle.Duplicate()

            ' Modify the required formatting characteristics of the created table style.
            ' For example, change background color of the table header.
            newTableStyle.TableStyleElements(TableStyleElementType.HeaderRow).Fill.BackgroundColor = Color.FromArgb(&HA7, &HEA, &H52)

            table1.Style = sourceTableStyle
            table2.Style = newTableStyle

            worksheet.Visible = True
'            #End Region ' #DuplicateTableStyle
        End Sub
    End Class
End Namespace
