Imports DevExpress.Spreadsheet
Imports System
Imports System.Drawing

Namespace SpreadsheetDocServerAPIPart2
	Public Module TableActions
		Public CreateTableAction As Action(Of Workbook) = AddressOf CreateTable
		Public TableRangesAction As Action(Of Workbook) = AddressOf TableRanges
		Public FormatTableAction As Action(Of Workbook) = AddressOf FormatTable
		Public CustomTableStyleAction As Action(Of Workbook) = AddressOf CustomTableStyle
		Public DuplicateTableStyleAction As Action(Of Workbook) = AddressOf DuplicateTableStyle

		Private Sub CreateTable(ByVal workbook As Workbook)
'			#Region "#CreateTable"
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Insert a table in a worksheet.
			Dim table As Table = worksheet.Tables.Add(worksheet("A1:F12"), False)

			' Apply a built-in table style to the table.
			table.Style = workbook.TableStyles(BuiltInTableStyleId.TableStyleMedium20)
'			#End Region ' #CreateTable
		End Sub

		Private Sub TableRanges(ByVal workbook As Workbook)
'			#Region "#TableRanges"
			Dim worksheet As Worksheet = workbook.Worksheets("TableRanges")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access a table.
			Dim table As Table = worksheet.Tables(0)

			' Obtain table columns.
			Dim productColumn As TableColumn = table.Columns(0)
			Dim priceColumn As TableColumn = table.Columns(1)
			Dim quantityColumn As TableColumn = table.Columns(2)
			Dim discountColumn As TableColumn = table.Columns(3)

			' Add a new column to the end of the table .
			Dim amountColumn As TableColumn = table.Columns.Add()

			' Specify the column name. 
			amountColumn.Name = "Amount"

			' Specify the formula to calculate the amount for each product 
			' and display the result in the "Amount" column.
			amountColumn.Formula = "=[Price]*[Quantity]*(1-[Discount])"

			' Display the total row for the table.
			table.ShowTotals = True

			' Use the SUM function to calculate the total value for the "Amount" column.
			discountColumn.TotalRowLabel = "Total:"
			amountColumn.TotalRowFunction = TotalRowFunction.Sum

			' Specify the number format for each column.
			priceColumn.DataRange.NumberFormat = "$#,##0.00"
			discountColumn.DataRange.NumberFormat = "0.0%"
			amountColumn.Range.NumberFormat = "$#,##0.00;$#,##0.00;"""";@"

			' Specify horizontal alignment for the header and total rows.
			table.HeaderRowRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
			table.TotalRowRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center

			' Specify horizontal alignment 
			' for all columns except the first column.
			For i As Integer = 1 To table.Columns.Count - 1
				table.Columns(i).DataRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
			Next i

			' Set the width of table columns.
			table.Range.ColumnWidthInCharacters = 10
			worksheet.Visible = True
'			#End Region ' #TableRanges
		End Sub
		Private Sub FormatTable(ByVal workbook As Workbook)
'			#Region "#FormatTable"
			Dim worksheet As Worksheet = workbook.Worksheets("FormatTable")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access a table.
			Dim table As Table = worksheet.Tables(0)

			' Access the workbook's collection of table styles.
			Dim tableStyles As TableStyleCollection = workbook.TableStyles

			' Access the built-in table style by its name.
			Dim tableStyle As TableStyle = tableStyles(BuiltInTableStyleId.TableStyleMedium16)

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
'			#End Region ' #FormatTable
		End Sub


		Private Sub CustomTableStyle(ByVal workbook As Workbook)
'			#Region "#CustomTableStyle"
			Dim worksheet As Worksheet = workbook.Worksheets("Custom Table Style")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access a table.
			Dim table As Table = worksheet.Tables(0)

			Dim styleName As String = "testTableStyle"

			' If a style with the specified name exists in the collection,
			' apply this style to the table.
			If workbook.TableStyles.Contains(styleName) Then
				table.Style = workbook.TableStyles(styleName)
			Else
				' Add a new table style under the "testTableStyle" name
				' to the table style collection.
'INSTANT VB NOTE: The variable customTableStyle was renamed since Visual Basic does not handle local variables named the same as class members well:
				Dim customTableStyle_Renamed As TableStyle = workbook.TableStyles.Add("testTableStyle")

				' Modify table style formatting. 
				' Specify format characteristics for different table elements.
				customTableStyle_Renamed.BeginUpdate()
				Try
					customTableStyle_Renamed.TableStyleElements(TableStyleElementType.WholeTable).Font.Color = Color.FromArgb(107, 107, 107)

					' Format the header row. 
					Dim headerRowStyle As TableStyleElement = customTableStyle_Renamed.TableStyleElements(TableStyleElementType.HeaderRow)
					headerRowStyle.Fill.BackgroundColor = Color.FromArgb(64, 66, 166)
					headerRowStyle.Font.Color = Color.White
					headerRowStyle.Font.Bold = True

					' Format the total row. 
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
				' Apply the custom style to the table.
				table.Style = customTableStyle_Renamed
			End If

			worksheet.Visible = True
'			#End Region ' #CustomTableStyle
		End Sub

		Private Sub DuplicateTableStyle(ByVal workbook As Workbook)
'			#Region "#DuplicateTableStyle"
			Dim worksheet As Worksheet = workbook.Worksheets("Duplicate Table Style")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access table.
			Dim table1 As Table = worksheet.Tables(0)
			Dim table2 As Table = worksheet.Tables(1)

			' Obtain the built-in table style.
			Dim sourceTableStyle As TableStyle = workbook.TableStyles(BuiltInTableStyleId.TableStyleMedium17)

			' Duplicate the table style.
			Dim newTableStyle As TableStyle = sourceTableStyle.Duplicate()

			' Modify the duplicated table style's formatting.
			newTableStyle.TableStyleElements(TableStyleElementType.HeaderRow).Fill.BackgroundColor = Color.FromArgb(&HA7, &HEA, &H52)

			' Apply styles to tables.
			table1.Style = sourceTableStyle
			table2.Style = newTableStyle

			worksheet.Visible = True
'			#End Region ' #DuplicateTableStyle
		End Sub
	End Module
End Namespace
