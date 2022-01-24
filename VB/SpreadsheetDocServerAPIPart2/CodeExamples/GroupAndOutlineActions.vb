Imports DevExpress.Spreadsheet
Imports System.Collections.Generic

Namespace SpreadsheetDocServerAPIPart2

    Public Module GroupAndOutlineActions

        Private Sub GroupRows(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#GroupRows    "
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Grouping")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Group rows 3 through 6 and collapse the group.
            worksheet.Rows.Group(2, 5, True)
            ' Group rows 9 through 12 and expand the group.
            worksheet.Rows.Group(8, 11, False)
            ' Group rows 2 through 13 to create the outer group. 
            worksheet.Rows.Group(1, 12, False)
#End Region  ' #GroupRows
        End Sub

        Private Sub GroupColumns(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#GroupColumns"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Grouping")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Group columns "C" through "F" and expand the group.
            worksheet.Columns.Group(2, 5, False)
#End Region  ' #GroupColumns
        End Sub

        Private Sub UngroupRows(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#UngroupRows    "
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Grouping and Outline")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Ungroup rows 3 through 6 and display collapsed data.
            worksheet.Rows.UnGroup(2, 5, True)
            ' Ungroup rows 9 through 12.
            worksheet.Rows.UnGroup(8, 11, False)
            ' Remove the outer row group.
            worksheet.Rows.UnGroup(1, 12, False)
#End Region  ' #UngroupRows
        End Sub

        Private Sub UngroupColumns(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#UngroupColumns"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Grouping and Outline")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Ungroup columns "C" through "F".
            worksheet.Columns.UnGroup(2, 5, False)
#End Region  ' #UngroupColumns
        End Sub

        Private Sub AutoOutline(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#AutoOutline"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Grouping")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Outline data automatically based on the summary formulas.
            worksheet.AutoOutline()
#End Region  ' #AutoOutline
        End Sub

        Private Sub Subtotal(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#Subtotal"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("Regional Sales")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Obtain the target cell range.
            Dim dataRange As DevExpress.Spreadsheet.CellRange = worksheet("B3:E23")
            ' Calculate subtotals for column "D".
            Dim subtotalColumnsList As System.Collections.Generic.List(Of Integer) = New System.Collections.Generic.List(Of Integer)()
            subtotalColumnsList.Add(3)
            ' Insert subtotals by each change in column "B"
            ' and calculate the SUM fuction for the related rows in column "D".
            worksheet.Subtotal(dataRange, 1, subtotalColumnsList, 9, "Total")
#End Region  ' #Subtotal
        End Sub
    End Module
End Namespace
