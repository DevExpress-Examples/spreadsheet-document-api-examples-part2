Imports DevExpress.Spreadsheet
Imports System.Collections.Generic

Namespace SpreadsheetDocServerAPIPart2
    Public NotInheritable Class GroupAndOutlineActions

        Private Sub New()
        End Sub

        Private Shared Sub GroupRows(ByVal workbook As Workbook)
'            #Region "#GroupRows    "
            Dim worksheet As Worksheet = workbook.Worksheets("Grouping")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Group four rows starting from the third row and collapse the group.
            worksheet.Rows.Group(2, 5, True)

            ' Group four rows starting from the ninth row and expand the group.
            worksheet.Rows.Group(8, 11, False)

            ' Create the outer group of rows by grouping rows 2 through 13. 
            worksheet.Rows.Group(1, 12, False)
'            #End Region ' #GroupRows
        End Sub

        Private Shared Sub GroupColumns(ByVal workbook As Workbook)
'            #Region "#GroupColumns"
            Dim worksheet As Worksheet = workbook.Worksheets("Grouping")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Group four columns starting from the third column "C" and expand the group.
            worksheet.Columns.Group(2, 5, False)
'            #End Region ' #GroupColumns
        End Sub

        Private Shared Sub UngroupRows(ByVal workbook As Workbook)
'            #Region "#UngroupRows    "
            Dim worksheet As Worksheet = workbook.Worksheets("Grouping and Outline")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Ungroup four rows (from the third row to the sixth row) and display collapsed data.
            worksheet.Rows.UnGroup(2, 5, True)

            ' Ungroup four rows (from the ninth row to the twelfth row).
            worksheet.Rows.UnGroup(8, 11, False)

            ' Remove the outer group of rows.
            worksheet.Rows.UnGroup(1, 12, False)
'            #End Region ' #UngroupRows
        End Sub

        Private Shared Sub UngroupColumns(ByVal workbook As Workbook)
'            #Region "#UngroupColumns"
            Dim worksheet As Worksheet = workbook.Worksheets("Grouping and Outline")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Ungroup four columns (from the column "C" to the column "F").
            worksheet.Columns.UnGroup(2, 5, False)
'            #End Region ' #UngroupColumns
        End Sub

        Private Shared Sub AutoOutline(ByVal workbook As Workbook)
'            #Region "#AutoOutline"
            Dim worksheet As Worksheet = workbook.Worksheets("Grouping")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Outline data automatically based on the summary formulas.
            worksheet.AutoOutline()
'            #End Region ' #AutoOutline
        End Sub

        Private Shared Sub Subtotal(ByVal workbook As Workbook)
'            #Region "#Subtotal"
            Dim worksheet As Worksheet = workbook.Worksheets("Regional Sales")
            workbook.Worksheets.ActiveWorksheet = worksheet

            Dim dataRange As Range = worksheet("B3:E23")
            ' Specify that subtotals should be calculated for the column "D". 
            Dim subtotalColumnsList As New List(Of Integer)()
            subtotalColumnsList.Add(3)
            ' Insert subtotals by each change in the column "B" and calculate the SUM fuction for the related rows in the column "D".
            worksheet.Subtotal(dataRange, 1, subtotalColumnsList, 9, "Total")
'            #End Region ' #Subtotal
        End Sub
    End Class
End Namespace
