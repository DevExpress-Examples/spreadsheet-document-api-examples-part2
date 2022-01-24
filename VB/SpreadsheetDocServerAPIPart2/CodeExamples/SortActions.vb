Imports DevExpress.Spreadsheet
Imports System.Collections.Generic

Namespace SpreadsheetDocServerAPIPart2

    Public Module SortActions

        Private Sub SimpleSort(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#SimpleSort"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets(0)
            ' Fill in the range.
            worksheet.Cells(CStr(("A2"))).Value = "Donald Dozier Bradley"
            worksheet.Cells(CStr(("A3"))).Value = "Tony Charles Mccallum-Geteer"
            worksheet.Cells(CStr(("A4"))).Value = "Calvin Liu"
            worksheet.Cells(CStr(("A5"))).Value = "Anita A Boyd"
            worksheet.Cells(CStr(("A6"))).Value = "Angela R. Scott"
            worksheet.Cells(CStr(("A7"))).Value = "D Fox"
            ' Sort the "A2:A7" range in ascending order.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range("A2:A7")
            worksheet.Sort(range)
            ' Create a heading.
            Dim header As DevExpress.Spreadsheet.CellRange = worksheet.Range("A1")
            header(CInt((0))).Value = "Ascending order"
            header.ColumnWidthInCharacters = 30
            header.Style = workbook.Styles("Heading 1")
#End Region  ' #SimpleSort
        End Sub

        Private Sub DescendingOrder(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#DescendingOrder"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets(0)
            ' Fill in the range.
            worksheet.Cells(CStr(("A2"))).Value = "Donald Dozier Bradley"
            worksheet.Cells(CStr(("A3"))).Value = "Tony Charles Mccallum-Geteer"
            worksheet.Cells(CStr(("A4"))).Value = "Calvin Liu"
            worksheet.Cells(CStr(("A5"))).Value = "Anita A Boyd"
            worksheet.Cells(CStr(("A6"))).Value = "Angela R. Scott"
            worksheet.Cells(CStr(("A7"))).Value = "D Fox"
            ' Sort the "A2:A7" range in descending order.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range("A2:A7")
            worksheet.Sort(range, False)
            ' Create a heading.
            Dim header As DevExpress.Spreadsheet.CellRange = worksheet.Range("A1")
            header(CInt((0))).Value = "Descending order"
            header.ColumnWidthInCharacters = 30
            header.Style = workbook.Styles("Heading 1")
#End Region  ' #DescendingOrder
        End Sub

        Private Sub SortBySpecifiedColumn(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#SortBySpecifiedColumn"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("SortSample")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Sort the "A3:F22" range by column "D" in ascending order.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range("A3:F22")
            worksheet.Sort(range, 3)
#End Region  ' #SortBySpecifiedColumn
        End Sub

        Private Sub SortByMultipleColumns(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#SortByMultipleColumns"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("SortSample")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Create sorting fields.
            Dim fields As System.Collections.Generic.List(Of DevExpress.Spreadsheet.SortField) = New System.Collections.Generic.List(Of DevExpress.Spreadsheet.SortField)()
            ' Create the first sorting field.
            Dim sortField1 As DevExpress.Spreadsheet.SortField = New DevExpress.Spreadsheet.SortField()
            ' Sort a cell range by column "A" in ascending order.
            sortField1.ColumnOffset = 0
            sortField1.Comparer = worksheet.Comparers.Ascending
            fields.Add(sortField1)
            ' Create the second sorting field.
            Dim sortField2 As DevExpress.Spreadsheet.SortField = New DevExpress.Spreadsheet.SortField()
            ' Sort a cell range by column "B" in ascending order.
            sortField2.ColumnOffset = 1
            sortField2.Comparer = worksheet.Comparers.Ascending
            fields.Add(sortField2)
            ' Sort the "A3:F22" cell range by sorting fields.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range("A3:F22")
            worksheet.Sort(range, fields)
#End Region  ' #SortByMultipleColumns
        End Sub
    End Module
End Namespace
