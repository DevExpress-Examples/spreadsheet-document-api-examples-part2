Imports DevExpress.Spreadsheet
Imports System.Collections.Generic

Namespace SpreadsheetDocServerAPIPart2
    Public NotInheritable Class SortActions

        Private Sub New()
        End Sub

        Private Shared Sub SimpleSort(ByVal workbook As Workbook)
'            #Region "#SimpleSort"
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Fill in the range.
            worksheet.Cells("A2").Value = "Donald Dozier Bradley"
            worksheet.Cells("A3").Value = "Tony Charles Mccallum-Geteer"
            worksheet.Cells("A4").Value = "Calvin Liu"
            worksheet.Cells("A5").Value = "Anita A Boyd"
            worksheet.Cells("A6").Value = "Angela R. Scott"
            worksheet.Cells("A7").Value = "D Fox"

            ' Sort the range in ascending order.
            Dim range As CellRange = worksheet.Range("A2:A7")
            worksheet.Sort(range)

            ' Create a heading.
            Dim header As CellRange = worksheet.Range("A1")
            header(0).Value = "Ascending order"
            header.ColumnWidthInCharacters = 30
            header.Style = workbook.Styles("Heading 1")
'            #End Region ' #SimpleSort
        End Sub

        Private Shared Sub DescendingOrder(ByVal workbook As Workbook)
'            #Region "#DescendingOrder"
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Fill in the range.
            worksheet.Cells("A2").Value = "Donald Dozier Bradley"
            worksheet.Cells("A3").Value = "Tony Charles Mccallum-Geteer"
            worksheet.Cells("A4").Value = "Calvin Liu"
            worksheet.Cells("A5").Value = "Anita A Boyd"
            worksheet.Cells("A6").Value = "Angela R. Scott"
            worksheet.Cells("A7").Value = "D Fox"

            ' Sort the range in descending order.
            Dim range As CellRange = worksheet.Range("A2:A7")
            worksheet.Sort(range, False)

            ' Create a heading.
            Dim header As CellRange = worksheet.Range("A1")
            header(0).Value = "Descending order"
            header.ColumnWidthInCharacters = 30
            header.Style = workbook.Styles("Heading 1")
'            #End Region ' #DescendingOrder
        End Sub

        Private Shared Sub SortBySpecifiedColumn(ByVal workbook As Workbook)
'            #Region "#SortBySpecifiedColumn"
            Dim worksheet As Worksheet = workbook.Worksheets("SortSample")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Sort by a column with the offset = 3 in the range being sorted.
            ' Use ascending order.
            Dim range As CellRange = worksheet.Range("A3:F22")
            worksheet.Sort(range, 3)

'            #End Region ' #SortBySpecifiedColumn
        End Sub

        Private Shared Sub SortByMultipleColumns(ByVal workbook As Workbook)
'            #Region "#SortByMultipleColumns"
            Dim worksheet As Worksheet = workbook.Worksheets("SortSample")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Create sorting fields.
            Dim fields As New List(Of SortField)()

            ' First sorting field. First column (offset = 0) will be sorted using ascending order.
            Dim sortField1 As New SortField()
            sortField1.ColumnOffset = 0
            sortField1.Comparer = worksheet.Comparers.Ascending
            fields.Add(sortField1)

            ' Second sorting field. Second column (offset = 1) will be sorted using ascending order.
            Dim sortField2 As New SortField()
            sortField2.ColumnOffset = 1
            sortField2.Comparer = worksheet.Comparers.Ascending
            fields.Add(sortField2)

            ' Sort the range by sorting fields.
            Dim range As CellRange = worksheet.Range("A3:F22")
            worksheet.Sort(range, fields)

'            #End Region ' #SortByMultipleColumns
        End Sub
    End Class
End Namespace
