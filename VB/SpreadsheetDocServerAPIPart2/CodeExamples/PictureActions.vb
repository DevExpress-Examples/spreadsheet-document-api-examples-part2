Imports DevExpress.Spreadsheet
Imports System.IO
Imports System.Reflection

Namespace SpreadsheetDocServerAPIPart2
    Public NotInheritable Class PictureActions

        Private Sub New()
        End Sub

        Private Shared Sub InsertPicture(ByVal workbook As Workbook)
'            #Region "#InsertPicture"
            workbook.BeginUpdate()
            ' Set the measurement unit to Millimeter.
            workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                ' Insert a picture from a file so that its top left corner is in the specified cell.
                ' By default the picture is named Picture 1.. Picture NN.
                worksheet.Pictures.AddPicture("Pictures\x-docserver.png", worksheet.Cells("A1"))
                ' Insert a picture at 70 mm from the left, 40 mm from the top, 
                ' and resize it to a width of 85 mm and a height of 25 mm, locking the aspect ratio.
                worksheet.Pictures.AddPicture("Pictures\x-docserver.png", 70, 40, 85, 25, True)
                ' Insert the picture to be removed.
                worksheet.Pictures.AddPicture("Pictures\x-docserver.png", 0, 0)
                ' Remove the last inserted picture.
                ' Find the shape by its name. The method returns a collection of shapes with the same name.
                Dim picShape As Picture = worksheet.Pictures.GetPicturesByName("Picture 3")(0)
                picShape.Delete()
            Finally
                workbook.EndUpdate()
            End Try
'            #End Region ' #InsertPicture
        End Sub

        Private Shared Sub InsertPictureFromUri(ByVal workbook As Workbook)
'            #Region "#InsertPictureFromUri"
            Dim imageUri As String = "https://www.devexpress.com/Products/NET/Controls/WinForms/spreadsheet/i/winforms-spreadsheet-control.png"
            ' Create an image from Uri.
            Dim imageSource As SpreadsheetImageSource = SpreadsheetImageSource.FromUri(imageUri, workbook)
            ' Set the measurement unit to point.
            workbook.Unit = DevExpress.Office.DocumentUnit.Point

            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                ' Insert a picture from the SpreadsheetImageSource at 100 pt from the left, 40 pt from the top, 
                ' and resize it to a width of 300 pt and a height of 200 pt.
                worksheet.Pictures.AddPicture(imageSource, 100, 40, 300, 200)
            Finally
                workbook.EndUpdate()
            End Try
'            #End Region ' #InsertPictureFromUri

        End Sub

        Private Shared Sub ModifyPicture(ByVal workbook As Workbook)

'            #Region "#ModifyPicture"
            ' Set the measurement unit to Millimeter.
            workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                ' Insert pictures from the file.
                Dim pic As Picture = worksheet.Pictures.AddPicture("Pictures\x-docserver.png", worksheet.Cells("A1"))
                ' Specify picture name and draw a border.
                pic.Name = "Logo"
                pic.AlternativeText = "Spreadsheet Logo"
                pic.BorderWidth = 1
                pic.BorderColor = DevExpress.Utils.DXColor.Black
                ' Move a picture.
                pic.Move(20, 30)
                ' Change picture behavior so it will move and size with underlying cells. 
                pic.Placement = Placement.MoveAndSize
                worksheet.Rows(5).Height += 10
                worksheet.Columns("D").Width += 10
                ' Specify rotation angle.
                pic.Rotation = 30
                ' Add a hyperlink.
                pic.InsertHyperlink("http://www.devexpress.com/Products/NET/Document-Server/", True)
            Finally
                workbook.EndUpdate()
            End Try
'            #End Region ' #ModifyPicture
        End Sub
    End Class
End Namespace
