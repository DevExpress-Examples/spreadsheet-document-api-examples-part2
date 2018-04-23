using DevExpress.Spreadsheet;
using System.IO;
using System.Reflection;

namespace SpreadsheetDocServerAPIPart2
{
    public static class PictureActions
    {
        static void InsertPicture(Workbook workbook)
        {
            #region #InsertPicture
            workbook.BeginUpdate();
            // Set the measurement unit to Millimeter.
            workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter;
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];
                // Insert a picture from a file so that its top left corner is in the specified cell.
                // By default the picture is named Picture 1.. Picture NN.
                worksheet.Pictures.AddPicture("Pictures\\x-docserver.png", worksheet.Cells["A1"]);
                // Insert a picture at 70 mm from the left, 40 mm from the top, 
                // and resize it to a width of 85 mm and a height of 25 mm, locking the aspect ratio.
                worksheet.Pictures.AddPicture("Pictures\\x-docserver.png", 70, 40, 85, 25, true);
                // Insert the picture to be removed.
                worksheet.Pictures.AddPicture("Pictures\\x-docserver.png", 0, 0);
                // Remove the last inserted picture.
                // Find the shape by its name. The method returns a collection of shapes with the same name.
                Picture picShape = worksheet.Pictures.GetPicturesByName("Picture 3")[0];
                picShape.Delete();
            }
            finally
            {
                workbook.EndUpdate();
            }
            #endregion #InsertPicture
        }

        static void InsertPictureFromUri(Workbook workbook)
        {
            #region #InsertPictureFromUri
            string imageUri = "https://www.devexpress.com/Products/NET/Controls/WinForms/spreadsheet/i/winforms-spreadsheet-control.png";
            // Create an image from Uri.
            SpreadsheetImageSource imageSource = SpreadsheetImageSource.FromUri(imageUri, workbook);
            // Set the measurement unit to point.
            workbook.Unit = DevExpress.Office.DocumentUnit.Point;

            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];
                // Insert a picture from the SpreadsheetImageSource at 100 pt from the left, 40 pt from the top, 
                // and resize it to a width of 300 pt and a height of 200 pt.
                worksheet.Pictures.AddPicture(imageSource, 100, 40, 300, 200);
            }
            finally
            {
                workbook.EndUpdate();
            }
            #endregion #InsertPictureFromUri

        }

        static void ModifyPicture(Workbook workbook)
        {

            #region #ModifyPicture
            // Set the measurement unit to Millimeter.
            workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter;
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];
                // Insert pictures from the file.
                Picture pic = worksheet.Pictures.AddPicture("Pictures\\x-docserver.png", worksheet.Cells["A1"]);
                // Specify picture name and draw a border.
                pic.Name = "Logo";
                pic.AlternativeText = "Spreadsheet Logo";
                pic.BorderWidth = 1;
                pic.BorderColor = DevExpress.Utils.DXColor.Black;
                // Move a picture.
                pic.Move(20, 30);
                // Change picture behavior so it will move and size with underlying cells. 
                pic.Placement = Placement.MoveAndSize;
                worksheet.Rows[5].Height += 10;
                worksheet.Columns["D"].Width += 10;
                // Specify rotation angle.
                pic.Rotation = 30;
                // Add a hyperlink.
                pic.InsertHyperlink("http://www.devexpress.com/Products/NET/Document-Server/", true);
            }
            finally
            {
                workbook.EndUpdate();
            }
            #endregion #ModifyPicture
        }
    }
}
