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
                // Insert a picture from a file so that its top left corner is in the "A1" cell.
                // Default picture names are Picture 1.. Picture NN.
                worksheet.Pictures.AddPicture("Pictures\\x-docserver.png", worksheet.Cells["A1"]);
                // Insert a picture at 70 mm from the left, 40 mm from the top, 
                // resize it to a width of 85 mm and a height of 25 mm, and lock the aspect ratio.
                worksheet.Pictures.AddPicture("Pictures\\x-docserver.png", 70, 40, 85, 25, true);
                // Insert a picture.
                worksheet.Pictures.AddPicture("Pictures\\x-docserver.png", 0, 0);
                // Find the last inserted picture by its name.
                Picture picShape = worksheet.Pictures.GetPicturesByName("Picture 3")[0];
                // Remove the last inserted picture.
                picShape.Delete();
            }
            finally
            {
                workbook.EndUpdate();
            }
            #endregion #InsertPicture
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
                // Insert a picture from a file.
                Picture pic = worksheet.Pictures.AddPicture("Pictures\\x-docserver.png", worksheet.Cells["A1"]);
                // Specify the picture name and draw a border.
                pic.Name = "Logo";
                pic.AlternativeText = "Spreadsheet Logo";
                pic.BorderWidth = 1;
                pic.BorderColor = DevExpress.Utils.DXColor.Black;
                // Move a picture.
                pic.Move(20, 30);
                // Specify picture behavior. 
                pic.Placement = Placement.MoveAndSize;
                worksheet.Rows[5].Height += 10;
                worksheet.Columns["D"].Width += 10;
                // Specify the rotation angle.
                pic.Rotation = 30;
                // Add a hyperlink.
                pic.InsertHyperlink("https://www.devexpress.com/products/net/office-file-api/", true);
            }
            finally
            {
                workbook.EndUpdate();
            }
            #endregion #ModifyPicture
        }
    }
}
