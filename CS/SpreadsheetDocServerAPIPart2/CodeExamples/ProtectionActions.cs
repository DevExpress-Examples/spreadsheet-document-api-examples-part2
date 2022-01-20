using DevExpress.Spreadsheet;
using System;
using System.Drawing;

namespace SpreadsheetDocServerAPIPart2
{
    public static class ProtectionActions
    {
        static void ProtectWorkbook(Workbook workbook)
        {
            #region #ProtectWorkbook
            Worksheet worksheet = workbook.Worksheets["ProtectionSample"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Protect workbook structure with a password.
            if (!workbook.IsProtected)
                workbook.Protect("password", true, false);
            // Add a note.
            worksheet["B2"].Value = "Workbook structure is protected with a password. \n You cannot add, move or delete worksheets until protection is removed.";
            worksheet.Visible = true;
            #endregion #ProtectWorkbook
        }
        static void UnprotectWorkbook(Workbook workbook)
        {
            #region #UnprotectWorkbook
            Worksheet worksheet = workbook.Worksheets["ProtectionSample"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Unprotect the workbook.
            if (workbook.IsProtected)
                workbook.Unprotect("password");
            // Add a note.
            worksheet["B2"].Value = "Workbook is unprotected. Workheets can be added, moved or deleted.";
            worksheet.Visible = true;
            #endregion #UnprotectWorkbook
        }
        static void ProtectWorksheet(Workbook workbook)
        {
            #region #ProtectWorksheet
            Worksheet worksheet = workbook.Worksheets["ProtectionSample"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Protect the worksheet with a password.
            if (!worksheet.IsProtected)
                worksheet.Protect("password", WorksheetProtectionPermissions.Default);
            // Add a note.
            worksheet["B2"].Value = "Worksheet is protected with a password. \n You cannot edit or format cells until protection is removed." +
                                    "\nTo remove protection, on the Review tab, in the Changes group," +
                                    "\nclick \"Unprotect Sheet\" and enter \"password\".";
            worksheet.Visible = true;
            #endregion #ProtectWorksheet
        }
        static void UnprotectWorksheet(Workbook workbook)
        {
            #region #UnprotectWorksheet
            Worksheet worksheet = workbook.Worksheets["ProtectionSample"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Unprotect the worksheet.
            if (worksheet.IsProtected)
                worksheet.Unprotect("password");
            // Add a note.
            worksheet["B2"].Value = "Worksheet is unprotected. You can edit and format cells.";
            worksheet.Visible = true;
            #endregion
        }
        static void ProtectRange(Workbook workbook)
        {
            #region #ProtectRange
            Worksheet worksheet = workbook.Worksheets["ProtectionSample"];
            workbook.Worksheets.ActiveWorksheet = worksheet;
            worksheet["B2:J5"].Borders.SetOutsideBorders(Color.Red, BorderLineStyle.Thin);

            // Specify user permission to edit a range in a protected worksheet.
            ProtectedRange protectedRange = worksheet.ProtectedRanges.Add("My Range", worksheet["B2:J5"]);
            EditRangePermission permission = new EditRangePermission();
            permission.UserName = Environment.UserName;
            permission.DomainName = Environment.UserDomainName;
            permission.Deny = false;
            protectedRange.SecurityDescriptor = protectedRange.CreateSecurityDescriptor(new EditRangePermission[] { permission });
            protectedRange.SetPassword("123");
            // Protect the worksheet with a password.
            if (!worksheet.IsProtected)
                worksheet.Protect("password", WorksheetProtectionPermissions.Default);
            // Add a note.
            worksheet["B2"].Value = "This cell range is protected with a password. \n You cannot edit or format it until protection is removed." +
                                    "\nTo remove protection, double-click the range and enter \"123\".";
            worksheet.Visible = true;
            #endregion #ProtectRange
        }
    }
}
