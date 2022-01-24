Imports DevExpress.Spreadsheet
Imports System
Imports System.Drawing

Namespace SpreadsheetDocServerAPIPart2

    Public Module ProtectionActions

        Private Sub ProtectWorkbook(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#ProtectWorkbook"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("ProtectionSample")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Protect workbook structure with a password.
            If Not workbook.IsProtected Then workbook.Protect("password", True, False)
            ' Add a note.
            worksheet(CStr(("B2"))).Value = "Workbook structure is protected with a password. " & Global.Microsoft.VisualBasic.Constants.vbLf & " You cannot add, move or delete worksheets until protection is removed."
            worksheet.Visible = True
#End Region  ' #ProtectWorkbook
        End Sub

        Private Sub UnprotectWorkbook(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#UnprotectWorkbook"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("ProtectionSample")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Unprotect the workbook.
            If workbook.IsProtected Then workbook.Unprotect("password")
            ' Add a note.
            worksheet(CStr(("B2"))).Value = "Workbook is unprotected. Workheets can be added, moved or deleted."
            worksheet.Visible = True
#End Region  ' #UnprotectWorkbook
        End Sub

        Private Sub ProtectWorksheet(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#ProtectWorksheet"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("ProtectionSample")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Protect the worksheet with a password.
            If Not worksheet.IsProtected Then worksheet.Protect("password", DevExpress.Spreadsheet.WorksheetProtectionPermissions.[Default])
            ' Add a note.
            worksheet(CStr(("B2"))).Value = "Worksheet is protected with a password. " & Global.Microsoft.VisualBasic.Constants.vbLf & " You cannot edit or format cells until protection is removed." & Global.Microsoft.VisualBasic.Constants.vbLf & "To remove protection, on the Review tab, in the Changes group," & Global.Microsoft.VisualBasic.Constants.vbLf & "click ""Unprotect Sheet"" and enter ""password""."
            worksheet.Visible = True
#End Region  ' #ProtectWorksheet
        End Sub

        Private Sub UnprotectWorksheet(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#UnprotectWorksheet"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("ProtectionSample")
            workbook.Worksheets.ActiveWorksheet = worksheet
            ' Unprotect the worksheet.
            If worksheet.IsProtected Then worksheet.Unprotect("password")
            ' Add a note.
            worksheet(CStr(("B2"))).Value = "Worksheet is unprotected. You can edit and format cells."
            worksheet.Visible = True
#End Region
        End Sub

        Private Sub ProtectRange(ByVal workbook As DevExpress.Spreadsheet.Workbook)
#Region "#ProtectRange"
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets("ProtectionSample")
            workbook.Worksheets.ActiveWorksheet = worksheet
            worksheet(CStr(("B2:J5"))).Borders.SetOutsideBorders(System.Drawing.Color.Red, DevExpress.Spreadsheet.BorderLineStyle.Thin)
            ' Specify user permission to edit a range in a protected worksheet.
            Dim protectedRange As DevExpress.Spreadsheet.ProtectedRange = worksheet.ProtectedRanges.Add("My Range", worksheet("B2:J5"))
            Dim permission As DevExpress.Spreadsheet.EditRangePermission = New DevExpress.Spreadsheet.EditRangePermission()
            permission.UserName = System.Environment.UserName
            permission.DomainName = System.Environment.UserDomainName
            permission.Deny = False
            protectedRange.SecurityDescriptor = protectedRange.CreateSecurityDescriptor(New DevExpress.Spreadsheet.EditRangePermission() {permission})
            protectedRange.SetPassword("123")
            ' Protect the worksheet with a password.
            If Not worksheet.IsProtected Then worksheet.Protect("password", DevExpress.Spreadsheet.WorksheetProtectionPermissions.[Default])
            ' Add a note.
            worksheet(CStr(("B2"))).Value = "This cell range is protected with a password. " & Global.Microsoft.VisualBasic.Constants.vbLf & " You cannot edit or format it until protection is removed." & Global.Microsoft.VisualBasic.Constants.vbLf & "To remove protection, double-click the range and enter ""123""."
            worksheet.Visible = True
#End Region  ' #ProtectRange
        End Sub
    End Module
End Namespace
