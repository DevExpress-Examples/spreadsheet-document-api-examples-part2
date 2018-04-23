Imports DevExpress.Spreadsheet
Imports System
Imports System.Drawing

Namespace SpreadsheetDocServerAPIPart2
    Public NotInheritable Class ProtectionActions

        Private Sub New()
        End Sub

        Private Shared Sub ProtectWorkbook(ByVal workbook As Workbook)
'            #Region "#ProtectWorkbook"
            Dim worksheet As Worksheet = workbook.Worksheets("ProtectionSample")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Protect workbook structure with the password (prevent users from adding or 
            'deleting worksheets or displaying hidden worksheets).
            If Not workbook.IsProtected Then
                workbook.Protect("password", True, False)
            End If
            ' Add a note.
            worksheet("B2").Value = "Workbook structure is protected with a password. " & ControlChars.Lf & " You cannot add, move or delete worksheets until protection is removed."
            worksheet.Visible = True
'            #End Region ' #ProtectWorkbook
        End Sub
        Private Shared Sub UnprotectWorkbook(ByVal workbook As Workbook)
'            #Region "#UnprotectWorkbook"
            Dim worksheet As Worksheet = workbook.Worksheets("ProtectionSample")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Unprotect the workbook using a password.
            If workbook.IsProtected Then
                workbook.Unprotect("password")
            End If
            ' Add a note.
            worksheet("B2").Value = "Workbook is unprotected. Workheets can be added, moved or deleted."
            worksheet.Visible = True
'            #End Region ' #UnprotectWorkbook
        End Sub
        Private Shared Sub ProtectWorksheet(ByVal workbook As Workbook)
'            #Region "#ProtectWorksheet"
            Dim worksheet As Worksheet = workbook.Worksheets("ProtectionSample")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Protect the worksheet. Prevent end-users from changing worksheet elements.
            If Not worksheet.IsProtected Then
                worksheet.Protect("password", WorksheetProtectionPermissions.Default)
            End If
            ' Add a note.
            worksheet("B2").Value = "Worksheet is protected with a password. " & ControlChars.Lf & " You cannot edit or format cells until protection is removed." & ControlChars.Lf & "To remove protection, on the Review tab, in the Changes group," & ControlChars.Lf & "click ""Unprotect Sheet"" and enter ""password""."
            worksheet.Visible = True
'            #End Region ' #ProtectWorksheet
        End Sub
        Private Shared Sub UnprotectWorksheet(ByVal workbook As Workbook)
'            #Region "#UnprotectWorksheet"
            Dim worksheet As Worksheet = workbook.Worksheets("ProtectionSample")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Remove worksheet protection using a password.
            If worksheet.IsProtected Then
                worksheet.Unprotect("password")
            End If
            ' Add a note.
            worksheet("B2").Value = "Worksheet is unprotected. You can edit and format cells."
            worksheet.Visible = True
'            #End Region
        End Sub
        Private Shared Sub ProtectRange(ByVal workbook As Workbook)
'            #Region "#ProtectRange"
            Dim worksheet As Worksheet = workbook.Worksheets("ProtectionSample")
            workbook.Worksheets.ActiveWorksheet = worksheet
            worksheet("B2:J5").Borders.SetOutsideBorders(Color.Red, BorderLineStyle.Thin)

            ' Specify user permission to edit a range in a protected worksheet.
            Dim protectedRange As ProtectedRange = worksheet.ProtectedRanges.Add("My Range", worksheet("B2:J5"))
            Dim permission As New EditRangePermission()
            permission.UserName = Environment.UserName
            permission.DomainName = Environment.UserDomainName
            permission.Deny = False
            protectedRange.SecurityDescriptor = protectedRange.CreateSecurityDescriptor(New EditRangePermission() { permission })
            protectedRange.SetPassword("123")
            ' Protect a worksheet.
            If Not worksheet.IsProtected Then
                worksheet.Protect("password", WorksheetProtectionPermissions.Default)
            End If
            ' Add a note.
            worksheet("B2").Value = "This cell range is protected with a password. " & ControlChars.Lf & " You cannot edit or format it until protection is removed." & ControlChars.Lf & "To remove protection, double-click the range and enter ""123""."
            worksheet.Visible = True
'            #End Region ' #ProtectRange
        End Sub
    End Class
End Namespace
