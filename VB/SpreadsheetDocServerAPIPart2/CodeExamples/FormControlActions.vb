Imports DevExpress.Spreadsheet
Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace SpreadsheetDocServerAPIPart2
	Public Module FormControlActions
		Public CreateFormControlsAction As Action(Of Workbook) = AddressOf CreateFormControls
		Public EditFormControlsAction As Action(Of Workbook) = AddressOf EditFormControls
		Private Sub CreateFormControls(ByVal workbook As IWorkbook)
'			#Region "#CreateFormControls"
			Dim formControls = workbook.Worksheets(0).FormControls

			' Create a button form control:
			Dim buttonCellRange = workbook.Worksheets(0).Range("B2:C2")
			Dim buttonFormControl = formControls.AddButton(buttonCellRange)
			buttonFormControl.PlainText = "Click Here"

			' Create a list box form control:
			Dim comboCellRange = workbook.Worksheets(0).Range("B4:C4")
			Dim comboBoxControl = formControls.AddComboBox(comboCellRange)
			comboBoxControl.DropDownLines = 3
			comboBoxControl.SourceRange = workbook.Worksheets(0).Range("E2:E6")
			comboBoxControl.SelectedIndex = 1

			' Create a check box form control:
			Dim checkRange = workbook.Worksheets(0).Range("D5:E5")
			Dim checkBoxControl = formControls.AddCheckBox(checkRange)
			checkBoxControl.CheckState = FormControlCheckState.Checked
			checkBoxControl.PlainText = "Reviewed"
'			#End Region ' #CreateFormControls
		End Sub
		Private Sub EditFormControls(ByVal workbook As IWorkbook)
'			#Region "#EditFormControls"
			workbook.LoadDocument("..\..\..\Documents\FormControls.xlsx")

			Dim formControls = workbook.Worksheets(0).FormControls

			For Each formControl As FormControl In formControls
				formControl.PrintObject = False
			Next formControl
'			#End Region ' #EditFormControls

		End Sub
	End Module
End Namespace
