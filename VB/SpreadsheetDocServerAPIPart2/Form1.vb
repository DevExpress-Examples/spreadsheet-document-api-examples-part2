Imports DevExpress.Spreadsheet
Imports System
Imports System.Diagnostics

Namespace SpreadsheetDocServerAPIPart2
	Partial Public Class Form1
		Inherits DevExpress.XtraEditors.XtraForm

		#Region "#CreateWorkbook"
		' Create a new Workbook object.
		Private workbook As New Workbook()
		#End Region ' #CreateWorkbook

		Public Sub New()
			InitializeComponent()
			InitTreeListControl()
			workbook.Options.CalculationMode = WorkbookCalculationMode.Automatic
		End Sub

		Private Sub InitTreeListControl()
			Dim examples As New GroupsOfSpreadsheetExamples()
			InitData(examples)
			DataBinding(examples)
		End Sub

		Private Sub InitData(ByVal examples As GroupsOfSpreadsheetExamples)
'			#Region "GroupNodes"
			examples.Add(New SpreadsheetNode("Auto Filter"))
			examples.Add(New SpreadsheetNode("Custom XML Parts"))
			examples.Add(New SpreadsheetNode("Data Validation"))
			examples.Add(New SpreadsheetNode("Export"))
			examples.Add(New SpreadsheetNode("Form Controls"))
			examples.Add(New SpreadsheetNode("Group and Outline"))
			examples.Add(New SpreadsheetNode("Pictures"))
			examples.Add(New SpreadsheetNode("Protection"))
			examples.Add(New SpreadsheetNode("Search"))
			examples.Add(New SpreadsheetNode("Sorting"))
			examples.Add(New SpreadsheetNode("Tables"))
'			#End Region

'			#Region "ExampleNodes"
			' Add nodes to the "Filter" group of examples.
			examples(0).Groups.Add(New SpreadsheetExample("Apply Filter", AutoFilterActions.ApplyFilterAction))
			examples(0).Groups.Add(New SpreadsheetExample("Sort and Filter by Single Column", AutoFilterActions.FilterAndSortBySingleColumnAction))
			examples(0).Groups.Add(New SpreadsheetExample("Sort and Filter by Multiple Columns", AutoFilterActions.FilterAndSortByMultipleColumnsAction))
			examples(0).Groups.Add(New SpreadsheetExample("Filter by Value", AutoFilterActions.FilterByValueAction))
			examples(0).Groups.Add(New SpreadsheetExample("Filter by Multiple Values", AutoFilterActions.FilterByMultipleValuesAction))
			examples(0).Groups.Add(New SpreadsheetExample("Numeric Filter by Condition", AutoFilterActions.FilterNumericByConditionAction))
			examples(0).Groups.Add(New SpreadsheetExample("Text Filter by Condition", AutoFilterActions.FilterTextByConditionAction))
			examples(0).Groups.Add(New SpreadsheetExample("Date Filter By Condition", AutoFilterActions.FilterDatesByConditionAction))
			examples(0).Groups.Add(New SpreadsheetExample("Filter Mixed Data Types by Values", AutoFilterActions.FilterMixedDataTypesByValuesAction))
			examples(0).Groups.Add(New SpreadsheetExample("Top 10 Filter", AutoFilterActions.Top10FilterValueAction))
			examples(0).Groups.Add(New SpreadsheetExample("Dynamic Filter", AutoFilterActions.DynamicFilterValueAction))
			examples(0).Groups.Add(New SpreadsheetExample("Sort and Filter by Color", AutoFilterActions.FilterAndSortByColorAction))
			examples(0).Groups.Add(New SpreadsheetExample("Filter by Background Color", AutoFilterActions.FilterByBackgroundColorAction))
			examples(0).Groups.Add(New SpreadsheetExample("Filter by Fill Color", AutoFilterActions.FilterByFillColorAction))
			examples(0).Groups.Add(New SpreadsheetExample("Numeric Filter by Condition", AutoFilterActions.FilterNumericByConditionAction))
			examples(0).Groups.Add(New SpreadsheetExample("Filter by Font Color", AutoFilterActions.FilterByFontColorAction))
			examples(0).Groups.Add(New SpreadsheetExample("Reapply Filter", AutoFilterActions.ReapplyFilterValueAction))
			examples(0).Groups.Add(New SpreadsheetExample("Clear Filter", AutoFilterActions.ClearFilterAction))
			examples(0).Groups.Add(New SpreadsheetExample("Disable Filter", AutoFilterActions.DisableFilterAction))



			' Add nodes to the "Custom Xml Parts" group of examples.
			examples(1).Groups.Add(New SpreadsheetExample("Obtain Custom XML Parts", CustomXmlPartActions.ObtainCustomXmlPartAction))
			examples(1).Groups.Add(New SpreadsheetExample("Modify Custom XML Parts", CustomXmlPartActions.ModifyCustomXmlPartAction))
			examples(1).Groups.Add(New SpreadsheetExample("Store Custom XML Parts", CustomXmlPartActions.StoreCustomXmlPartAction))

			' Add nodes to the "Data Validation" group of examples.
			examples(2).Groups.Add(New SpreadsheetExample("Add Data Validation", DataValidationActions.AddDataValidationAction))
			examples(2).Groups.Add(New SpreadsheetExample("Change Validation Criteria", DataValidationActions.ChangeCriteriaAction))
			examples(2).Groups.Add(New SpreadsheetExample("Get Data Validation", DataValidationActions.GetDataValidationAction))
			examples(2).Groups.Add(New SpreadsheetExample("Validate Cell Value", DataValidationActions.ValidateCellValueAction))
			examples(2).Groups.Add(New SpreadsheetExample("Show Error Message", DataValidationActions.ShowErrorMessageAction))
			examples(2).Groups.Add(New SpreadsheetExample("Show Input Message", DataValidationActions.ShowInputMessageAction))
			examples(2).Groups.Add(New SpreadsheetExample("Use Union Range", DataValidationActions.UseUnionRangeAction))
			examples(2).Groups.Add(New SpreadsheetExample("Remove Data Validaiton", DataValidationActions.RemoveDataValidationAction))
			examples(2).Groups.Add(New SpreadsheetExample("Remove All Data Validations", DataValidationActions.RemoveAllDataValidationsAction))

			' Add nodes to the "Export" group of examples.
			examples(3).Groups.Add(New SpreadsheetExample("Export to HTML", ExportActions.ExportDocToHTMLAction))

			' Add nodes to the "Form Controls" group of examples.
			examples(4).Groups.Add(New SpreadsheetExample("Create Form Controls", FormControlActions.CreateFormControlsAction))
			examples(4).Groups.Add(New SpreadsheetExample("Edit Form Controls", FormControlActions.EditFormControlsAction))

			' Add nodes to the "Group and Outline" group of examples.
			examples(5).Groups.Add(New SpreadsheetExample("Group Rows", GroupAndOutlineActions.GroupRowsAction))
			examples(5).Groups.Add(New SpreadsheetExample("Ungroup Rows", GroupAndOutlineActions.UngroupRowsAction))
			examples(5).Groups.Add(New SpreadsheetExample("Group Columns", GroupAndOutlineActions.GroupColumnsAction))
			examples(5).Groups.Add(New SpreadsheetExample("Ungroup Columns", GroupAndOutlineActions.UngroupColumnsAction))
			examples(5).Groups.Add(New SpreadsheetExample("Auto Outline", GroupAndOutlineActions.AutoOutlineAction))
			examples(5).Groups.Add(New SpreadsheetExample("Subtotal", GroupAndOutlineActions.SubtotalAction))

			' Add nodes to the "Pictures" group of examples. 
			examples(6).Groups.Add(New SpreadsheetExample("Insert a Picture", PictureActions.InsertPictureAction))
			examples(6).Groups.Add(New SpreadsheetExample("Modify a Picture", PictureActions.ModifyPictureAction))
			examples(6).Groups.Add(New SpreadsheetExample("Place Picture In Cell", PictureActions.PlacePictureInCellAction))

			' Add nodes to the "Protection" group of examples.
			examples(7).Groups.Add(New SpreadsheetExample("Protect Workbook", ProtectionActions.ProtectWorkbookAction))
			examples(7).Groups.Add(New SpreadsheetExample("Protect Worksheet", ProtectionActions.ProtectWorksheetAction))
			examples(7).Groups.Add(New SpreadsheetExample("Unprotect Workbook", ProtectionActions.UnprotectWorkbookAction))
			examples(7).Groups.Add(New SpreadsheetExample("Unprotect Worksheet", ProtectionActions.UnprotectWorksheetAction))
			examples(7).Groups.Add(New SpreadsheetExample("Protect Range", ProtectionActions.ProtectRangeAction))

			' Add nodes to the "Search" group of examples.
			examples(8).Groups.Add(New SpreadsheetExample("Simple Search", SearchActions.SimpleSearchValueAction))
			examples(8).Groups.Add(New SpreadsheetExample("Advanced Search", SearchActions.AdvancedSearchValueAction))

			' Add nodes to the "Sort" group of examples.
			examples(9).Groups.Add(New SpreadsheetExample("Simple Sort", SortActions.SimpleSortAction))
			examples(9).Groups.Add(New SpreadsheetExample("Sort in Descending Order", SortActions.DescendingOrderAction))
			examples(9).Groups.Add(New SpreadsheetExample("Sort by a Column", SortActions.SortBySpecifiedColumnAction))
			examples(9).Groups.Add(New SpreadsheetExample("Sort by Multiple Columns", SortActions.SortByMultipleColumnsAction))
			examples(9).Groups.Add(New SpreadsheetExample("Sort by Fill Color", SortActions.SortByFillColorAction))
			examples(9).Groups.Add(New SpreadsheetExample("Sort by Font Color", SortActions.SortByFontColorAction))


			' Add nodes to the "Tables" group of examples.
			examples(10).Groups.Add(New SpreadsheetExample("Create a Table", TableActions.CreateTableAction))
			examples(10).Groups.Add(New SpreadsheetExample("Format a Table", TableActions.FormatTableAction))
			examples(10).Groups.Add(New SpreadsheetExample("Duplicate Table Style", TableActions.DuplicateTableStyleAction))
			examples(10).Groups.Add(New SpreadsheetExample("Table Ranges", TableActions.TableRangesAction))
			examples(10).Groups.Add(New SpreadsheetExample("Custom Table Style", TableActions.CustomTableStyleAction))
'			#End Region
		End Sub

		Private Sub DataBinding(ByVal examples As GroupsOfSpreadsheetExamples)
			treeList1.DataSource = examples
			treeList1.ExpandAll()
			treeList1.BestFitColumns()
		End Sub


		Private Sub btnOpenExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnOpenExcel.Click
			LoadDocumentFromFile()
			Dim example As SpreadsheetExample = TryCast(treeList1.GetDataRecordByNode(treeList1.FocusedNode), SpreadsheetExample)
			If example Is Nothing Then
				Return
			End If
			Dim action As Action(Of Workbook) = example.Action
			action(workbook)
			SaveDocumentToFile()
		End Sub

		' ------------------- Load and Save a Document -------------------
		Private Sub LoadDocumentFromFile()
'			#Region "#LoadDocumentFromFile"
			' Load a workbook from the file.
			workbook.LoadDocument("Document.xlsx", DocumentFormat.OpenXml)
'			#End Region ' #LoadDocumentFromFile
		End Sub

		Private Sub SaveDocumentToFile()
'			#Region "#SaveDocumentToFile"
			' Save the modified document to the file.
			workbook.SaveDocument("SavedDocument.xlsx", DocumentFormat.OpenXml)
'			#End Region ' #SaveDocumentToFile
			Process.Start(New ProcessStartInfo("SavedDocument.xlsx") With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
