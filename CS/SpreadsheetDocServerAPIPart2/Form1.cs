using DevExpress.Spreadsheet;
using System;
using System.Diagnostics;

namespace SpreadsheetDocServerAPIPart2
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        #region #CreateWorkbook
        // Create a new Workbook object.
        Workbook workbook = new Workbook();
        #endregion #CreateWorkbook

        public Form1()
        {
            InitializeComponent();
            InitTreeListControl();
            workbook.Options.CalculationMode = WorkbookCalculationMode.Automatic;
        }

        void InitTreeListControl()
        {
            GroupsOfSpreadsheetExamples examples = new GroupsOfSpreadsheetExamples();
            InitData(examples);
            DataBinding(examples);
        }

        void InitData(GroupsOfSpreadsheetExamples examples)
        {
            #region GroupNodes
            examples.Add(new SpreadsheetNode("Auto Filter"));
            examples.Add(new SpreadsheetNode("Custom XML Parts"));
            examples.Add(new SpreadsheetNode("Data Validation"));
            examples.Add(new SpreadsheetNode("Export"));
            examples.Add(new SpreadsheetNode("Form Controls"));
            examples.Add(new SpreadsheetNode("Group and Outline"));
            examples.Add(new SpreadsheetNode("Pictures"));
            examples.Add(new SpreadsheetNode("Protection"));
            examples.Add(new SpreadsheetNode("Search"));
            examples.Add(new SpreadsheetNode("Sorting"));
            examples.Add(new SpreadsheetNode("Tables"));
            #endregion

            #region ExampleNodes
            // Add nodes to the "Filter" group of examples.
            examples[0].Groups.Add(new SpreadsheetExample("Apply Filter", AutoFilterActions.ApplyFilterAction));
            examples[0].Groups.Add(new SpreadsheetExample("Sort and Filter by Single Column", AutoFilterActions.FilterAndSortBySingleColumnAction));
            examples[0].Groups.Add(new SpreadsheetExample("Sort and Filter by Multiple Columns", AutoFilterActions.FilterAndSortByMultipleColumnsAction));
            examples[0].Groups.Add(new SpreadsheetExample("Filter by Value", AutoFilterActions.FilterByValueAction));
            examples[0].Groups.Add(new SpreadsheetExample("Filter by Multiple Values", AutoFilterActions.FilterByMultipleValuesAction)); examples[0].Groups.Add(new SpreadsheetExample("Numeric Filter by Condition", AutoFilterActions.FilterNumericByConditionAction));
            examples[0].Groups.Add(new SpreadsheetExample("Text Filter by Condition", AutoFilterActions.FilterTextByConditionAction));
            examples[0].Groups.Add(new SpreadsheetExample("Date Filter By Condition", AutoFilterActions.FilterDatesByConditionAction));
            examples[0].Groups.Add(new SpreadsheetExample("Filter Mixed Data Types by Values", AutoFilterActions.FilterMixedDataTypesByValuesAction));
            examples[0].Groups.Add(new SpreadsheetExample("Top 10 Filter", AutoFilterActions.Top10FilterValueAction));
            examples[0].Groups.Add(new SpreadsheetExample("Dynamic Filter", AutoFilterActions.DynamicFilterValueAction));
            examples[0].Groups.Add(new SpreadsheetExample("Sort and Filter by Color", AutoFilterActions.FilterAndSortByColorAction));
            examples[0].Groups.Add(new SpreadsheetExample("Filter by Background Color", AutoFilterActions.FilterByBackgroundColorAction));
            examples[0].Groups.Add(new SpreadsheetExample("Filter by Fill Color", AutoFilterActions.FilterByFillColorAction)); examples[0].Groups.Add(new SpreadsheetExample("Numeric Filter by Condition", AutoFilterActions.FilterNumericByConditionAction));
            examples[0].Groups.Add(new SpreadsheetExample("Filter by Font Color", AutoFilterActions.FilterByFontColorAction));
            examples[0].Groups.Add(new SpreadsheetExample("Reapply Filter", AutoFilterActions.ReapplyFilterValueAction));
            examples[0].Groups.Add(new SpreadsheetExample("Clear Filter", AutoFilterActions.ClearFilterAction));
            examples[0].Groups.Add(new SpreadsheetExample("Disable Filter", AutoFilterActions.DisableFilterAction));

            // Add nodes to the "Custom Xml Parts" group of examples.
            examples[1].Groups.Add(new SpreadsheetExample("Obtain Custom XML Parts", CustomXmlPartActions.ObtainCustomXmlPartAction));
            examples[1].Groups.Add(new SpreadsheetExample("Modify Custom XML Parts", CustomXmlPartActions.ModifyCustomXmlPartAction));
            examples[1].Groups.Add(new SpreadsheetExample("Store Custom XML Parts", CustomXmlPartActions.StoreCustomXmlPartAction));

            // Add nodes to the "Data Validation" group of examples.
            examples[2].Groups.Add(new SpreadsheetExample("Add Data Validation", DataValidationActions.AddDataValidationAction));
            examples[2].Groups.Add(new SpreadsheetExample("Change Validation Criteria", DataValidationActions.ChangeCriteriaAction));
            examples[2].Groups.Add(new SpreadsheetExample("Get Data Validation", DataValidationActions.GetDataValidationAction));
            examples[2].Groups.Add(new SpreadsheetExample("Validate Cell Value", DataValidationActions.ValidateCellValueAction));
            examples[2].Groups.Add(new SpreadsheetExample("Show Error Message", DataValidationActions.ShowErrorMessageAction));
            examples[2].Groups.Add(new SpreadsheetExample("Show Input Message", DataValidationActions.ShowInputMessageAction));
            examples[2].Groups.Add(new SpreadsheetExample("Use Union Range", DataValidationActions.UseUnionRangeAction));
            examples[2].Groups.Add(new SpreadsheetExample("Remove Data Validation", DataValidationActions.RemoveDataValidationAction));
            examples[2].Groups.Add(new SpreadsheetExample("Remove All Data Validations", DataValidationActions.RemoveAllDataValidationsAction));

            // Add nodes to the "Export" group of examples.
            examples[3].Groups.Add(new SpreadsheetExample("Export to HTML", ExportActions.ExportDocToHTMLAction));

            // Add nodes to the "Form Controls" group of examples.
            examples[4].Groups.Add(new SpreadsheetExample("Create Form Controls", FormControlActions.CreateFormControlsAction));
            examples[4].Groups.Add(new SpreadsheetExample("Edit Form Controls", FormControlActions.EditFormControlsAction));

            // Add nodes to the "Group and Outline" group of examples.
            examples[5].Groups.Add(new SpreadsheetExample("Group Rows", GroupAndOutlineActions.GroupRowsAction));
            examples[5].Groups.Add(new SpreadsheetExample("Ungroup Rows", GroupAndOutlineActions.UngroupRowsAction));
            examples[5].Groups.Add(new SpreadsheetExample("Group Columns", GroupAndOutlineActions.GroupColumnsAction));
            examples[5].Groups.Add(new SpreadsheetExample("Ungroup Columns", GroupAndOutlineActions.UngroupColumnsAction));
            examples[5].Groups.Add(new SpreadsheetExample("Auto Outline", GroupAndOutlineActions.AutoOutlineAction));
            examples[5].Groups.Add(new SpreadsheetExample("Subtotal", GroupAndOutlineActions.SubtotalAction));

            // Add nodes to the "Pictures" group of examples. 
            examples[6].Groups.Add(new SpreadsheetExample("Insert a Picture", PictureActions.InsertPictureAction));
            examples[6].Groups.Add(new SpreadsheetExample("Modify a Picture", PictureActions.ModifyPictureAction));
            examples[6].Groups.Add(new SpreadsheetExample("Place Picture In Cell", PictureActions.PlacePictureInCellAction));

            // Add nodes to the "Protection" group of examples.
            examples[7].Groups.Add(new SpreadsheetExample("Protect Workbook", ProtectionActions.ProtectWorkbookAction));
            examples[7].Groups.Add(new SpreadsheetExample("Protect Worksheet", ProtectionActions.ProtectWorksheetAction));
            examples[7].Groups.Add(new SpreadsheetExample("Unprotect Workbook", ProtectionActions.UnprotectWorkbookAction));
            examples[7].Groups.Add(new SpreadsheetExample("Unprotect Worksheet", ProtectionActions.UnprotectWorksheetAction));
            examples[7].Groups.Add(new SpreadsheetExample("Protect Range", ProtectionActions.ProtectRangeAction));

            // Add nodes to the "Search" group of examples.
            examples[8].Groups.Add(new SpreadsheetExample("Simple Search", SearchActions.SimpleSearchValueAction));
            examples[8].Groups.Add(new SpreadsheetExample("Advanced Search", SearchActions.AdvancedSearchValueAction));

            // Add nodes to the "Sort" group of examples.
            examples[9].Groups.Add(new SpreadsheetExample("Simple Sort", SortActions.SimpleSortAction));
            examples[9].Groups.Add(new SpreadsheetExample("Sort in Descending Order", SortActions.DescendingOrderAction));
            examples[9].Groups.Add(new SpreadsheetExample("Sort by a Column", SortActions.SortBySpecifiedColumnAction));
            examples[9].Groups.Add(new SpreadsheetExample("Sort by Multiple Columns", SortActions.SortByMultipleColumnsAction));
            examples[9].Groups.Add(new SpreadsheetExample("Sort by Fill Color", SortActions.SortByFillColorAction));
            examples[9].Groups.Add(new SpreadsheetExample("Sort by Font Color", SortActions.SortByFontColorAction));


            // Add nodes to the "Tables" group of examples.
            examples[10].Groups.Add(new SpreadsheetExample("Create a Table", TableActions.CreateTableAction));
            examples[10].Groups.Add(new SpreadsheetExample("Format a Table", TableActions.FormatTableAction));
            examples[10].Groups.Add(new SpreadsheetExample("Duplicate Table Style", TableActions.DuplicateTableStyleAction));
            examples[10].Groups.Add(new SpreadsheetExample("Table Ranges", TableActions.TableRangesAction));
            examples[10].Groups.Add(new SpreadsheetExample("Custom Table Style", TableActions.CustomTableStyleAction));
            #endregion
        }

        void DataBinding(GroupsOfSpreadsheetExamples examples)
        {
            treeList1.DataSource = examples;
            treeList1.ExpandAll();
            treeList1.BestFitColumns();
        }


        private void btnOpenExcel_Click(object sender, EventArgs e)
        {
            LoadDocumentFromFile();
            SpreadsheetExample example = treeList1.GetDataRecordByNode(treeList1.FocusedNode) as SpreadsheetExample;
            if (example == null)
                return;
            Action<Workbook> action = example.Action;
            action(workbook);
            SaveDocumentToFile();
        }

        // ------------------- Load and Save a Document -------------------
        private void LoadDocumentFromFile()
        {
            #region #LoadDocumentFromFile
            // Load a workbook from the file.
            workbook.LoadDocument("Document.xlsx", DocumentFormat.OpenXml);
            #endregion #LoadDocumentFromFile
        }

        private void SaveDocumentToFile()
        {
            #region #SaveDocumentToFile
            // Save the modified document to the file.
            workbook.SaveDocument("SavedDocument.xlsx", DocumentFormat.OpenXml);
            #endregion #SaveDocumentToFile
            Process.Start(new ProcessStartInfo("SavedDocument.xlsx") { UseShellExecute = true });
        }
    }
}
