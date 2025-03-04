using System;
using DevExpress.Spreadsheet;
using System.Drawing;
using System.Linq;
using DevExpress.XtraCharts;
using DevExpress.XtraExport.Helpers;

namespace SpreadsheetDocServerAPIPart2
{
    public static class DataValidationActions {

        public static Action<Workbook> AddDataValidationAction = AddDataValidation;
        public static Action<Workbook> ChangeCriteriaAction = ChangeCriteria;
        public static Action<Workbook> UseUnionRangeAction = UseUnionRange;
        public static Action<Workbook> ShowInputMessageAction = ShowInputMessage;
        public static Action<Workbook> ShowErrorMessageAction = ShowErrorMessage;
        public static Action<Workbook> GetDataValidationAction = GetDataValidation;
        public static Action<Workbook> ValidateCellValueAction = ValidateCellValue;
        public static Action<Workbook> RemoveDataValidationAction = RemoveDataValidation;
        public static Action<Workbook> RemoveAllDataValidationsAction = RemoveAllDataValidations;

        static void AddDataValidation(IWorkbook workbook) {
            #region #AddDataValidation
            // Load a document from a file.
            workbook.LoadDocument("..\\..\\..\\Documents\\DataValidation.xlsx");

            // Access a worksheet.
            Worksheet worksheet = workbook.Worksheets[0];
            
            // Specify the "C1" cell value.
            worksheet["C1"].SetValue(DateTime.Now);
            worksheet["C1"].NumberFormat = "mmm/d/yyyy h:mm";

            // Restrict data entry to a whole number from 10 to 20.
            worksheet.DataValidations.Add(worksheet["B1"], 
                DataValidationType.WholeNumber, DataValidationOperator.Between, 10, 20);

            // Restrict data entry to a number within limits.
            DataValidation validation = worksheet.DataValidations.Add(worksheet["F4:F11"], DataValidationType.Decimal, DataValidationOperator.Between, 10, 40);

            // Restrict data entry based on a formula.
            worksheet.DataValidations.Add(worksheet["B4:B11"], DataValidationType.Custom, "=AND(ISNUMBER(B4),LEN(B4)=5)");

            // Restrict data entry to 3 symbols.
            worksheet.DataValidations.Add(worksheet["D4:D11"], DataValidationType.TextLength, DataValidationOperator.Equal, 3);

            // Restrict data entry to values in a drop-down list specified in code. 
            // Note that the list in code uses comma to separate entries, 
            // but the list in UI displays a culture-specific list separator.
            worksheet.DataValidations.Add(worksheet["A4:A11"], DataValidationType.List, "PASS, FAIL");

            // Restrict data entry to values in a drop-down list obtained from a worksheet.
            worksheet.DataValidations.Add(worksheet["E4:E11"], DataValidationType.List, ValueObject.FromRange(worksheet["H4:H9"].GetRangeWithAbsoluteReference()));
            
            // Restrict data entry to a time before the specified time.
            worksheet.DataValidations.Add(worksheet["C1"], DataValidationType.Time, DataValidationOperator.LessThanOrEqual, DateTime.Now);

            // Highlight data validation ranges.
            worksheet["H4:H9"].FillColor = Color.LightGray;
            int[] MyColorScheme = new int[] { 0xFFC4C4, 0xFFD9D9, 0xFFF6F6, 0xFFECEC, 0xE9D3D3, 0xFFDFC4, 0xFFDAE9};
            for (int i = 0; i < worksheet.DataValidations.Count; i++){
                worksheet.DataValidations[i].Range.FillColor = Color.FromArgb(MyColorScheme[i]);
            }
            #endregion #AddDataValidation
        }

        static void ChangeCriteria(IWorkbook workbook) {
            #region #ChangeCriteria
            // Load a document from a file.
            workbook.LoadDocument("..\\..\\..\\Documents\\DataValidation.xlsx");

            // Access a worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Restrict data entry to a number within limits.
            DataValidation validation = worksheet.DataValidations.Add
                (worksheet["F4:F11"], DataValidationType.Decimal, DataValidationOperator.Between, 10, 40);

            // Change the validation operator and criteria.
            // The "F4:F11" range's cell values are valid
            // if they are greater than or equal 20.
            validation.Operator = DataValidationOperator.GreaterThanOrEqual;
            validation.Criteria = 20;
            validation.Criteria2 = ValueObject.Empty;

            // Highlight data validation ranges.
            int[] MyColorScheme = new int[] { 0xFFC4C4, 0xFFD9D9, 0xFFF6F6, 0xFFECEC, 0xE9D3D3 };
            for (int i = 0; i < worksheet.DataValidations.Count; i++)
            {
                worksheet.DataValidations[i].Range.FillColor = Color.FromArgb(MyColorScheme[i]);
            }
            #endregion #ChangeCriteria
        }

        static void UseUnionRange(IWorkbook workbook) {
            #region #UseUnionRange
            // Load a document from a file.
            workbook.LoadDocument("..\\..\\..\\Documents\\DataValidation.xlsx");

            // Access a worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a union range.
            CellRange range = worksheet.Range.Union(worksheet["F4:F5"], worksheet["F6:F11"]);
            
            // Restrict data entry to a number within limits.
            worksheet.DataValidations.Add(range, DataValidationType.Decimal, DataValidationOperator.Between, 10, 40);

            // Highlight data validation ranges.
            int[] MyColorScheme = new int[] { 0xFFC4C4, 0xFFD9D9, 0xFFF6F6, 0xFFECEC, 0xE9D3D3 };
            for (int i = 0; i < worksheet.DataValidations.Count; i++)
            {
                worksheet.DataValidations[i].Range.FillColor = Color.FromArgb(MyColorScheme[i]);
            }
            #endregion #UseUnionRange
        }

        static void ShowInputMessage(IWorkbook workbook) {
            #region #ShowInputMessage
            // Load a document from a file.
            workbook.LoadDocument("..\\..\\..\\Documents\\DataValidation.xlsx");

            // Access a worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Restrict data entry to a 5-digit number.
            DataValidation validation = worksheet.DataValidations.Add(worksheet["B4:B11"], DataValidationType.Custom, "=AND(ISNUMBER(B4),LEN(B4)=5)");

            // Show input message.
            validation.InputTitle = "Employee Id";
            validation.InputMessage = "Please enter 5-digit number";
            validation.ShowInputMessage = true;

            // Highlight data validation ranges.
            int[] MyColorScheme = new int[] { 0xFFC4C4, 0xFFD9D9, 0xFFF6F6, 0xFFECEC, 0xE9D3D3 };
            for (int i = 0; i < worksheet.DataValidations.Count; i++)
            {
                worksheet.DataValidations[i].Range.FillColor = Color.FromArgb(MyColorScheme[i]);
            }
            #endregion #ShowInputMessage
        }

        static void ShowErrorMessage(IWorkbook workbook) {
            #region #ShowErrorMessage
            // Load a document from a file.
            workbook.LoadDocument("..\\..\\..\\Documents\\DataValidation.xlsx");

            // Access a worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Restrict data entry to a 5-digit number.
            DataValidation validation = worksheet.DataValidations.Add(worksheet["B4:B11"], DataValidationType.Custom, "=AND(ISNUMBER(B4),LEN(B4)=5)");

            // Show error message.
            validation.ErrorTitle = "Wrong Employee Id";
            validation.ErrorMessage = "The value you entered is not valid. Use 5-digit number for the employee ID.";
            validation.ErrorStyle = DataValidationErrorStyle.Information;
            validation.ShowErrorMessage = true;

            // Highlight data validation ranges.
            int[] MyColorScheme = new int[] { 0xFFC4C4, 0xFFD9D9, 0xFFF6F6, 0xFFECEC, 0xE9D3D3 };
            for (int i = 0; i < worksheet.DataValidations.Count; i++)
            {
                worksheet.DataValidations[i].Range.FillColor = Color.FromArgb(MyColorScheme[i]);
            }
            #endregion #ShowErrorMessage
        }

        static void GetDataValidation(IWorkbook workbook)
        {
            #region #GetDataValidation
            // Load a document from a file.
            workbook.LoadDocument("..\\..\\..\\Documents\\DataValidation.xlsx");

            // Access a worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Add data validations.
            worksheet.DataValidations.Add(worksheet["D4:D11"], DataValidationType.TextLength, DataValidationOperator.Equal, 3);
            worksheet.DataValidations.Add(worksheet["E4:E11"], DataValidationType.List, ValueObject.FromRange(worksheet["H4:H9"].GetRangeWithAbsoluteReference()));

            // Get data validation entry associated with the "E4" cell.
            worksheet.DataValidations.GetDataValidation(worksheet.Cells["E4"]).Criteria = ValueObject.FromRange(worksheet["H4:H5"]);

            // Get data validation entries for the "D4:E11" range.
            var myValidation = worksheet.DataValidations.GetDataValidations(worksheet["D4:E11"])
                .Where(d => d.ValidationType == DataValidationType.TextLength).SingleOrDefault();
            if (myValidation != null) myValidation.Criteria = 4;
            
            // Get data validation entries that meet a specific criteria.
            foreach (var d in worksheet.DataValidations.GetDataValidations(DataValidationType.TextLength, DataValidationOperator.Equal, 4, ValueObject.Empty))
            {
                // Change the validation operator.
                // The "D4:D11" range's cell values are valid
                // if they contain text with more than 4 characters.
                d.Operator = DataValidationOperator.GreaterThan;
            }              

            // Highlight data validation ranges.
            int[] MyColorScheme = new int[] { 0xFFC4C4, 0xFFD9D9, 0xFFF6F6, 0xFFECEC, 0xE9D3D3 };
            for (int i = 0; i < worksheet.DataValidations.Count; i++)
            {
                worksheet.DataValidations[i].Range.FillColor = Color.FromArgb(MyColorScheme[i]);
            }
            #endregion #GetDataValidation
        }

        static void ValidateCellValue(IWorkbook workbook)
        {
            #region #ValidateCellValue
            // Load a document from a file.
            workbook.LoadDocument("..\\..\\..\\Documents\\DataValidation.xlsx");

            // Access a worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Add data validations.
            worksheet.DataValidations.Add(worksheet["D4:D11"], DataValidationType.TextLength, DataValidationOperator.Equal, 3);
            
            // Check whether the cell value meets a validation criteria.
            bool isValid = worksheet.DataValidations.Validate(worksheet.Cells["D4"], worksheet.Cells["J4"].Value);
            if (isValid)  { worksheet["D4"].CopyFrom(worksheet["J4"]); }
            #endregion #ValidateCellValue
        }

        static void RemoveDataValidation(IWorkbook workbook) {
            #region #RemoveDataValidation
            // Load a document from a file.
            workbook.LoadDocument("..\\..\\..\\Documents\\DataValidation.xlsx");
            
            // Access a worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Add data validations.
            worksheet.DataValidations.Add(worksheet["D4:D11"], DataValidationType.TextLength, DataValidationOperator.Equal, 3);
            worksheet.DataValidations.Add(worksheet["E4:E11"], DataValidationType.List, ValueObject.FromRange(worksheet["H4:H9"].GetRangeWithAbsoluteReference()));

            // Remove a data validation by its index.
            worksheet.DataValidations.RemoveAt(1);

            // Highlight data validation ranges.
            int[] MyColorScheme = new int[] { 0xFFC4C4, 0xFFD9D9, 0xFFF6F6, 0xFFECEC, 0xE9D3D3 };
            for (int i = 0; i < worksheet.DataValidations.Count; i++)
            {
                worksheet.DataValidations[i].Range.FillColor = Color.FromArgb(MyColorScheme[i]);
            }
            #endregion #RemoveDataValidation
        }

        static void RemoveAllDataValidations(IWorkbook workbook) {
            #region #RemoveAllDataValidations
            // Load a document from a file.
            workbook.LoadDocument("..\\..\\..\\Documents\\DataValidation.xlsx");

            // Access a worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Add data validations.
            worksheet.DataValidations.Add(worksheet["D4:D11"], DataValidationType.TextLength, DataValidationOperator.Equal, 3);
            worksheet.DataValidations.Add(worksheet["E4:E11"], DataValidationType.List, ValueObject.FromRange(worksheet["H4:H9"].GetRangeWithAbsoluteReference()));

            // Remove all data validations.
            worksheet.DataValidations.Clear();

            // Highlight data validation ranges.
            int[] MyColorScheme = new int[] { 0xFFC4C4, 0xFFD9D9, 0xFFF6F6, 0xFFECEC, 0xE9D3D3 };
            for (int i = 0; i < worksheet.DataValidations.Count; i++)
            {
                worksheet.DataValidations[i].Range.FillColor = Color.FromArgb(MyColorScheme[i]);
            }
            #endregion #RemoveAllDataValidations
        }
    }
}
