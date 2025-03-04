using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetDocServerAPIPart2
{
    public static class FormControlActions
    {
        public static Action<Workbook> CreateFormControlsAction = CreateFormControls;
        public static Action<Workbook> EditFormControlsAction = EditFormControls;
        static void CreateFormControls(IWorkbook workbook)
        {
            #region #CreateFormControls
            var formControls = workbook.Worksheets[0].FormControls;

            // Create a button form control:
            var buttonCellRange = workbook.Worksheets[0].Range["B2:C2"];
            var buttonFormControl = formControls.AddButton(buttonCellRange);
            buttonFormControl.PlainText = "Click Here";

            // Create a list box form control:
            var comboCellRange = workbook.Worksheets[0].Range["B4:C4"];
            var comboBoxControl = formControls.AddComboBox(comboCellRange);
            comboBoxControl.DropDownLines = 3;
            comboBoxControl.SourceRange = workbook.Worksheets[0].Range["E2:E6"];
            comboBoxControl.SelectedIndex = 1;

            // Create a check box form control:
            var checkRange = workbook.Worksheets[0].Range["D5:E5"];
            var checkBoxControl = formControls.AddCheckBox(checkRange);
            checkBoxControl.CheckState = FormControlCheckState.Checked;
            checkBoxControl.PlainText = "Reviewed";
            #endregion #CreateFormControls
        }
        static void EditFormControls(IWorkbook workbook)
        {
            #region #EditFormControls
            workbook.LoadDocument("..\\..\\..\\Documents\\FormControls.xlsx");

            var formControls = workbook.Worksheets[0].FormControls;

            foreach (FormControl formControl in formControls)
            {
                formControl.PrintObject = false;
            }
            #endregion #EditFormControls

        }
    }
}
