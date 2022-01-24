using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace SpreadsheetDocServerAPIPart2
{
    public static class SearchActions
    {
        static void SimpleSearchValue(Workbook workbook)
        {
            #region #SimpleSearch
            workbook.Calculate();
            Worksheet worksheet = workbook.Worksheets["ExpenseReport"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Find and highlight cells that contain the word "holiday".
            IEnumerable<Cell> searchResult = worksheet.Search("holiday");
            foreach (Cell cell in searchResult)
                cell.Fill.BackgroundColor = Color.LightGreen;
            #endregion #SimpleSearch
        }

        static void AdvancedSearchValue(Workbook workbook)
        {
            #region #AdvancedSearch
            workbook.Calculate();
            Worksheet worksheet = workbook.Worksheets["ExpenseReport"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Specify the search term.
            string searchString = DateTime.Today.ToString("d");

            // Specify search options.
            SearchOptions options = new SearchOptions();
            options.SearchBy = SearchBy.Columns;
            options.SearchIn = SearchIn.Values;
            options.MatchEntireCellContents = true;

            // Find and highlight all cells that contain today's date.
            IEnumerable<Cell> searchResult = worksheet.Search(searchString, options);
            foreach (Cell cell in searchResult)
                cell.Fill.BackgroundColor = Color.LightGreen;
            #endregion #AdvancedSearch
        }
    }
}
