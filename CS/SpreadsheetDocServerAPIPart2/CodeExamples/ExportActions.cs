using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet.Export;
using System.IO;

namespace SpreadsheetDocServerAPIPart2
{
    public static class ExportActions
    {
        private static void ExportDocToHTML(Workbook workbook)
        {
            #region #ExportToHTML
            Worksheet worksheet = workbook.Worksheets["Grouping"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            HtmlDocumentExporterOptions options = new HtmlDocumentExporterOptions();

            // Specify the part of the document to be exported to HTML.
            options.SheetIndex = worksheet.Index;
            options.Range = "B2:G7";

            // Export the active worksheet to a stream as HTML with the specified options.
            using (FileStream htmlStream = new FileStream("OutputWorksheet.html", FileMode.Create))
            {
                workbook.ExportToHtml(htmlStream, options);
            }

            System.Diagnostics.Process.Start("OutputWorksheet.html");

            #endregion #ExportToHTML
        }
    }
}
