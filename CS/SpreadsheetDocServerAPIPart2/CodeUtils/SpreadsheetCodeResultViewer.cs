using Microsoft.CSharp;
using System.CodeDom.Compiler;

namespace SpreadsheetDocServerAPIPart2
{
    public abstract class SpreadsheetExampleCodeEvaluator : ExampleCodeEvaluator
    {
        protected override string GetModuleAssembly()
        {
            return AssemblyInfo.SRAssemblySpreadsheetCore;
        }
        protected override string GetExampleClassName()
        {
            return "SpreadsheetCodeResultViewer.ExampleItem";
        }
    }
    #region SpreadsheetCSExampleCodeEvaluator
    public partial class SpreadsheetCSExampleCodeEvaluator : SpreadsheetExampleCodeEvaluator
    {

        protected override CodeDomProvider GetCodeDomProvider()
        {
            return new CSharpCodeProvider();
        }

        const string codeStart =
      "using System;\r\n" +
      "using DevExpress.Spreadsheet;\r\n" +
      "using DevExpress.Spreadsheet.Charts;\r\n" +
      "using DevExpress.Spreadsheet.Drawings;\r\n" +
      "using DevExpress.XtraSpreadsheet.Export;\r\n" +
      "using System.Drawing;\r\n" +
      "using System.Windows.Forms;\r\n" +
      "using DevExpress.XtraPrinting;\r\n" +
      "using DevExpress.XtraPrinting.Control;\r\n" +
      "using DevExpress.Utils;\r\n" +
      "using System.IO;\r\n" +
      "using System.Diagnostics;\r\n" +
      "using System.Xml;\r\n" +
      "using System.Data;\r\n" +
      "using System.Collections.Generic;\r\n" +
      "using System.Globalization;\r\n" +
      "using Formatting = DevExpress.Spreadsheet.Formatting;\r\n" +
      "namespace SpreadsheetCodeResultViewer { \r\n" +
      "public class ExampleItem { \r\n" +
      "        public static void Process(Workbook workbook) { \r\n" +
      "\r\n";

        const string codeEnd =
        "       \r\n }\r\n" +
        "    }\r\n" +
        "}\r\n";
        protected override string CodeStart { get { return codeStart; } }
        protected override string CodeEnd { get { return codeEnd; } }
    }
    #endregion
    #region SpreadsheetVbExampleCodeEvaluator
    public partial class SpreadsheetVbExampleCodeEvaluator : SpreadsheetExampleCodeEvaluator
    {

        protected override CodeDomProvider GetCodeDomProvider()
        {
            return new Microsoft.VisualBasic.VBCodeProvider();
        }
        const string codeStart =
      "Imports Microsoft.VisualBasic\r\n" +
      "Imports System\r\n" +
      "Imports DevExpress.Spreadsheet\r\n" +
      "Imports DevExpress.Spreadsheet.Charts\r\n" +
      "Imports DevExpress.Spreadsheet.Drawings\r\n" +
      "Imports DevExpress.XtraSpreadsheet.Export\r\n" +
      "Imports System.Drawing\r\n" +
      "Imports System.Windows.Forms\r\n" +
      "Imports DevExpress.XtraPrinting\r\n" +
      "Imports DevExpress.XtraPrinting.Control\r\n" +
      "Imports DevExpress.Utils\r\n" +
      "Imports System.IO\r\n" +
      "Imports System.Diagnostics\r\n" +
      "Imports System.Xml\r\n" +
      "Imports System.Data\r\n" +
      "Imports System.Collections.Generic\r\n" +
      "Imports System.Globalization\r\n" +
      "Imports Formatting = DevExpress.Spreadsheet.Formatting\r\n" +
      "Namespace SpreadsheetCodeResultViewer\r\n" +
      "	Public Class ExampleItem\r\n" +
      "		Public Shared Sub Process(ByVal workbook As Workbook)\r\n" +
      "\r\n";

        const string codeEnd =
        "\r\n		End Sub\r\n" +
        "	End Class\r\n" +
        "End Namespace\r\n";

        protected override string CodeStart { get { return codeStart; } }
        protected override string CodeEnd { get { return codeEnd; } }
    }
    #endregion
}
