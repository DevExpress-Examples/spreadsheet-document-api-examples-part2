Imports Microsoft.CSharp
Imports System.CodeDom.Compiler

Namespace SpreadsheetDocServerAPIPart2

    Public MustInherit Class SpreadsheetExampleCodeEvaluator
        Inherits ExampleCodeEvaluator

        Protected Overrides Function GetModuleAssembly() As String
            Return AssemblyInfo.SRAssemblySpreadsheetCore
        End Function

        Protected Overrides Function GetExampleClassName() As String
            Return "SpreadsheetCodeResultViewer.ExampleItem"
        End Function
    End Class

#Region "SpreadsheetCSExampleCodeEvaluator"
    Public Partial Class SpreadsheetCSExampleCodeEvaluator
        Inherits SpreadsheetExampleCodeEvaluator

        Protected Overrides Function GetCodeDomProvider() As CodeDomProvider
            Return New CSharpCodeProvider()
        End Function

        Const codeStartField As String = "using System;" & Microsoft.VisualBasic.Constants.vbCrLf & "using DevExpress.Spreadsheet;" & Microsoft.VisualBasic.Constants.vbCrLf & "using DevExpress.Spreadsheet.Charts;" & Microsoft.VisualBasic.Constants.vbCrLf & "using DevExpress.Spreadsheet.Drawings;" & Microsoft.VisualBasic.Constants.vbCrLf & "using DevExpress.XtraSpreadsheet.Export;" & Microsoft.VisualBasic.Constants.vbCrLf & "using System.Drawing;" & Microsoft.VisualBasic.Constants.vbCrLf & "using System.Windows.Forms;" & Microsoft.VisualBasic.Constants.vbCrLf & "using DevExpress.XtraPrinting;" & Microsoft.VisualBasic.Constants.vbCrLf & "using DevExpress.XtraPrinting.Control;" & Microsoft.VisualBasic.Constants.vbCrLf & "using DevExpress.Utils;" & Microsoft.VisualBasic.Constants.vbCrLf & "using System.IO;" & Microsoft.VisualBasic.Constants.vbCrLf & "using System.Diagnostics;" & Microsoft.VisualBasic.Constants.vbCrLf & "using System.Xml;" & Microsoft.VisualBasic.Constants.vbCrLf & "using System.Data;" & Microsoft.VisualBasic.Constants.vbCrLf & "using System.Collections.Generic;" & Microsoft.VisualBasic.Constants.vbCrLf & "using System.Globalization;" & Microsoft.VisualBasic.Constants.vbCrLf & "using Formatting = DevExpress.Spreadsheet.Formatting;" & Microsoft.VisualBasic.Constants.vbCrLf & "namespace SpreadsheetCodeResultViewer { " & Microsoft.VisualBasic.Constants.vbCrLf & "public class ExampleItem { " & Microsoft.VisualBasic.Constants.vbCrLf & "        public static void Process(Workbook workbook) { " & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf

        Const codeEndField As String = "       " & Microsoft.VisualBasic.Constants.vbCrLf & " }" & Microsoft.VisualBasic.Constants.vbCrLf & "    }" & Microsoft.VisualBasic.Constants.vbCrLf & "}" & Microsoft.VisualBasic.Constants.vbCrLf

        Protected Overrides ReadOnly Property CodeStart As String
            Get
                Return codeStartField
            End Get
        End Property

        Protected Overrides ReadOnly Property CodeEnd As String
            Get
                Return codeEndField
            End Get
        End Property
    End Class

#End Region
#Region "SpreadsheetVbExampleCodeEvaluator"
    Public Partial Class SpreadsheetVbExampleCodeEvaluator
        Inherits SpreadsheetExampleCodeEvaluator

        Protected Overrides Function GetCodeDomProvider() As CodeDomProvider
            Return New Microsoft.VisualBasic.VBCodeProvider()
        End Function

        Const codeStartField As String = "Imports Microsoft.VisualBasic" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports System" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports DevExpress.Spreadsheet" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports DevExpress.Spreadsheet.Charts" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports DevExpress.Spreadsheet.Drawings" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports DevExpress.XtraSpreadsheet.Export" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports System.Drawing" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports System.Windows.Forms" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports DevExpress.XtraPrinting" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports DevExpress.XtraPrinting.Control" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports DevExpress.Utils" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports System.IO" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports System.Diagnostics" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports System.Xml" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports System.Data" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports System.Collections.Generic" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports System.Globalization" & Microsoft.VisualBasic.Constants.vbCrLf & "Imports Formatting = DevExpress.Spreadsheet.Formatting" & Microsoft.VisualBasic.Constants.vbCrLf & "Namespace SpreadsheetCodeResultViewer" & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbTab & "Public Class ExampleItem" & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbTab & Microsoft.VisualBasic.Constants.vbTab & "Public Shared Sub Process(ByVal workbook As Workbook)" & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf

        Const codeEndField As String = Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbTab & Microsoft.VisualBasic.Constants.vbTab & "End Sub" & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbTab & "End Class" & Microsoft.VisualBasic.Constants.vbCrLf & "End Namespace" & Microsoft.VisualBasic.Constants.vbCrLf

        Protected Overrides ReadOnly Property CodeStart As String
            Get
                Return codeStartField
            End Get
        End Property

        Protected Overrides ReadOnly Property CodeEnd As String
            Get
                Return codeEndField
            End Get
        End Property
    End Class
#End Region
End Namespace
