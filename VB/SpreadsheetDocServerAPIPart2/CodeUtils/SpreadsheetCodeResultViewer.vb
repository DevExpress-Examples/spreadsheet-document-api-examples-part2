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
    Partial Public Class SpreadsheetCSExampleCodeEvaluator
        Inherits SpreadsheetExampleCodeEvaluator

        Protected Overrides Function GetCodeDomProvider() As CodeDomProvider
            Return New CSharpCodeProvider()
        End Function


        Private Const codeStart_Renamed As String = "using System;" & ControlChars.CrLf & "using DevExpress.Spreadsheet;" & ControlChars.CrLf & "using DevExpress.Spreadsheet.Charts;" & ControlChars.CrLf & "using DevExpress.Spreadsheet.Drawings;" & ControlChars.CrLf & "using DevExpress.XtraSpreadsheet.Export;" & ControlChars.CrLf & "using System.Drawing;" & ControlChars.CrLf & "using System.Windows.Forms;" & ControlChars.CrLf & "using DevExpress.XtraPrinting;" & ControlChars.CrLf & "using DevExpress.XtraPrinting.Control;" & ControlChars.CrLf & "using DevExpress.Utils;" & ControlChars.CrLf & "using System.IO;" & ControlChars.CrLf & "using System.Diagnostics;" & ControlChars.CrLf & "using System.Xml;" & ControlChars.CrLf & "using System.Data;" & ControlChars.CrLf & "using System.Collections.Generic;" & ControlChars.CrLf & "using System.Globalization;" & ControlChars.CrLf & "using Formatting = DevExpress.Spreadsheet.Formatting;" & ControlChars.CrLf & "namespace SpreadsheetCodeResultViewer { " & ControlChars.CrLf & "public class ExampleItem { " & ControlChars.CrLf & "        public static void Process(Workbook workbook) { " & ControlChars.CrLf & ControlChars.CrLf


        Private Const codeEnd_Renamed As String = "       " & ControlChars.CrLf & " }" & ControlChars.CrLf & "    }" & ControlChars.CrLf & "}" & ControlChars.CrLf
        Protected Overrides ReadOnly Property CodeStart() As String
            Get
                Return codeStart_Renamed
            End Get
        End Property
        Protected Overrides ReadOnly Property CodeEnd() As String
            Get
                Return codeEnd_Renamed
            End Get
        End Property
    End Class
    #End Region
    #Region "SpreadsheetVbExampleCodeEvaluator"
    Partial Public Class SpreadsheetVbExampleCodeEvaluator
        Inherits SpreadsheetExampleCodeEvaluator

        Protected Overrides Function GetCodeDomProvider() As CodeDomProvider
            Return New Microsoft.VisualBasic.VBCodeProvider()
        End Function

        Private Const codeStart_Renamed As String = "Imports Microsoft.VisualBasic" & ControlChars.CrLf & "Imports System" & ControlChars.CrLf & "Imports DevExpress.Spreadsheet" & ControlChars.CrLf & "Imports DevExpress.Spreadsheet.Charts" & ControlChars.CrLf & "Imports DevExpress.Spreadsheet.Drawings" & ControlChars.CrLf & "Imports DevExpress.XtraSpreadsheet.Export" & ControlChars.CrLf & "Imports System.Drawing" & ControlChars.CrLf & "Imports System.Windows.Forms" & ControlChars.CrLf & "Imports DevExpress.XtraPrinting" & ControlChars.CrLf & "Imports DevExpress.XtraPrinting.Control" & ControlChars.CrLf & "Imports DevExpress.Utils" & ControlChars.CrLf & "Imports System.IO" & ControlChars.CrLf & "Imports System.Diagnostics" & ControlChars.CrLf & "Imports System.Xml" & ControlChars.CrLf & "Imports System.Data" & ControlChars.CrLf & "Imports System.Collections.Generic" & ControlChars.CrLf & "Imports System.Globalization" & ControlChars.CrLf & "Imports Formatting = DevExpress.Spreadsheet.Formatting" & ControlChars.CrLf & "Namespace SpreadsheetCodeResultViewer" & ControlChars.CrLf & "	Public Class ExampleItem" & ControlChars.CrLf & "		Public Shared Sub Process(ByVal workbook As Workbook)" & ControlChars.CrLf & ControlChars.CrLf


        Private Const codeEnd_Renamed As String = ControlChars.CrLf & "		End Sub" & ControlChars.CrLf & "	End Class" & ControlChars.CrLf & "End Namespace" & ControlChars.CrLf

        Protected Overrides ReadOnly Property CodeStart() As String
            Get
                Return codeStart_Renamed
            End Get
        End Property
        Protected Overrides ReadOnly Property CodeEnd() As String
            Get
                Return codeEnd_Renamed
            End Get
        End Property
    End Class
    #End Region
End Namespace
