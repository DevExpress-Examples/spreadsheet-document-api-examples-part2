Imports DevExpress.Spreadsheet
Imports DevExpress.XtraTab
Imports DevExpress.XtraTreeList
Imports DevExpress.XtraTreeList.Columns
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms

Namespace SpreadsheetDocServerAPIPart2

    Public Partial Class Form1
        Inherits Form

        Private workbook As Workbook = New Workbook()

        Private defaultCulture As CultureInfo = New CultureInfo("en-US")

        Private codeEditor As ExampleCodeEditor

        Private evaluator As ExampleEvaluatorByTimer

        Private examples As List(Of CodeExampleGroup)

        Private treeListRootNodeLoading As Boolean = True

        Public Sub New()
            InitializeComponent()
            Dim examplePath As String = GetExamplePath("CodeExamples")
            Dim examplesCS As Dictionary(Of String, FileInfo) = GatherExamplesFromProject(examplePath, ExampleLanguage.Csharp)
            Dim examplesVB As Dictionary(Of String, FileInfo) = GatherExamplesFromProject(examplePath, ExampleLanguage.VB)
            DisableTabs(examplesCS.Count, examplesVB.Count)
            examples = FindExamples(examplePath, examplesCS, examplesVB)
            MergeGroups()
            ShowExamplesInTreeList(treeList1, examples)
            codeEditor = New ExampleCodeEditor(richEditControlCS, richEditControlVB)
            CurrentExampleLanguage = DetectExampleLanguage("SpreadsheetDocServerAPIPart2")
            evaluator = New SpreadsheetExampleEvaluatorByTimer()
            AddHandler evaluator.QueryEvaluate, AddressOf OnExampleEvaluatorQueryEvaluate
            AddHandler evaluator.OnBeforeCompile, AddressOf evaluator_OnBeforeCompile
            AddHandler evaluator.OnAfterCompile, AddressOf evaluator_OnAfterCompile
            ShowFirstExample()
            AddHandler xtraTabControl1.SelectedPageChanged, New TabPageChangedEventHandler(AddressOf xtraTabControl1_SelectedPageChanged)
        End Sub

        Private Sub MergeGroups()
            Dim uniqueNameGroup = New Dictionary(Of String, CodeExampleGroup)()
            For Each n As CodeExampleGroup In examples
                If uniqueNameGroup.ContainsKey(n.Name) Then
                    uniqueNameGroup(n.Name).Merge(n)
                Else
                    uniqueNameGroup(n.Name) = n
                End If
            Next

            examples.Clear()
            For Each value In uniqueNameGroup.Values
                examples.Add(value)
            Next
        End Sub

        Private Sub evaluator_OnAfterCompile(ByVal sender As Object, ByVal args As OnAfterCompileEventArgs)
            codeEditor.AfterCompile(args.Result)
            workbook.Worksheets.ActiveWorksheet.Visible = True
            workbook.EndUpdate()
        End Sub

        Private Sub evaluator_OnBeforeCompile(ByVal sender As Object, ByVal e As EventArgs)
            workbook.BeginUpdate()
            codeEditor.BeforeCompile()
            workbook.Options.Culture = defaultCulture
            Dim loaded As Boolean = workbook.LoadDocument("Document.xlsx")
            Call Debug.Assert(loaded)
        End Sub

        Private Property CurrentExampleLanguage As ExampleLanguage
            Get
                Return CType(xtraTabControl1.SelectedTabPageIndex, ExampleLanguage)
            End Get

            Set(ByVal value As ExampleLanguage)
                codeEditor.CurrentExampleLanguage = value
                xtraTabControl1.SelectedTabPageIndex = If(value = ExampleLanguage.Csharp, 0, 1)
            End Set
        End Property

        Private Sub ShowExamplesInTreeList(ByVal treeList As TreeList, ByVal examples As List(Of CodeExampleGroup))
#Region "InitializeTreeList"
            treeList.OptionsPrint.UsePrintStyles = True
            AddHandler treeList.FocusedNodeChanged, New FocusedNodeChangedEventHandler(AddressOf OnNewExampleSelected)
            treeList.OptionsView.ShowColumns = False
            treeList.OptionsView.ShowIndicator = False
            AddHandler treeList.VirtualTreeGetChildNodes, AddressOf treeList_VirtualTreeGetChildNodes
            AddHandler treeList.VirtualTreeGetCellValue, AddressOf treeList_VirtualTreeGetCellValue
#End Region
            Dim col1 As TreeListColumn = New TreeListColumn()
            col1.VisibleIndex = 0
            col1.OptionsColumn.AllowEdit = False
            col1.OptionsColumn.AllowMove = False
            col1.OptionsColumn.ReadOnly = True
            treeList.Columns.AddRange(New TreeListColumn() {col1})
            treeList.DataSource = New [Object]()
            treeList.ExpandAll()
        End Sub

        Private Sub treeList_VirtualTreeGetCellValue(ByVal sender As Object, ByVal args As VirtualTreeGetCellValueInfo)
            Dim group As CodeExampleGroup = TryCast(args.Node, CodeExampleGroup)
            If group IsNot Nothing Then args.CellData = group.Name
            Dim example As CodeExample = TryCast(args.Node, CodeExample)
            If example IsNot Nothing Then args.CellData = example.RegionName
        End Sub

        Private Sub treeList_VirtualTreeGetChildNodes(ByVal sender As Object, ByVal args As VirtualTreeGetChildNodesInfo)
            If treeListRootNodeLoading Then
                args.Children = examples
                treeListRootNodeLoading = False
            Else
                If args.Node Is Nothing Then Return
                Dim group As CodeExampleGroup = TryCast(args.Node, CodeExampleGroup)
                If group IsNot Nothing Then args.Children = group.Examples
            End If
        End Sub

        Private Sub ShowFirstExample()
            treeList1.ExpandAll()
            If treeList1.Nodes.Count > 0 Then treeList1.FocusedNode = treeList1.MoveFirst().FirstNode
        End Sub

        Private Sub OnNewExampleSelected(ByVal sender As Object, ByVal e As FocusedNodeChangedEventArgs)
            Dim newExample As CodeExample = TryCast(TryCast(sender, TreeList).GetDataRecordByNode(e.Node), CodeExample)
            Dim oldExample As CodeExample = TryCast(TryCast(sender, TreeList).GetDataRecordByNode(e.OldNode), CodeExample)
            If newExample Is Nothing Then Return
            Dim exampleCode As String = codeEditor.ShowExample(oldExample, newExample)
            codeExampleNameLbl.Text = ConvertStringToMoreHumanReadableForm(newExample.RegionName)
            Dim args As CodeEvaluationEventArgs = New CodeEvaluationEventArgs()
            InitializeCodeEvaluationEventArgs(args, newExample.RegionName)
            evaluator.ForceCompile(args)
        End Sub

        Private Sub InitializeCodeEvaluationEventArgs(ByVal e As CodeEvaluationEventArgs, ByVal regionName As String)
            e.Result = True
            e.Code = codeEditor.CurrentCodeEditor.Text
            e.Language = CurrentExampleLanguage
            e.EvaluationParameter = workbook
            e.RegionName = regionName
        End Sub

        Private Sub OnExampleEvaluatorQueryEvaluate(ByVal sender As Object, ByVal e As CodeEvaluationEventArgs)
            e.Result = False
            If codeEditor.RichEditTextChanged Then ' && compileComplete) {
                Dim span As TimeSpan = Date.Now - codeEditor.LastExampleCodeModifiedTime
                If span < TimeSpan.FromMilliseconds(1000) Then 'CompileTimeIntervalInMilliseconds  1900
                    codeEditor.ResetLastExampleModifiedTime()
                    Return
                End If

                'e.Result = true;
                InitializeCodeEvaluationEventArgs(e, e.RegionName)
            End If
        End Sub

        Private Sub xtraTabControl1_SelectedPageChanged(ByVal sender As Object, ByVal e As TabPageChangedEventArgs)
            Dim value As ExampleLanguage = CType(xtraTabControl1.SelectedTabPageIndex, ExampleLanguage)
            If codeEditor IsNot Nothing Then codeEditor.CurrentExampleLanguage = value
        End Sub

        Private Sub SpreadsheetAPIModule_Disposed(ByVal sender As Object, ByVal e As EventArgs)
            evaluator.Dispose()
        End Sub

        Private Sub DisableTabs(ByVal examplesCSCount As Integer, ByVal examplesVBCount As Integer)
            If examplesCSCount = 0 Then xtraTabControl1.TabPages(CInt(ExampleLanguage.Csharp)).PageEnabled = False
            If examplesVBCount = 0 Then xtraTabControl1.TabPages(CInt(ExampleLanguage.VB)).PageEnabled = False
        End Sub

        Private Sub btnOpenExcel_Click(ByVal sender As Object, ByVal e As EventArgs)
            Dim fileName As String = "SampleDocument.xlsx"
            workbook.SaveDocument(fileName, DocumentFormat.Xlsx)
            Call Process.Start(fileName)
        End Sub
    End Class
End Namespace
