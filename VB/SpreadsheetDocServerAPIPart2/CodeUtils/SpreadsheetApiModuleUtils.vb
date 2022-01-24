Imports DevExpress.Utils
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.Internal
Imports System

Namespace SpreadsheetDocServerAPIPart2

    Public Class ExampleCodeEditor

        Private ReadOnly codeEditorCs As IRichEditControl

        Private ReadOnly codeEditorVb As IRichEditControl

        Private current As ExampleLanguage

        Private forceTextChangesCounter As Integer

        Private richEditTextChangedField As Boolean = False

        Private lastExampleCodeModifiedTimeField As Date = Date.Now

        Public Sub New(ByVal codeEditorCs As IRichEditControl, ByVal codeEditorVb As IRichEditControl)
            Me.codeEditorCs = codeEditorCs
            Me.codeEditorVb = codeEditorVb
            AddHandler Me.codeEditorCs.InnerControl.InitializeDocument, New EventHandler(AddressOf InitializeSyntaxHighlightForCs)
            AddHandler Me.codeEditorVb.InnerControl.InitializeDocument, New EventHandler(AddressOf InitializeSyntaxHighlightForVb)
        End Sub

        Public ReadOnly Property CurrentCodeEditor As InnerRichEditControl
            Get
                If CurrentExampleLanguage = ExampleLanguage.Csharp Then
                    Return codeEditorCs.InnerControl
                Else
                    Return codeEditorVb.InnerControl
                End If
            End Get
        End Property

        Public ReadOnly Property LastExampleCodeModifiedTime As Date
            Get
                Return lastExampleCodeModifiedTimeField
            End Get
        End Property

        Public ReadOnly Property RichEditTextChanged As Boolean
            Get
                Return richEditTextChangedField
            End Get
        End Property

        Public Property CurrentExampleLanguage As ExampleLanguage
            Get
                Return current
            End Get

            Set(ByVal value As ExampleLanguage)
                Try
                    UnsubscribeRichEditEvents()
                    current = value
                Finally
                    SubscribeRichEditEvent()
                    forceTextChangesCounter = 0 ' actualy where are no changes in that richEdit (CurrentCodeEditor)
                    richEditTextChangedField = True
                End Try
            End Set
        End Property

        Private Sub richEditControl_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
            If forceTextChangesCounter <= 0 Then
                richEditTextChangedField = True
                lastExampleCodeModifiedTimeField = Date.Now
            Else
                forceTextChangesCounter -= 1
            End If
        End Sub

        Public Function ShowExample(ByVal oldExample As CodeExample, ByVal newExample As CodeExample) As String
            Dim richEditControlCs As InnerRichEditControl = codeEditorCs.InnerControl
            Dim richEditControlVb As InnerRichEditControl = codeEditorVb.InnerControl
            If oldExample IsNot Nothing Then
                'save edited example
                oldExample.CodeCS = richEditControlCs.Text
                oldExample.CodeVB = richEditControlVb.Text
            End If

            Dim exampleCode As String = String.Empty
            If newExample IsNot Nothing Then
                Try
                    forceTextChangesCounter = 2
                    exampleCode = If(CurrentExampleLanguage = ExampleLanguage.Csharp, newExample.CodeCS, newExample.CodeVB)
                    richEditControlCs.Text = newExample.CodeCS
                    richEditControlVb.Text = newExample.CodeVB
                    richEditTextChangedField = False
                Finally
                    richEditTextChangedField = True
                End Try
            End If

            Return exampleCode
        End Function

        Private Sub UpdatePageBackground(ByVal codeEvaluated As Boolean)
            CurrentCodeEditor.Document.SetPageBackground(If(codeEvaluated, DXColor.Empty, DXColor.FromArgb(&HFF, &HBC, &HC8)), True)
        End Sub

        Friend Sub BeforeCompile()
            UnsubscribeRichEditEvents()
        End Sub

        Friend Sub AfterCompile(ByVal codeExcecutedWithoutExceptions As Boolean)
            UpdatePageBackground(codeExcecutedWithoutExceptions)
            richEditTextChangedField = False
            ResetLastExampleModifiedTime()
            SubscribeRichEditEvent()
        End Sub

        Public Sub ResetLastExampleModifiedTime()
            lastExampleCodeModifiedTimeField = Date.Now
        End Sub

        Private Sub UnsubscribeRichEditEvents()
            RemoveHandler CurrentCodeEditor.ContentChanged, AddressOf richEditControl_TextChanged
        End Sub

        Private Sub SubscribeRichEditEvent()
            AddHandler CurrentCodeEditor.ContentChanged, AddressOf richEditControl_TextChanged
        End Sub

        Private Sub InitializeSyntaxHighlightForCs(ByVal sender As Object, ByVal e As EventArgs)
            Dim syntaxHightlightInitializator As SyntaxHightlightInitializeHelper = New SyntaxHightlightInitializeHelper()
            syntaxHightlightInitializator.Initialize(codeEditorCs, GetCodeExampleFileExtension(ExampleLanguage.Csharp))
            DisableRichEditFeatures(codeEditorCs)
        End Sub

        Private Sub InitializeSyntaxHighlightForVb(ByVal sender As Object, ByVal e As EventArgs)
            Dim syntaxHightlightInitializator As SyntaxHightlightInitializeHelper = New SyntaxHightlightInitializeHelper()
            syntaxHightlightInitializator.Initialize(codeEditorVb, GetCodeExampleFileExtension(ExampleLanguage.VB))
            DisableRichEditFeatures(codeEditorVb)
        End Sub

        Private Sub DisableRichEditFeatures(ByVal codeEditor As IRichEditControl)
            Dim options As RichEditControlOptionsBase = codeEditor.InnerDocumentServer.Options
            options.DocumentCapabilities.Hyperlinks = DocumentCapability.Disabled
            options.DocumentCapabilities.Numbering.Bulleted = DocumentCapability.Disabled
            options.DocumentCapabilities.Numbering.Simple = DocumentCapability.Disabled
            options.DocumentCapabilities.Numbering.MultiLevel = DocumentCapability.Disabled
            options.DocumentCapabilities.Tables = DocumentCapability.Disabled
            'options.DocumentCapabilities.CharacterFormatting = DocumentCapability.Disabled;
            'options.DocumentCapabilities.ParagraphFormatting = DocumentCapability.Disabled;
            options.DocumentCapabilities.Bookmarks = DocumentCapability.Disabled
            options.DocumentCapabilities.CharacterStyle = DocumentCapability.Disabled
            options.DocumentCapabilities.ParagraphStyle = DocumentCapability.Disabled
        End Sub
    End Class
End Namespace
