using DevExpress.Utils;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.Internal;
using System;

namespace SpreadsheetDocServerAPIPart2
{
    public class ExampleCodeEditor
    {
        readonly IRichEditControl codeEditorCs;
        readonly IRichEditControl codeEditorVb;
        ExampleLanguage current;

        int forceTextChangesCounter;
        bool richEditTextChanged = false;
        DateTime lastExampleCodeModifiedTime = DateTime.Now;

        public ExampleCodeEditor(IRichEditControl codeEditorCs, IRichEditControl codeEditorVb)
        {
            this.codeEditorCs = codeEditorCs;
            this.codeEditorVb = codeEditorVb;

            this.codeEditorCs.InnerControl.InitializeDocument += new System.EventHandler(this.InitializeSyntaxHighlightForCs);
            this.codeEditorVb.InnerControl.InitializeDocument += new System.EventHandler(this.InitializeSyntaxHighlightForVb);
        }
        public InnerRichEditControl CurrentCodeEditor
        {
            get
            {
                if (CurrentExampleLanguage == ExampleLanguage.Csharp)
                    return codeEditorCs.InnerControl;
                else
                    return codeEditorVb.InnerControl;
            }
        }
        public DateTime LastExampleCodeModifiedTime { get { return lastExampleCodeModifiedTime; } }

        public bool RichEditTextChanged { get { return richEditTextChanged; } }


        public ExampleLanguage CurrentExampleLanguage
        {
            get { return current; }
            set
            {
                try
                {
                    UnsubscribeRichEditEvents();
                    current = value;
                }
                finally
                {
                    SubscribeRichEditEvent();
                    forceTextChangesCounter = 0; // actualy where are no changes in that richEdit (CurrentCodeEditor)
                    richEditTextChanged = true;
                }
            }
        }
        void richEditControl_TextChanged(object sender, EventArgs e)
        {
            if (forceTextChangesCounter <= 0)
            {
                richEditTextChanged = true;
                lastExampleCodeModifiedTime = DateTime.Now;
            }
            else
                forceTextChangesCounter--;
        }

        public string ShowExample(CodeExample oldExample, CodeExample newExample)
        {
            InnerRichEditControl richEditControlCs = codeEditorCs.InnerControl;
            InnerRichEditControl richEditControlVb = codeEditorVb.InnerControl;

            if (oldExample != null)
            {
                //save edited example
                oldExample.CodeCS = richEditControlCs.Text;
                oldExample.CodeVB = richEditControlVb.Text;
            }
            string exampleCode = String.Empty;
            if (newExample != null)
            {
                try
                {
                    forceTextChangesCounter = 2;
                    exampleCode = (CurrentExampleLanguage == ExampleLanguage.Csharp) ? newExample.CodeCS : newExample.CodeVB;
                    richEditControlCs.Text = newExample.CodeCS;
                    richEditControlVb.Text = newExample.CodeVB;

                    richEditTextChanged = false;
                }
                finally
                {
                    richEditTextChanged = true;
                }
            }
            return exampleCode;
        }

        void UpdatePageBackground(bool codeEvaluated)
        {
            CurrentCodeEditor.Document.SetPageBackground((codeEvaluated) ? DXColor.Empty : DXColor.FromArgb(0xFF, 0xBC, 0xC8), true);
        }

        internal void BeforeCompile()
        {
            UnsubscribeRichEditEvents();
        }

        internal void AfterCompile(bool codeExcecutedWithoutExceptions)
        {
            UpdatePageBackground(codeExcecutedWithoutExceptions);

            richEditTextChanged = false;
            ResetLastExampleModifiedTime();

            SubscribeRichEditEvent();
        }
        public void ResetLastExampleModifiedTime()
        {
            lastExampleCodeModifiedTime = DateTime.Now;
        }
        private void UnsubscribeRichEditEvents()
        {
            CurrentCodeEditor.ContentChanged -= richEditControl_TextChanged;
        }
        void SubscribeRichEditEvent()
        {
            CurrentCodeEditor.ContentChanged += richEditControl_TextChanged;
        }
        void InitializeSyntaxHighlightForCs(object sender, EventArgs e)
        {
            SyntaxHightlightInitializeHelper syntaxHightlightInitializator = new SyntaxHightlightInitializeHelper();
            syntaxHightlightInitializator.Initialize(codeEditorCs, CodeExampleDemoUtils.GetCodeExampleFileExtension(ExampleLanguage.Csharp));

            DisableRichEditFeatures(codeEditorCs);
        }


        void InitializeSyntaxHighlightForVb(object sender, EventArgs e)
        {
            SyntaxHightlightInitializeHelper syntaxHightlightInitializator = new SyntaxHightlightInitializeHelper();
            syntaxHightlightInitializator.Initialize(codeEditorVb, CodeExampleDemoUtils.GetCodeExampleFileExtension(ExampleLanguage.VB));

            DisableRichEditFeatures(codeEditorVb);
        }
        void DisableRichEditFeatures(IRichEditControl codeEditor)
        {
            RichEditControlOptionsBase options = codeEditor.InnerDocumentServer.Options;
            options.DocumentCapabilities.Hyperlinks = DocumentCapability.Disabled;
            options.DocumentCapabilities.Numbering.Bulleted = DocumentCapability.Disabled;
            options.DocumentCapabilities.Numbering.Simple = DocumentCapability.Disabled;
            options.DocumentCapabilities.Numbering.MultiLevel = DocumentCapability.Disabled;

            options.DocumentCapabilities.Tables = DocumentCapability.Disabled;
            //options.DocumentCapabilities.CharacterFormatting = DocumentCapability.Disabled;
            //options.DocumentCapabilities.ParagraphFormatting = DocumentCapability.Disabled;
            options.DocumentCapabilities.Bookmarks = DocumentCapability.Disabled;

            options.DocumentCapabilities.CharacterStyle = DocumentCapability.Disabled;
            options.DocumentCapabilities.ParagraphStyle = DocumentCapability.Disabled;
        }
    }
}
