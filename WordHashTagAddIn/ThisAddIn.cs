using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Unity;

namespace WordHashTagAddIn
{
    public partial class ThisAddIn : IAddIn
    {
        private IHashTagsViewModel vm;
        private string paneName = "HashTags";
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var container = new UnityContainer();
            container.RegisterSingleton<IHashTagsViewModel, HashTagsViewModel>();
            container.RegisterInstance<IAddIn>(this);
            vm = container.Resolve<IHashTagsViewModel>();
            this.Application.DocumentBeforeSave +=
                Application_DocumentBeforeSave;
            CustomTaskPane pane = CustomTaskPanes.FirstOrDefault(p => p.Title == paneName) ??
                                  this.CustomTaskPanes.Add(container.Resolve<HashTagsForm>(), paneName);
            //pane.Visible = true;
            Panes.HashTags = pane;

        }

       


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            //Doc.Paragraphs[1].Range.InsertParagraphBefore();
            //Doc.Paragraphs[1].Range.Text = "This text was added by using code.";
            UpdateTags(Doc);
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        public void UpdateTags()
        {
            UpdateTags(this.Application.ActiveDocument);
        }

        public void NavigateToParagraph(HashTagItem selectedParagraph)
        {
            foreach (Word.Paragraph activeDocumentParagraph in this.Application.ActiveDocument.Paragraphs)
            {
                if (activeDocumentParagraph.Range.Text == selectedParagraph.Paragraph)
                {
                    var start = activeDocumentParagraph.Range.Start;
                    var end = start;
                    var range = this.Application.ActiveDocument.Range(start, end);
                    range.Select();
                }
            }
        }

        public void UpdateTags(Word.Document doc)
        {
            
            if (vm.IsShowingHashTagsView)
            {
                
                vm.ClearTags();
                foreach (Word.Paragraph docParagraph in doc.Paragraphs)
                {
                    var text = docParagraph.Range.Text;
                    var hashTags = text.Split(' ', '\r', '\n').Where(p => p.StartsWith("#"));
                    foreach (var hashTag in hashTags)
                    {
                        if (vm.IsHighlightingTags)
                        {
                            var start = docParagraph.Range.Start +
                                        text.IndexOf(hashTag, StringComparison.InvariantCultureIgnoreCase);
                            var end = start + hashTag.Length;
                            var range = doc.Range(start, end);
                            range.HighlightColorIndex = Word.WdColorIndex.wdYellow;
                        }
                        vm.AddTag(new HashTagItem() { Name = hashTag, Paragraph = text, });

                    }
                }
            }
        }
    }
}
