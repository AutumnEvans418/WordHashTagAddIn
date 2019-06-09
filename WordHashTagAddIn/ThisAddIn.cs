using System;
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
    public partial class ThisAddIn
    {
        private IHashTagsViewModel vm;
        private string paneName = "HashTags";
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var container = new UnityContainer();
            container.RegisterSingleton<IHashTagsViewModel, HashTagsViewModel>();
            vm = container.Resolve<IHashTagsViewModel>();
            this.Application.DocumentBeforeSave +=
                new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);
            CustomTaskPane pane = CustomTaskPanes.FirstOrDefault(p => p.Title == paneName) ?? 
                                  this.CustomTaskPanes.Add(container.Resolve<HashTagsForm>(), paneName);
            Panes.HashTags = pane;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            Doc.Paragraphs[1].Range.InsertParagraphBefore();
            Doc.Paragraphs[1].Range.Text = "This text was added by using code.";
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
    }
}
