using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms.VisualStyles;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace WordAddInConfidential
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //this.Application.ActiveDocument.Sections
            this.Application.DocumentOpen += Application_DocumentOpen;
        }

        private void Application_DocumentOpen(Word.Document Doc)
        {
            //throw new NotImplementedException();

            foreach (Word.Section section in this.Application.ActiveDocument.Sections)
            {

                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                //headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                //headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

                //section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range
                if (!File.Exists("Secret.png"))
                {
                    Properties.Resources.Secret.Save("Secret.png");
                }


                int left = 0;
                int top = 0;
                int height = 40;
                int width = 180;

                object lt = false;
                object sw = true;
                object lf = (object) left;
                object tp = (object) top;
                object wd = (object) width;
                object hg = (object) height;

                
                //Word.InlineShape varShape= headerRange.InlineShapes.AddPicture("Secret.png");

                section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.AddPicture("Secret.png",ref lt,ref sw,ref lf,ref tp,ref wd,ref hg);

            }

            foreach (Word.Section wordSection in this.Application.ActiveDocument.Sections)
            {
                Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = Word.WdColorIndex.wdDarkRed;
                footerRange.Font.Size = 20;
                footerRange.Text = "Confidential";
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
