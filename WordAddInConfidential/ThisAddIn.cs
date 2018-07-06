using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
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
            //this.Application.DocumentOpen += Application_DocumentOpen;

            //this.Application.DocumentChange;
            this.Application.WindowActivate += Application_WindowActivate;

            if (Properties.Settings.Default.IsMask)
            {
                Globals.Ribbons.Sensitive.toggleButtonMarkYes.Checked = true;
                Globals.Ribbons.Sensitive.toggleButtonMarkNo.Checked = false;
            }
            else
            {
                Globals.Ribbons.Sensitive.toggleButtonMarkYes.Checked = false;
                Globals.Ribbons.Sensitive.toggleButtonMarkNo.Checked = true;
            }

            

        }


        private void Application_WindowActivate(Word.Document Doc, Word.Window Wn)
        {
            //throw new NotImplementedException();
            try
            {
                Office.DocumentProperties prp = this.Application.ActiveDocument.CustomDocumentProperties;

                bool isSenitive = false;

                foreach (Office.DocumentProperty documentProperty in prp)
                {
                    if (documentProperty.Name.Equals("Sensitive"))
                    {
                        InitRabbionControl(documentProperty.Value);
                        isSenitive = true;
                    }
                }

                if (isSenitive == false)
                {
                    InitRabbionControl(string.Empty);
                }
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.ToString());
                throw;
            }
        }

        private void Application_DocumentOpen(Word.Document Doc)
        {
            //throw new NotImplementedException();

            

            /*
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

                //Word.Shape varShape=section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.AddPicture("Secret.png",ref lt,ref sw,ref lf,ref tp,ref wd,ref hg);
                Word.Shape varShape = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes
                    .AddPicture("Secret.png");
                
                varShape.RelativeHorizontalPosition =
                    Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                varShape.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                varShape.TopRelative = 0;
                varShape.LeftRelative = 0;

            }
            */
            
        }



        /// <summary>
        /// set rabbion controls state corrding the sensitive properes
        /// </summary>
        /// <param name="sensitive"></param>
        public void InitRabbionControl(string sensitive)
        {
            try
            {
                if (sensitive != string.Empty)
                {
                    if (sensitive.Equals("Secret"))
                    {
                        Globals.Ribbons.Sensitive.toggleButtonSecret.Checked = true;
                        Globals.Ribbons.Sensitive.toggleButtonConfidential.Checked = false;
                        Globals.Ribbons.Sensitive.toggleButtonInternal.Checked = false;
                        Globals.Ribbons.Sensitive.toggleButtonPublic.Checked = false;
                    }
                    else if (sensitive.Equals("Confidential"))
                    {
                        Globals.Ribbons.Sensitive.toggleButtonSecret.Checked = false;
                        Globals.Ribbons.Sensitive.toggleButtonConfidential.Checked = true;
                        Globals.Ribbons.Sensitive.toggleButtonInternal.Checked = false;
                        Globals.Ribbons.Sensitive.toggleButtonPublic.Checked = false;
                    }
                    else if (sensitive.Equals("Internal"))
                    {
                        Globals.Ribbons.Sensitive.toggleButtonSecret.Checked = false;
                        Globals.Ribbons.Sensitive.toggleButtonConfidential.Checked = false;
                        Globals.Ribbons.Sensitive.toggleButtonInternal.Checked = true;
                        Globals.Ribbons.Sensitive.toggleButtonPublic.Checked = false;
                    }
                    else if (sensitive.Equals("Public"))
                    {
                        Globals.Ribbons.Sensitive.toggleButtonSecret.Checked = false;
                        Globals.Ribbons.Sensitive.toggleButtonConfidential.Checked = false;
                        Globals.Ribbons.Sensitive.toggleButtonInternal.Checked = false;
                        Globals.Ribbons.Sensitive.toggleButtonPublic.Checked = true;
                    }
                    else
                    {
                        Globals.Ribbons.Sensitive.toggleButtonSecret.Checked = false;
                        Globals.Ribbons.Sensitive.toggleButtonConfidential.Checked = false;
                        Globals.Ribbons.Sensitive.toggleButtonInternal.Checked = false;
                        Globals.Ribbons.Sensitive.toggleButtonPublic.Checked = false;
                    }

                }
                else
                {
                    Globals.Ribbons.Sensitive.toggleButtonSecret.Checked = false;
                    Globals.Ribbons.Sensitive.toggleButtonConfidential.Checked = false;
                    Globals.Ribbons.Sensitive.toggleButtonInternal.Checked = false;
                    Globals.Ribbons.Sensitive.toggleButtonPublic.Checked = false;
                }
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.ToString());
                throw;
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
