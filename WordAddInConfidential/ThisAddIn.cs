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
            this.Application.DocumentOpen += Application_DocumentOpen;

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

            Microsoft.Office.Core.DocumentProperties prp = Doc.CustomDocumentProperties;

            bool isFileGuid = false;

            string fileGuid = string.Empty;

            foreach (Microsoft.Office.Core.DocumentProperty documentProperty in prp)
            {
                
                if (documentProperty.Name.Equals("FileGuid"))
                {
                    fileGuid = documentProperty.Value;
                    isFileGuid = true;
                }
            }

            //add Guid to the file
            if (isFileGuid == false)
            {
                fileGuid = System.Guid.NewGuid().ToString();
                prp.Add("FileGuid", false, Office.MsoDocProperties.msoPropertyTypeString, fileGuid, null);
            }

            //MessageBox.Show(fileGuid);

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
