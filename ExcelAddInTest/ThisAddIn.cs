using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Utilities.DataTypes;
using Utilities.IO;

namespace ExcelAddInTest
{
    public partial class ThisAddIn
    {

//        public AddinConfig AppConfig;
//        public FileInfo AppConfigFileInfo;
        public string IsMask;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

           
            try
            {
                //get the appsetting config form app.config file

                var t = ConfigurationManager.AppSettings["IsMask"];
                IsMask = t;

                //inite this rabbion

                if (IsMask.Equals("true"))
                {
                    Globals.Ribbons.Sensitive.toggleButtonMarkYes.Checked = true;
                    Globals.Ribbons.Sensitive.toggleButtonMarkNo.Checked = false;
                }
                else if (IsMask.Equals("false"))
                {
                    Globals.Ribbons.Sensitive.toggleButtonMarkYes.Checked = false;
                    Globals.Ribbons.Sensitive.toggleButtonMarkNo.Checked = true;
                }
                else
                {
                    
                }
                                   
            }
            catch (Exception)
            {
                
                throw;
            }
                      
        }



        /// <summary>
        /// Workbook changed or opened ,this function will be process
        /// Get the workbook's meat data to init the ribbons
        /// </summary>
        /// <param name="Wb"></param>
        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {

            try
            {
                Office.DocumentProperties prp = this.Application.ActiveWorkbook.CustomDocumentProperties;

                bool isSenitive = false;

                foreach (Office.DocumentProperty documentProperty in prp)
                {
                    if (documentProperty.Name.Equals("Sensitive"))
                    {
                        MessageBox.Show(documentProperty.Value);
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
                MessageBox.Show(e.ToString());                
                throw;
            }
            
        }



        /// <summary>
        /// set rabbion controls state corrding the sensitive properes
        /// </summary>
        /// <param name="sensitive"></param>
        public void InitRabbionControl(string sensitive)
        {
            try
            {
                if (sensitive!=string.Empty)
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
                MessageBox.Show(e.ToString());
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
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
            
        }



        #endregion
    }
}
