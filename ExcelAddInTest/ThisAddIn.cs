using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddInTest
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
           // this.Application.WorkbookOpen += Application_WorkbookOpen;
            //(this.Application as Excel.AppEvents_Event).NewWorkbook += ThisAddIn_NewWorkbook;
            this.Application.WorkbookActivate += Application_WorkbookActivate;

        }


        /// <summary>
        /// Workbook changed or opened ,this function will be process
        /// Get the workbook's meat data to init the ribbons
        /// </summary>
        /// <param name="Wb"></param>
        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            //MessageBox.Show(Wb.Path);
            //throw new NotImplementedException();
            //Globals.Ribbons.Sensitive.toggleButtonSecret.Checked = true;

            Office.DocumentProperties prp= this.Application.ActiveWorkbook.CustomDocumentProperties;

            bool isSenitive = false;

            foreach (Office.DocumentProperty documentProperty in prp)
            {
                if (documentProperty.Name.Equals("Sensitive"))
                {
                    MessageBox.Show(documentProperty.Value);
                    isSenitive = true;
                }
                
            }

            if (isSenitive==false)
            {
                prp.Add("Sensitive", false, Office.MsoDocProperties.msoPropertyTypeString, "Secret", null);
            }
            

        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void ThisAddin_NewWorkbook(object sender, EventArgs e)
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
