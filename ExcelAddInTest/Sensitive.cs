using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;


namespace ExcelAddInTest
{
    public partial class Sensitive
    {
        private Excel.Application App;

        private void Sensitive_Load(object sender, RibbonUIEventArgs e)
        {
            App = Globals.ThisAddIn.Application;
            
        }


        /// <summary>
        /// secret level
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleButtonSecret_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Secret");
            var activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            if (Properties.Settings.Default.IsMask)
            {
                
                activeSheet.PageSetup.LeftHeaderPicture.Filename = "Secret.png";
                activeSheet.PageSetup.LeftHeader = "&G";
            }
            else
            {
                var picname = activeSheet.PageSetup.LeftHeaderPicture.Filename;
                if (picname.Equals("Secret.png") || picname.Equals("Confidential.png") || picname.Equals("Internal.png"))
                {
                    activeSheet.PageSetup.LeftHeaderPicture.Filename = "";
                    activeSheet.PageSetup.LeftHeader = "&G";
                }
            }
            
        }

        /// <summary>
        /// set the confidential level
        /// </summary>
        /// <param name="sensitive"> confidential level string</param>
        private void SetWorkbookSensitive(string sensitive)
        {

            try
            {

                Office.DocumentProperties prp = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties;

                bool isSenitive = false;

                foreach (Office.DocumentProperty documentProperty in prp)
                {
                    if (documentProperty.Name.Equals("Sensitive"))
                    {
                        //MessageBox.Show(documentProperty.Value);
                        //InitRabbionControl(documentProperty.Value);
                        documentProperty.Value = sensitive;
                        isSenitive = true;
                    }
                }

                if (isSenitive == false)
                {
                    prp.Add("Sensitive", false, Office.MsoDocProperties.msoPropertyTypeString, sensitive, null);
                }

                Globals.ThisAddIn.InitRabbionControl(sensitive);

            }
            catch (Exception)
            {

                throw;
            }
        }


        /// <summary>
        /// confidential level
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleButtonConfidential_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Confidential");
            var activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            if (Properties.Settings.Default.IsMask)
            {                
                activeSheet.PageSetup.LeftHeaderPicture.Filename = "Confidential.png";
                activeSheet.PageSetup.LeftHeader = "&G";
            }
            else
            {
                var picname = activeSheet.PageSetup.LeftHeaderPicture.Filename;
                if (picname.Equals("Secret.png") || picname.Equals("Confidential.png") || picname.Equals("Internal.png"))
                {
                    activeSheet.PageSetup.LeftHeaderPicture.Filename = "";
                    activeSheet.PageSetup.LeftHeader = "&G";
                }
            }

        }


        /// <summary>
        /// Internal level
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleButtonInternal_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Internal");

            var activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            if (Properties.Settings.Default.IsMask)
            {
               
                activeSheet.PageSetup.LeftHeaderPicture.Filename = "Internal.png";
                activeSheet.PageSetup.LeftHeader = "&G";
            }
            else
            {
                var picname = activeSheet.PageSetup.LeftHeaderPicture.Filename;
                if (picname.Equals("Secret.png") || picname.Equals("Confidential.png") || picname.Equals("Internal.png"))
                {
                    activeSheet.PageSetup.LeftHeaderPicture.Filename = "";
                    activeSheet.PageSetup.LeftHeader = "&G";
                }
            }


        }

        /// <summary>
        /// Public level
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleButtonPublic_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Public");

            var activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            if (Properties.Settings.Default.IsMask)
            {
                
                activeSheet.PageSetup.LeftHeaderPicture.Filename = "";
                activeSheet.PageSetup.LeftHeader = "&G";
            }
            else
            {
                var picname = activeSheet.PageSetup.LeftHeaderPicture.Filename;
                if (picname.Equals("Secret.png") || picname.Equals("Confidential.png") || picname.Equals("Internal.png"))
                {
                    activeSheet.PageSetup.LeftHeaderPicture.Filename = "";
                    activeSheet.PageSetup.LeftHeader = "&G";
                }
            }

            //activeSheet.PageSetup.LeftHeader = "&G";
        }

        /// <summary>
        /// Yes button to mark 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleButtonMarkYes_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.Ribbons.Sensitive.toggleButtonMarkYes.Checked = true;
            Globals.Ribbons.Sensitive.toggleButtonMarkNo.Checked = false;

            Properties.Settings.Default.IsMask = true;
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// NO button to mark
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleButtonMarkNo_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.Ribbons.Sensitive.toggleButtonMarkYes.Checked = false;
            Globals.Ribbons.Sensitive.toggleButtonMarkNo.Checked = true;

            try
            {
                Properties.Settings.Default.IsMask = false;
                Properties.Settings.Default.Save();
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
