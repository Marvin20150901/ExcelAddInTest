using System;
using System.Collections.Generic;
using System.Configuration;
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
            var activeSheet = (Excel.Worksheet) Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            //activeSheet.PageSetup.LeftHeaderPicture.Filename


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

        }


        /// <summary>
        /// Internal level
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleButtonInternal_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Internal");
        }

        /// <summary>
        /// Public level
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleButtonPublic_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Public");
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

            try
            {
                Globals.ThisAddIn.IsMask = "true";
                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                config.AppSettings.Settings["isMask"].Value = "true";
                config.Save();
                }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        /// <summary>
        /// Yes button to mark
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleButtonMarkNo_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.Ribbons.Sensitive.toggleButtonMarkYes.Checked = false;
            Globals.Ribbons.Sensitive.toggleButtonMarkNo.Checked = true;

            try
            {
                Globals.ThisAddIn.IsMask = "false";
                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                config.AppSettings.Settings["isMask"].Value = "false";
                config.Save();
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
