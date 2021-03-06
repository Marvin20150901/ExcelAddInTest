﻿using System;
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
        private string appPath;
        private void Sensitive_Load(object sender, RibbonUIEventArgs e)
        {
            appPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\OfficeAddinConfidential\";

            if (!Directory.Exists(appPath))
            {
                Directory.CreateDirectory(appPath);
            }

            if (!File.Exists(appPath + @"Secret.pang"))
            {
                Properties.Resources.Secret.Save(appPath + "Secret.png");
                Properties.Resources.Internal.Save(appPath + "Internal.png");
                Properties.Resources.Confidential.Save(appPath + "Confidential.png");

                Properties.Settings.Default.IsImageUpdata = false;
                Properties.Settings.Default.Save();
            }
            
            //Globals.ThisAddIn.Application.UserLibraryPath

        }


        /// <summary>
        /// secret level
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleButtonSecret_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Secret");
//            var activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            foreach (Excel.Worksheet activeSheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets)
            {
                if (Properties.Settings.Default.IsMask)
                {
                    if (!File.Exists(appPath + "Secret.png"))
                    {
                        Properties.Resources.Secret.Save(appPath + "Secret.png");
                    }

                    activeSheet.PageSetup.LeftHeaderPicture.Filename = appPath + "Secret.png";
                    activeSheet.PageSetup.LeftHeader = "&G";
                }
                else
                {
                    var picname = activeSheet.PageSetup.LeftHeaderPicture.Filename;
                    if (picname.Equals(appPath + "Secret.png") || picname.Equals(appPath + "Confidential.png") || picname.Equals(appPath + "Internal.png"))
                    {
                        activeSheet.PageSetup.LeftHeaderPicture.Filename = "";
                        activeSheet.PageSetup.LeftHeader = "&G";
                    }
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
//            var activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            
            foreach (Excel.Worksheet activeSheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets)
            {
                if (Properties.Settings.Default.IsMask)
                {
                    if (!File.Exists(appPath + "Confidential.png"))
                    {
                        Properties.Resources.Confidential.Save(appPath + "Confidential.png");
                    }

                    activeSheet.PageSetup.LeftHeaderPicture.Filename = appPath + "Confidential.png";
                    activeSheet.PageSetup.LeftHeader = "&G";
                }
                else
                {
                    var picname = activeSheet.PageSetup.LeftHeaderPicture.Filename;
                    if (picname.Equals(appPath + "Secret.png") || picname.Equals(appPath + "Confidential.png") || picname.Equals(appPath + "Internal.png"))
                    {
                        activeSheet.PageSetup.LeftHeaderPicture.Filename = "";
                        activeSheet.PageSetup.LeftHeader = "&G";
                    }
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

            //var activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            foreach (Excel.Worksheet activeSheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets)
            {
                if (Properties.Settings.Default.IsMask)
                {
                    if (!File.Exists(appPath + "Internal.png"))
                    {
                        Properties.Resources.Confidential.Save(appPath + "Internal.png");
                    }

                    activeSheet.PageSetup.LeftHeaderPicture.Filename = appPath + "Internal.png";
                    activeSheet.PageSetup.LeftHeader = "&G";
                }
                else
                {
                    var picname = activeSheet.PageSetup.LeftHeaderPicture.Filename;
                    if (picname.Equals(appPath + "Secret.png") || picname.Equals(appPath + "Confidential.png") || picname.Equals(appPath + "Internal.png"))
                    {
                        activeSheet.PageSetup.LeftHeaderPicture.Filename = "";
                        activeSheet.PageSetup.LeftHeader = "&G";
                    }
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

            //var activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;


            foreach (Excel.Worksheet activeSheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets)
            {
                if (Properties.Settings.Default.IsMask)
                {

                    activeSheet.PageSetup.LeftHeaderPicture.Filename = "";
                    activeSheet.PageSetup.LeftHeader = "&G";
                }
                else
                {
                    var picname = activeSheet.PageSetup.LeftHeaderPicture.Filename;
                    if (picname.Equals(appPath + "Secret.png") || picname.Equals(appPath + "Confidential.png") || picname.Equals(appPath + "Internal.png"))
                    {
                        activeSheet.PageSetup.LeftHeaderPicture.Filename = "";
                        activeSheet.PageSetup.LeftHeader = "&G";
                    }
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

                foreach (Excel.Worksheet activeSheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets)
                {

                    var picname = activeSheet.PageSetup.LeftHeaderPicture.Filename;
                    if (picname.Equals(appPath + "Secret.png") || picname.Equals(appPath + "Confidential.png") ||
                        picname.Equals(appPath + "Internal.png"))
                    {
                        activeSheet.PageSetup.LeftHeaderPicture.Filename = "";
                        activeSheet.PageSetup.LeftHeader = "&G";
                    }


                }

            }
            catch (Exception)
            {

                throw;
            }
        }

        private void buttonClearTags_Click(object sender, RibbonControlEventArgs e)
        {
            foreach (Excel.Worksheet activeSheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets)
            {
                activeSheet.PageSetup.LeftHeaderPicture.Filename = "";
                activeSheet.PageSetup.LeftHeader = "&G";

            }


        }
    }
}
