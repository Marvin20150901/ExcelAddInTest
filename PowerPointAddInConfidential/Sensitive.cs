using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;

namespace PowerPointAddInConfidential
{
    public partial class Sensitive
    {
        private void Sensitive_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void toggleButtonSecret_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Secret");
            //            var activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            var activePpt = Globals.ThisAddIn.Application.ActivePresentation;
            var size = Properties.Resources.Secret.Size;
            if (Properties.Settings.Default.IsMask)
            {
                string logo = "Secret.png";
                var shap=activePpt.SlideMaster.Shapes.AddPicture(logo, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, size.Width,size.Height);
                shap.Line.Visible=MsoTriState.msoFalse;
            }
            else
            {

            }


        }

        private void toggleButtonConfidential_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Confidential");

            var activePpt = Globals.ThisAddIn.Application.ActivePresentation;

            if (Properties.Settings.Default.IsMask)
            {

                //add  logic
            }
            else
            {

            }

        }

        private void toggleButtonInternal_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Internal");

            var activePpt = Globals.ThisAddIn.Application.ActivePresentation;

            if (Properties.Settings.Default.IsMask)
            {

                //add  logic
            }
            else
            {

            }
        }

        private void toggleButtonPublic_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Public");

            var activePpt = Globals.ThisAddIn.Application.ActivePresentation;

            if (Properties.Settings.Default.IsMask)
            {

                //add  logic
            }
            else
            {

            }
        }

        private void toggleButtonMarkYes_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void toggleButtonMarkNo_Click(object sender, RibbonControlEventArgs e)
        {

        }




        /// <summary>
        /// set the confidential level
        /// </summary>
        /// <param name="sensitive"> confidential level string</param>
        private void SetWorkbookSensitive(string sensitive)
        {

            try
            {

                Microsoft.Office.Core.DocumentProperties prp = Globals.ThisAddIn.Application.ActivePresentation.CustomDocumentProperties;

                bool isSenitive = false;

                foreach (Microsoft.Office.Core.DocumentProperty documentProperty in prp)
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
                    prp.Add("Sensitive", false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, sensitive, null);
                }

                Globals.ThisAddIn.InitRabbionControl(sensitive);

            }
            catch (Exception)
            {

                throw;
            }
        }

    }
}
