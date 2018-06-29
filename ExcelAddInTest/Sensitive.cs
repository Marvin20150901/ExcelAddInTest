using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelAddInTest.Config;
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

        private void toggleButtonSecret_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Secret");
        }

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

        private void toggleButtonConfidential_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Confidential");

        }

        private void toggleButtonInternal_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Internal");
        }

        private void toggleButtonPublic_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Public");
        }

        private void toggleButtonMarkYes_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.Ribbons.Sensitive.toggleButtonMarkYes.Checked = true;
            Globals.Ribbons.Sensitive.toggleButtonMarkNo.Checked = false;

            var appConfig = Globals.ThisAddIn.AppConfig;
            var jsonFile = Globals.ThisAddIn.AppConfigFileInfo;

            if (appConfig!=null && jsonFile!=null)
            {
                appConfig.IsMark = true;
                var jsonSerialize = new Utilities.IO.Serializers.Default.JSONSerializer();
                var varCofig= jsonSerialize.Serialize(typeof(AddinConfig), appConfig);
                jsonFile.Write(varCofig);
            }

        }

        private void toggleButtonMarkNo_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.Ribbons.Sensitive.toggleButtonMarkYes.Checked = false;
            Globals.Ribbons.Sensitive.toggleButtonMarkNo.Checked = true;

            var appConfig = Globals.ThisAddIn.AppConfig;
            var jsonFile = Globals.ThisAddIn.AppConfigFileInfo;

            if (appConfig != null && jsonFile != null)
            {
                appConfig.IsMark = false;
                var jsonSerialize = new Utilities.IO.Serializers.Default.JSONSerializer();
                var varCofig = jsonSerialize.Serialize(typeof(AddinConfig), appConfig);
                jsonFile.Write(varCofig);
            }
        }
    }
}
