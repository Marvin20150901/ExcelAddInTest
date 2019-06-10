using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointAddInConfidential
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

            if (!File.Exists(appPath + @"Secret.png"))
            {
                Properties.Resources.Secret.Save(appPath + @"Secret.png");
                Properties.Resources.Internal.Save(appPath + "Internal.png");
                Properties.Resources.Confidential.Save(appPath + "Confidential.png");

                Properties.Settings.Default.IsImageUpdata = false;
                Properties.Settings.Default.Save();

            }
            
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleButtonSecret_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Secret");
            //            var activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            var activePpt = Globals.ThisAddIn.Application.ActivePresentation;
            //var size = Properties.Resources.Secret.Size;

            
            if (Properties.Settings.Default.IsMask)
            {
                foreach (Shape slideMasterShape in activePpt.SlideMaster.Shapes)
                {
                    if (slideMasterShape.Name.Equals(Properties.Settings.Default.ShapName))
                    {
                        slideMasterShape.Delete();
                    }
                }


                if (!File.Exists(appPath + "Secret.png"))
                {
                    Properties.Resources.Secret.Save(appPath + "Secret.png");
                }

                string logo = appPath + @"Secret.png";
                var shap=activePpt.SlideMaster.Shapes.AddPicture(logo, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0);
                shap.Line.Visible=MsoTriState.msoFalse;
                shap.Name = Properties.Settings.Default.ShapName;
            }
            else
            {

                 //activePpt.SlideMaster.Shapes
                foreach (Shape slideMasterShape in activePpt.SlideMaster.Shapes)
                {
                    if (slideMasterShape.Name.Equals(Properties.Settings.Default.ShapName))
                    {
                       // MessageBox.Show("heheheeh");
                        slideMasterShape.Delete();
                    }
                    //slideMasterShape.
                }
            }


        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleButtonConfidential_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Confidential");

            var activePpt = Globals.ThisAddIn.Application.ActivePresentation;
            var size = Properties.Resources.Confidential.Size;
            if (Properties.Settings.Default.IsMask)
            {
                foreach (Shape slideMasterShape in activePpt.SlideMaster.Shapes)
                {
                    if (slideMasterShape.Name.Equals(Properties.Settings.Default.ShapName))
                    {
                        slideMasterShape.Delete();
                    }
                }
                
                if (!File.Exists(appPath + "Confidential.png"))
                {
                    Properties.Resources.Confidential.Save(appPath + "Confidential.png");
                }

                string logo = appPath + @"Confidential.png";
                var shap = activePpt.SlideMaster.Shapes.AddPicture(logo, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0);
                shap.Line.Visible = MsoTriState.msoFalse;
                shap.Name = Properties.Settings.Default.ShapName;
                //add  logic
            }
            else
            {
                foreach (Shape slideMasterShape in activePpt.SlideMaster.Shapes)
                {
                    if (slideMasterShape.Name.Equals(Properties.Settings.Default.ShapName))
                    {
                        // MessageBox.Show("heheheeh");
                        slideMasterShape.Delete();
                    }
                    //slideMasterShape.
                }
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleButtonInternal_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Internal");
            var size = Properties.Resources.Internal.Size;
            var activePpt = Globals.ThisAddIn.Application.ActivePresentation;

            if (Properties.Settings.Default.IsMask)
            {
                foreach (Shape slideMasterShape in activePpt.SlideMaster.Shapes)
                {
                    if (slideMasterShape.Name.Equals(Properties.Settings.Default.ShapName))
                    {
                        slideMasterShape.Delete();
                    }
                }

                
                if (!File.Exists(appPath + "Internal.png"))
                {
                    Properties.Resources.Internal.Save(appPath + "Internal.png");
                }

                string logo = appPath + @"Internal.png";
                var shap = activePpt.SlideMaster.Shapes.AddPicture(logo, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0);
                shap.Line.Visible = MsoTriState.msoFalse;
                shap.Name = Properties.Settings.Default.ShapName;
                //add  logic
            }
            else
            {
                foreach (Shape slideMasterShape in activePpt.SlideMaster.Shapes)
                {
                    if (slideMasterShape.Name.Equals(Properties.Settings.Default.ShapName))
                    {
                        // MessageBox.Show("heheheeh");
                        slideMasterShape.Delete();
                    }
                    //slideMasterShape.
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleButtonPublic_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Public");

            var activePpt = Globals.ThisAddIn.Application.ActivePresentation;

            foreach (Shape slideMasterShape in activePpt.SlideMaster.Shapes)
            {
                if (slideMasterShape.Name.Equals(Properties.Settings.Default.ShapName))
                {
                    // MessageBox.Show("heheheeh");
                    slideMasterShape.Delete();
                }
                //slideMasterShape.
            }
        }


        /// <summary>
        /// 
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
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleButtonMarkNo_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.Ribbons.Sensitive.toggleButtonMarkYes.Checked = false;
            Globals.Ribbons.Sensitive.toggleButtonMarkNo.Checked = true;


            Properties.Settings.Default.IsMask = false;
            Properties.Settings.Default.Save();

            var activePpt = Globals.ThisAddIn.Application.ActivePresentation;

            foreach (Shape slideMasterShape in activePpt.SlideMaster.Shapes)
            {
                if (slideMasterShape.Name.Equals(Properties.Settings.Default.ShapName))
                {
                    // MessageBox.Show("heheheeh");
                    slideMasterShape.Delete();
                }
                //slideMasterShape.
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


        /// <summary>
        /// clear the slide master picture
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonClearTags_Click(object sender, RibbonControlEventArgs e)
        {
            var activePpt = Globals.ThisAddIn.Application.ActivePresentation;

            foreach (Shape slideMasterShape in activePpt.SlideMaster.Shapes)
            {
                if (slideMasterShape.Name.Equals(Properties.Settings.Default.ShapName))
                {
                    // MessageBox.Show("heheheeh");
                    slideMasterShape.Delete();
                }
                //slideMasterShape.
            }
        }
    }
}
