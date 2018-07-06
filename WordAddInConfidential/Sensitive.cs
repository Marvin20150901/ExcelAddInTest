using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Shape = Microsoft.Office.Interop.Word.Shape;

using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace WordAddInConfidential
{
    public partial class Sensitive
    {
        private  string appPath;
        private void Sensitive_Load(object sender, RibbonUIEventArgs e)
        {
            appPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\OfficeAddinConfidential\";

            if (!Directory.Exists(appPath))
            {
                Directory.CreateDirectory(appPath);
            }

            if (!File.Exists(appPath+@"Secret.pang"))
            {
                Properties.Resources.Secret.Save(appPath+@"Secret.png");
                Properties.Resources.Internal.Save(appPath+@"Internal.png");
                Properties.Resources.Confidential.Save(appPath+@"Confidential.png");

                Properties.Settings.Default.IsImageUpdata = false;
                Properties.Settings.Default.Save();
            }

        }

        private void toggleButtonSecret_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Secret");
            //            var activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            var activeWord = Globals.ThisAddIn.Application.ActiveDocument;
            //var size = Properties.Resources.Secret.Size;
            if (Properties.Settings.Default.IsMask)
            {

                foreach (Word.Section section in activeWord.Sections)
                {

                    //Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    //headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    //headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

                    //section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range
                    foreach (Shape headerShape in section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes)
                    {
                        if (headerShape.Name.Equals(Properties.Settings.Default.ShapName))
                        {
                            headerShape.Delete();
                        }
                    }

                    if (!File.Exists(appPath + @"Secret.png"))
                    {
                        Properties.Resources.Secret.Save(appPath + @"Secret.png");
                    }


                    //Word.InlineShape varShape= headerRange.InlineShapes.AddPicture("Secret.png");

                    //Word.Shape varShape=section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.AddPicture("Secret.png",ref lt,ref sw,ref lf,ref tp,ref wd,ref hg);
                    Word.Shape varShape = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes
                        .AddPicture(appPath + @"Secret.png");

                    varShape.RelativeHorizontalPosition =
                        Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                    varShape.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    varShape.Left =10;
                    varShape.Top = 10;
                    varShape.Name = Properties.Settings.Default.ShapName;
                }
            }
            
            else
            {
                //delete the image header

                foreach (Word.Section section in activeWord.Sections)
                {
                    foreach (Shape headerShape in section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes)
                    {
                        if (headerShape.Name.Equals(Properties.Settings.Default.ShapName))
                        {
                            headerShape.Delete();
                        }
                    }
                }
                 
            }
        }

        private void toggleButtonConfidential_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Confidential");
            //            var activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            var activeWord = Globals.ThisAddIn.Application.ActiveDocument;
            //var size = Properties.Resources.Secret.Size;
            if (Properties.Settings.Default.IsMask)
            {

                foreach (Word.Section section in activeWord.Sections)
                {

                    //Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    //headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    //headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

                    //section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range
                    foreach (Shape headerShape in section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes)
                    {
                        if (headerShape.Name.Equals(Properties.Settings.Default.ShapName))
                        {
                            headerShape.Delete();
                        }
                    }

                    if (!File.Exists(appPath + @"Confidential.png"))
                    {
                        Properties.Resources.Confidential.Save(appPath + @"Confidential.png");
                    }


                    //Word.InlineShape varShape= headerRange.InlineShapes.AddPicture("Secret.png");

                    //Word.Shape varShape=section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.AddPicture("Secret.png",ref lt,ref sw,ref lf,ref tp,ref wd,ref hg);
                    Word.Shape varShape = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes
                        .AddPicture(appPath + @"Confidential.png");

                    varShape.RelativeHorizontalPosition =
                        Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                    varShape.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    varShape.Top = 10;
                    varShape.Left = 10;
                    varShape.Name = Properties.Settings.Default.ShapName;
                    
                }
            }

            else
            {
                //delete the image header

                foreach (Word.Section section in activeWord.Sections)
                {
                    foreach (Shape headerShape in section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes)
                    {
                        if (headerShape.Name.Equals(Properties.Settings.Default.ShapName))
                        {
                            headerShape.Delete();
                        }
                    }
                }

            }
        }

        private void toggleButtonInternal_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Internal");
            //            var activeSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            var activeWord = Globals.ThisAddIn.Application.ActiveDocument;
            //var size = Properties.Resources.Secret.Size;
            if (Properties.Settings.Default.IsMask)
            {

                foreach (Word.Section section in activeWord.Sections)
                {

                    //Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    //headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    //headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

                    //section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range
                    foreach (Shape headerShape in section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes)
                    {
                        if (headerShape.Name.Equals(Properties.Settings.Default.ShapName))
                        {
                            headerShape.Delete();
                        }
                    }

                    if (!File.Exists(appPath + @"Internal.png"))
                    {
                        Properties.Resources.Internal.Save(appPath + @"Internal.png");
                    }


                    //Word.InlineShape varShape= headerRange.InlineShapes.AddPicture("Secret.png");

                    //Word.Shape varShape=section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.AddPicture("Secret.png",ref lt,ref sw,ref lf,ref tp,ref wd,ref hg);
                    Word.Shape varShape = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes
                        .AddPicture(appPath + @"Internal.png");

                    varShape.RelativeHorizontalPosition =
                        Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                    varShape.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    varShape.Top = 10;
                    varShape.Left = 10;
                    varShape.Name = Properties.Settings.Default.ShapName;
                }
            }

            else
            {
                //delete the image header

                foreach (Word.Section section in activeWord.Sections)
                {
                    foreach (Shape headerShape in section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes)
                    {
                        if (headerShape.Name.Equals(Properties.Settings.Default.ShapName))
                        {
                            headerShape.Delete();
                        }
                    }
                }

            }
        }

        private void toggleButtonPublic_Click(object sender, RibbonControlEventArgs e)
        {
            SetWorkbookSensitive("Public");
            var activeWord = Globals.ThisAddIn.Application.ActiveDocument;
            foreach (Word.Section section in activeWord.Sections)
            {
                foreach (Shape headerShape in section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes)
                {
                    if (headerShape.Name.Equals(Properties.Settings.Default.ShapName))
                    {
                        headerShape.Delete();
                    }
                }
            }
        }

        private void buttonClearTags_Click(object sender, RibbonControlEventArgs e)
        {
            var activeWord = Globals.ThisAddIn.Application.ActiveDocument;
            foreach (Word.Section section in activeWord.Sections)
            {
                foreach (Shape headerShape in section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes)
                {
                    if (headerShape.Name.Equals(Properties.Settings.Default.ShapName))
                    {
                        headerShape.Delete();
                    }
                }
            }
        }

        private void toggleButtonMarkYes_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.Ribbons.Sensitive.toggleButtonMarkYes.Checked = true;
            Globals.Ribbons.Sensitive.toggleButtonMarkNo.Checked = false;

            Properties.Settings.Default.IsMask = true;
            Properties.Settings.Default.Save();
        }

        private void toggleButtonMarkNo_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.Ribbons.Sensitive.toggleButtonMarkYes.Checked = false;
            Globals.Ribbons.Sensitive.toggleButtonMarkNo.Checked = true;


            Properties.Settings.Default.IsMask = false;
            Properties.Settings.Default.Save();

            var activeWord = Globals.ThisAddIn.Application.ActiveDocument;
            foreach (Word.Section section in activeWord.Sections)
            {
                foreach (Shape headerShape in section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes)
                {
                    if (headerShape.Name.Equals(Properties.Settings.Default.ShapName))
                    {
                        headerShape.Delete();
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

                Microsoft.Office.Core.DocumentProperties prp = Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties;

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
