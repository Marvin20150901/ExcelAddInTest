using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

        }
    }
}
