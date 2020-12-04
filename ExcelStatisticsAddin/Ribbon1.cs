using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelStatisticsAddin
{
    public partial class Ribbon1
    {
        private frmMain fmMain;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if(fmMain is null || fmMain.IsDisposed)
            {
                fmMain = new frmMain();
            }
            if(!fmMain.Visible)
            {
                fmMain.Show();
            }  
            else
            {
                fmMain.BringToFront();
            }
        }
    }
}
