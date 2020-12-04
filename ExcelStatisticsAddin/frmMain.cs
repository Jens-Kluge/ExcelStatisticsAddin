using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelStatisticsAddin
{
    public partial class frmMain : Form
    {
        private frmDistFit fmDistFit;

        public frmMain()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (fmDistFit is null || fmDistFit.IsDisposed) 
            {
                fmDistFit = new frmDistFit();
                // Initialize input and output selection controls
                fmDistFit.refedit1._Excel = Globals.ThisAddIn.Application;
                fmDistFit.refedit2._Excel = Globals.ThisAddIn.Application;
            }
            if (!fmDistFit.Visible)
            {
                fmDistFit.Show();
            }
            else
            {
                fmDistFit.BringToFront();
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            listBox1.SelectedIndex = 0;
        }
    }
}
