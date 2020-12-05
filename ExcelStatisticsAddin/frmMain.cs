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
        private frmPlotHist fmPlotHist;

        public frmMain()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if(listBox1.SelectedIndex == 0)
            {
                LoadDistFitWindow();
            }
            else if(listBox1.SelectedIndex == 1)
            {
                LoadDistPlotWindow();
            }
        }

        private void LoadDistFitWindow()
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

        private void LoadDistPlotWindow()
        {
            if (fmPlotHist is null || fmPlotHist.IsDisposed)
            {
                fmPlotHist = new frmPlotHist();
                // Initialize input and output selection controls
                fmPlotHist.refedData._Excel = Globals.ThisAddIn.Application;
                fmPlotHist.refedBins._Excel = Globals.ThisAddIn.Application;
                fmPlotHist.refedOutput._Excel = Globals.ThisAddIn.Application;
            }
            if (!fmPlotHist.Visible)
            {
                fmPlotHist.Show();
            }
            else
            {
                fmPlotHist.BringToFront();
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
