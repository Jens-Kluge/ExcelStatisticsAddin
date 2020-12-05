using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;


namespace ExcelStatisticsAddin
{
    public partial class frmPlotHist : Form
    {
        public frmPlotHist()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Range rgData = null, rgBins = null, rgOutput=null;
            bool ok = false;
            //Check ranges
            try
            {
                //GetValues
                Utilities.GetRange(ref rgData, refedData.Text, ref ok);
                Utilities.GetRange(ref rgBins, refedBins.Text, ref ok);
                Utilities.GetRange(ref rgOutput, refedOutput.Text, ref ok);

                //Plot the histogram
                long[] values = new long[rgBins.Rows.Count];
                String s = String.Format("=FREQUENCY({0}, {1})", rgData.Address[false,false], rgBins.Address[false,false]);
                rgOutput = rgOutput.Resize[rgBins.Rows.Count, 1];
                rgOutput.FormulaArray = s;
                for (int i = 0; i < rgOutput.Rows.Count; i++)
                {
                    values[i] = (long)(rgOutput.Cells[i + 1, 1].Value);
                }
                histogram1.DrawHistogram(values);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnExtendRg1_Click(object sender, EventArgs e)
        {
            Range rg1 = null, rg2 = null;
            bool ok = false;

            Utilities.GetRange(ref rg1, refedData.Text, ref ok);
            Utilities.ExtendRange(rg1, ref rg2);
            rg2.Select();
            refedData.Text = rg2.Address[false, false];
        }

        private void btnExtendRg2_Click(object sender, EventArgs e)
        {
            Range rg1 = null, rg2 = null;
            bool ok = false;

            Utilities.GetRange(ref rg1, refedBins.Text, ref ok);
            Utilities.ExtendRange(rg1, ref rg2);
            rg2.Select();
            refedBins.Text = rg2.Address[false, false];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmPlotHist_Load(object sender, EventArgs e)
        {
            refedData.Focus();
        }

        private void histogram1_Load(object sender, EventArgs e)
        {

        }
    }
}
