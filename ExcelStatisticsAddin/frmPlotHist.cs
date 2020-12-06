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

        private int m_XUnits = 1;

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
                String s = String.Format("=FREQUENCY({0}, {1})", rgData.Address[false, false], rgBins.Address[false, false]);
                rgOutput[1,1].Value = "frequency";
                rgOutput = rgOutput.Offset[1, 0];
                rgOutput = rgOutput.Resize[rgBins.Rows.Count, 1];
                rgOutput.FormulaArray = s;
               
                // create a series of bars and populate them with data
                var seriesA = new OxyPlot.Series.ColumnSeries()
                {
                    Title = "Series A",
                    StrokeColor = OxyPlot.OxyColors.Black,
                    FillColor = OxyPlot.OxyColors.Red,
                    StrokeThickness = 1
                };
                seriesA.ColumnWidth = 1;
                seriesA.LabelPlacement = OxyPlot.Series.LabelPlacement.Outside;

                for (int i = 0; i < rgOutput.Rows.Count; i++)
                {
                    //values[i] = (long)(rgOutput.Cells[i + 1, 1].Value);
                    seriesA.Items.Add(new OxyPlot.Series.ColumnItem(rgOutput.Cells[i + 1, 1].Value, i));
                }
                
                // create a model and add the bars into it
                var model = new OxyPlot.PlotModel
                {
                    Title = "Histogram"
                };

                String[] catLabels = new String[rgOutput.Rows.Count];
                double dvalue, remainder, eps;
                double BinSize = rgBins[2, 1].Value - rgBins[1, 1].Value;
                double Xunit = m_XUnits * BinSize; 
                eps = m_XUnits*BinSize / 1000;

                for (int i = 0; i < rgBins.Rows.Count; i++)
                {
                    dvalue = rgBins.Cells[i + 1, 1].Value;

                    remainder = Math.Abs(Math.IEEERemainder(dvalue, Xunit));
                    if (remainder < eps)
                    {
                        catLabels[i] = String.Format("{0}", dvalue);
                    }
        
                }

                model.Axes.Add(new OxyPlot.Axes.CategoryAxis()
                {
                    ItemsSource = catLabels,
                    Angle = 90
                }); 
               
                
                model.Series.Add(seriesA);
                
                plotView1.Model = model;
                
            }
            catch (Exception ex)
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
            refedData.Select();
        }

        private void histogram1_Load(object sender, EventArgs e)
        {

        }


        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                m_XUnits = Convert.ToInt32(numericUpDown1.Value);
            }
            catch { }
        }
    }
}
