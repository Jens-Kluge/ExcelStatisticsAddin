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
            
            try
            {
                //GetValues
                Utilities.GetRange(ref rgData, refedData.Text, ref ok);
                Utilities.GetRange(ref rgBins, refedBins.Text, ref ok);
                Utilities.GetRange(ref rgOutput, refedOutput.Text, ref ok);

                //Todo: Check ranges
                if(!IsValidBinRange(rgBins)) {
                    MessageBox.Show("Select a valid bin range");
                    return;
                }
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

                //create a model and add the series to it
                var model = new OxyPlot.PlotModel
                {
                    Title = "Histogram"
                };

                if (chkOverlay.Checked)
                {
                    double alpha = 3, beta = 2;
                    EstimateWeibullDistParams(rgData, ref alpha, ref beta);
                    double value;

                    double scalefactor = rgData.Rows.Count*(rgBins[2,1].Value-rgBins[1,1].Value);

                    //line series overlay
                    var lsFit = new OxyPlot.Series.LineSeries()
                    {
                        Color = OxyPlot.OxyColors.Green
                    };
                   
                    for (int i = 0; i < rgOutput.Rows.Count; i++)
                    {
                        value = Globals.ThisAddIn.Application.WorksheetFunction.Weibull(rgBins[i + 1, 1].Value, beta, alpha, false);
                        value *= scalefactor;
                        lsFit.Points.Add(new OxyPlot.DataPoint(i + 0.5, value));
                    }
                    model.Series.Add(lsFit);
                }
               

                // Add labels to the X-Axis
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

        private void EstimateWeibullDistParams(Range rgData, ref double alpha, ref double beta)
        {
            //use the power density method 
            double x3_avg = Globals.ThisAddIn.Application.WorksheetFunction.SumProduct(rgData, rgData, rgData);
            long samplesize = rgData.Rows.Count;
            x3_avg = x3_avg / samplesize;
            double avg = Globals.ThisAddIn.Application.WorksheetFunction.Average(rgData);
            double avgx3 = Math.Pow(avg,3);
            double epattern = x3_avg / avgx3;
            beta = 1 + 3.69 / Math.Pow(epattern, 2);
            alpha = avg / Globals.ThisAddIn.Application.WorksheetFunction.Gamma(1 + 1 / beta);
        }

        private bool IsValidBinRange(Range rgBins)
        {
            bool bRet=true;
            try
            {
                if (rgBins is null)
                {
                    return false;
                }
                //there need to be at least two Bins
                if (rgBins.Rows.Count <= 2)
                {
                    return false;
                }
                //the bins should be arranged in one column
                if (rgBins.Columns.Count > 1)
                {
                    return false;
                }
                //the bins should be in ascending order
                for (int i = 0; i <= (rgBins.Rows.Count - 1); i++)
                {
                    if (rgBins[i + 2, 1].Value <= rgBins[i + 1, 1].Value)
                    {
                        return false;
                    }

                }
                return bRet;
            }
            catch(Exception ex) {
                MessageBox.Show(ex.Message);
                return false;
            }
        }
    }
}
