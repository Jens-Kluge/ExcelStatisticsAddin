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
    /// <summary>
    /// Form to call the Distribution fit
    /// </summary>
    public partial class frmDistFit : Form
    {
        /// <summary>
        /// Constructor of frmDistFit
        /// </summary>
        public frmDistFit()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        /*  
         *  Retrieve input range and output ranges from content of SelEdit controls
         *  Check if the ranges are valid
         *  Return the retrieved ranges in the reference parameters
         */
        private bool IsValid(ref Range rgIn, ref Range rgOut)
        {
            // return true if input is valid
            bool ok = false;

            GetRange(ref rgIn, refedit1.Text, ref ok);
            if (!ok)
            {
                MessageBox.Show("Invalid input range selected");
                return false;
            }

            long m;     //number of data rows in rg
            m = rgIn.Rows.Count;
            long n;     //number of columns in rg
            n = rgIn.Columns.Count;

            if (n != 1)
            {
                MessageBox.Show("Input Range must have one column");
                return false;
            }

            if (m < 2)
            {
                MessageBox.Show("Input Range must have at least two rows");
                return false;
            }

            if (Globals.ThisAddIn.Application.WorksheetFunction.Count(rgIn) != m)
            {
                MessageBox.Show("Input Range must not contain any non-numeric values");
                return false;
            }

            GetOutputRange(ref rgOut, refedit2.Text, ref ok);
            if (!ok)
            {
                MessageBox.Show("Invalid output range selected");
                return false;
            }
            return IsEmptyRange(rgIn, rgOut);
        }

        private bool IsEmptyRange(Range rgIn, Range rgOut)
        {

            //return true if output range is empty
            long m;       // # rows in output range
            long n;       // # columns in output range
            bool retval;

            m = 12;
            n = 9;

            Range rg1 = null;        //range under consideration
            SetRange2(ref rg1, rgOut, m, n);

            if (Globals.ThisAddIn.Application.WorksheetFunction.CountA(rg1) == 0)
            {
                retval = true;
            }
            else
            {
                retval = MessageBox.Show("Overwrite existing data?", "", MessageBoxButtons.OKCancel) == DialogResult.OK;
            }
            return retval;
        }

        void SetRange2(ref Range rg, Range rg0, long r, long c)
        {
            // set rg to a range with first cell in rg and height r and width c
            rg = rg0.Cells[1, 1].Resize(r, c);
        }

        void GetRange(ref Range rg, String s, ref bool ok)
        {
            // Get range with address s, set ok to true if this address is valid
            ok = true;
            try
            {
                rg = Globals.ThisAddIn.Application.ActiveSheet.Range(s);
                return;
            }
            catch (Exception ex)
            {
                ok = false;
                return;
            }
        }

        void GetOutputRange(ref Range rg, String s, ref bool ok)
        {
            //Get output range with address s, set ok to true if this address is valid
            // If s = "" then set rg to a new worksheet
            Worksheet ws;
            ok = true;
            try
            {
                if (s == "")
                {
                    ws = Globals.ThisAddIn.Application.Worksheets.Add();
                    rg = ws.Range["A1"];
                }
                else
                {
                    rg = Globals.ThisAddIn.Application.ActiveSheet.Range(s);
                }
                return;
            }
            catch (Exception ex)
            {
                ok = false;
                return;
            }
        }

        /// <summary>
        ///  Launch the distribution fitting
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOK_Click(object sender, EventArgs e)
        {
            //OK button is clicked
            Range rg = null;
            Range ce = null;
            long dist;

            dist = 3; //Weibull

            if (IsValid(ref rg, ref ce)) {
                RunDistFit(rg, ce);
            }
        }

        void RunDistFit(Range rgIn, Range rgout)
        {
            int rowOut = ((Range) rgout.Cells.Item[1, 1]).Row;
            int colOut = ((Range) rgout.Cells.Item[1, 1]).Column;

            int rowsIn = rgIn.Rows.Count;
            String sAddrFrom = rgIn.Cells.Item[1, 1].AddressLocal[false, false]; 
            String sAddrTo = rgIn.Cells.Item[rowsIn, 1].AddressLocal[false, false];
            String s;
            String sCellrefMean, sCellrefStd, sCellrefBeta1, sCellrefBeta;
            try
            {
                Worksheet ws = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                // compute Weibull parameters alpha, beta using method of moments

                if (rbMoments.Checked) { 
                    //average of input range
                    s = String.Format("=AVERAGE({0}:{1})", sAddrFrom, sAddrTo);
                    rgout[1, 1].Value = "mean";
                    rgout.Cells[1, 2].Formula = s;
                    sCellrefMean = ((Range)rgout.Cells[1, 2]).AddressLocal[false, false];

                    //standard deviation of input range
                    s = String.Format("=STDEV.S({0}:{1})", sAddrFrom, sAddrTo);
                    rgout.Cells[2, 1].Value = "deviation";
                    rgout.Cells[2, 2].Formula = s;
                    sCellrefStd = ((Range)rgout.Cells[2, 2]).AddressLocal;

                    //initial guess for beta
                    rgout[1, 4].Value = "β (initial guess)";
                    s = String.Format("=0.5"); //to be replaced by an appropriate formula
                    rgout[1, 5].Formula = s;
                    sCellrefBeta1 = ((Range)rgout.Cells[1, 5]).AddressLocal;

                    //copy this value into the output cell for beta
                    rgout[4, 4].Value = "β";
                    rgout[4, 5].Value = rgout[1, 5].Value;
                    sCellrefBeta = ((Range)rgout.Cells[4, 5]).AddressLocal;

                    //calculate alpha from beta and mean value
                    rgout[3, 4].Value = "α";
                    s = String.Format("={0}/EXP(GAMMALN(1 + 1/{1}))", sCellrefMean, sCellrefBeta);
                    rgout[3, 5].Formula = s;

                    //implicit formula for beta
                    //to solve this for beta adjust beta such that the value is close to zero
                    rgout.Cells[4, 1].Value = "implicit formula";
                    s = String.Format("=GAMMALN(1+2/{0})-2*GAMMALN(1+1/{0})-LN({1}^2+{2}^2)+2*LN({1})", sCellrefBeta, sCellrefMean, sCellrefStd);
                    rgout.Cells[4, 2].Formula = s;

                    //Call the Goalseeker
                    ((Range)rgout.Cells[4, 2]).GoalSeek(0, rgout.Cells[4, 5]);
                }
                else if(rbDensity.Checked)
                {
                    //average of input range
                    s = String.Format("=AVERAGE({0}:{1})", sAddrFrom, sAddrTo);
                    rgout[1, 1].Value = "mean";
                    rgout.Cells[1, 2].Formula = s;
                    sCellrefMean = ((Range)rgout.Cells[1, 2]).AddressLocal[false, false];

                    //average of third power
                    rgout.Cells[2, 1].Value = "mean(x^3)";
                    s = String.Format("=SUMPRODUCT({0}:{1}^3)/COUNT({0}:{1})", sAddrFrom, sAddrTo);
                    rgout.Cells[2, 2].Formula = s;
                    String sCellrefX3 = ((Range)rgout.Cells[2, 2]).AddressLocal;

                    //third power of average
                    rgout[3, 1].Value = "mean(x)^3";
                    s = String.Format("=AVERAGE({0}:{1})^3", sAddrFrom, sAddrTo);
                    rgout.Cells[3, 2].Formula = s;
                    String sCellrefAVG3 = ((Range)rgout.Cells[3, 2]).AddressLocal[false, false];

                    //energy pattern
                    rgout[4, 1].Value = "energy pattern";
                    s = String.Format("={0}/{1}", sCellrefX3, sCellrefAVG3);
                    rgout.Cells[4, 2].Formula = s;
                    String sCellrefPattern = ((Range)rgout.Cells[4, 2]).AddressLocal[false, false];

                    //alpha
                    rgout[5, 1].Value = "β";
                    s = String.Format("=1 + 3.69/{0}^2", sCellrefPattern);
                    rgout.Cells[5, 2].Formula = s;
                    String sCellrefAlpha = ((Range)rgout.Cells[5, 2]).AddressLocal[false, false];

                    //beta
                    rgout[6, 1].Value = "α";
                    s = String.Format("={0}/GAMMA(1 + 1/{1})", sCellrefMean, sCellrefAlpha);
                    rgout.Cells[6, 2].Formula = s;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void frmDistFit_Load(object sender, EventArgs e)
        {
            rbMoments.Checked = true;
        }

        private void btnFill_Click(object sender, EventArgs e)
        {
            Range rg1= null, rg2=null;
            bool ok = false;

            GetRange(ref rg1, refedit1.Text, ref ok);
            Utilities.ExtendRange(rg1, ref rg2);
            rg2.Select();
            refedit1.Text = rg2.Address[false, false];
        }
    }
}
