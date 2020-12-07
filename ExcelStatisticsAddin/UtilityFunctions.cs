using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace ExcelStatisticsAddin
{
    static class Utilities
    {
        public static void GetRange(ref Range rg, String s, ref bool ok)
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

        public static void BringFormsToFront()
        {
            FormCollection fms = System.Windows.Forms.Application.OpenForms;
            if(fms is null) { return; }

            foreach(Form fm in fms)
            {
                fm.BringToFront();
            }
        }

        public static void ExtendRange(Range rg, ref Range rg1)
        {
            //For a given range rg, return range rg1 which extends the first row of rg until a row of empty cells is found
            //If first row of rg is empty then perform above on second row of rg
            long n, j;
            bool bFinish;

            Range ce;
            Range ce1;
            String s;
            
            n = rg.Columns.Count;
            ce = rg.Cells[1, 1];
            try
            {
                do
                {
                    bFinish = true;
                    ce = ce.Offset[1, 0];
                    while (!(ce.Value is null))
                    {
                        ce = ce.Offset[1, 0];
                    }

                    ce1 = ce.Offset[0, 1];
                    for (j = 2; j <= n; j++)
                    {
                        if (!(ce1.Value is null))
                        {
                            bFinish = false;
                            break;
                        }
                        ce1 = ce1.Offset[0, 1];
                    }
                } while (!bFinish);
                rg1 = rg.Cells[1, 1].Resize(ce.Row - rg.Cells[1, 1].Row, n);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }//class
}// namespace