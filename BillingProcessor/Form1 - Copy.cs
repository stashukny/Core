using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace BillingProcessor
{
    public partial class Form1 : Form
    {

        public bool formulasInserted;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;


            int rCnt = 0;



            string filepath = @"C:\Users\sshatkin\Documents\Billing\Client Templates\Alaska Federal Credit Union\Cor360BillingTemplate - Alaska Federal Credit Union.xlsx";

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filepath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {

                if (range.Cells[rCnt, 6].Value2 != null)
                {
                    if (range.Cells[rCnt, 6].Value2.ToString().Trim() == "Payment Processing SaaS fee")
                    {
                        //insert next new row
                        Excel.Range Line = (Excel.Range)xlWorkSheet.Rows[rCnt + 1];
                        Line.Insert();
                        range.Cells[rCnt + 1, 6].Value2 = "Expense Report SaaS fee";
                        range.Cells[rCnt + 1, 7].Value2 = 0;
                    }

                    if (range.Cells[rCnt, 6].Value2.ToString().Trim() == "Email Monthly Fee")
                    {
                        //insert next new row
                        Excel.Range Line = (Excel.Range)xlWorkSheet.Rows[rCnt + 1];
                        Line.Insert();
                        range.Cells[rCnt + 1, 2].Value2 = "Single Sign-On Monthly Fee";
                        range.Cells[rCnt + 1, 6].Value2 = "Single Sign-On Monthly Fee";
                        range.Cells[rCnt + 1, 7].Value2 = 0;
                        range.Cells[rCnt + 1, 8].Value2 = "Specify Single Sign-On Monthly Fee";
                        xlWorkSheet.Range[range.Cells[rCnt + 1, 2], range.Cells[rCnt + 1, 3]].Merge();

                        //break;
                    }

                }

            }

            xlApp.DisplayAlerts = false;
            xlWorkBook.Save();

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;


            int rCnt = 0;



            string filepath = @"C:\Users\sshatkin\Documents\Billing\Client Templates\Alaska Federal Credit Union\Cor360BillingTemplate - Alaska Federal Credit Union.xlsx";

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filepath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {
                if (range.Cells[rCnt, 1].Value2 != null)
                {
                    if (range.Cells[rCnt, 1].Value2.ToString().Trim() == "Subtotal of Charges: USD")
                    {
                        int Position = rCnt - 5;

                        Excel.Range Line = (Excel.Range)xlWorkSheet.Rows[Position];
                        Line.Insert();
                        string initialFormula = range.Cells[Position - 1, 1].Formula;
                        string updateFormula = initialFormula.Substring(initialFormula.Length - 10);
                        int rowNum = Convert.ToInt32(updateFormula.Replace("), \" \")", "").Replace("F", ""));
                        range.Cells[Position, 1].Formula = initialFormula.Replace((rowNum).ToString(), (rowNum + 1).ToString()).Replace((Position - 1).ToString(), Position.ToString());
                        xlWorkSheet.Range[range.Cells[Position, 4], range.Cells[Position, 5]].Merge();
                        range.Cells[Position, 4].Formula = range.Cells[Position - 1, 4].Formula.ToString().Replace(rowNum.ToString(), (rowNum + 1).ToString());

                        Position = rCnt - 2;

                        Line = (Excel.Range)xlWorkSheet.Rows[Position];
                        Line.Insert();
                        initialFormula = range.Cells[Position - 1, 1].Formula;
                        updateFormula = initialFormula.Substring(initialFormula.Length - 10);
                        rowNum = rowNum = Convert.ToInt32(updateFormula.Replace("), \" \")", "").Replace("F", ""));
                        range.Cells[Position, 1].Formula = initialFormula.Replace((rowNum).ToString(), (rowNum + 1).ToString()).Replace((Position - 1).ToString(), Position.ToString());
                        xlWorkSheet.Range[range.Cells[Position, 4], range.Cells[Position, 5]].Merge();
                        range.Cells[Position, 4].Formula = range.Cells[Position - 1, 4].Formula.ToString().Replace(rowNum.ToString(), (rowNum + 1).ToString());

                        break;
                    }
                }
            }

            xlApp.DisplayAlerts = false;
            xlWorkBook.Save();

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DirSearch(@"C:\Users\sshatkin\Documents\Billing\Client Templates\");
        }

        private List<String> DirSearch(string sDir)
        {

            List<String> files = new List<String>();

            try
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"listClients.txt", true))                    
                {
                    file.AutoFlush = true;
                    foreach (string f in Directory.GetFiles(sDir))
                    {
                        files.Add(f);

                        if (f.Contains("xlsx") && !f.Contains("~") && !f.Contains("copy"))
                        {
                            file.WriteLine(f);
                        }

                    }
                }
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    files.AddRange(DirSearch(d));

                }
            }

            catch (System.Exception excpt)
            {
                MessageBox.Show(excpt.Message);
            }

            return files;
        }

        private void processData ()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;


            int rCnt = 0;



            string filepath = @"C:\Users\sshatkin\Documents\Billing\Client Templates\Alaska Federal Credit Union\Cor360BillingTemplate - Alaska Federal Credit Union.xlsx";

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filepath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {

                if (range.Cells[rCnt, 6].Value2 != null)
                {
                    if (range.Cells[rCnt, 6].Value2.ToString().Trim() == "Payment Processing SaaS fee")
                    {
                        //insert next new row
                        Excel.Range Line = (Excel.Range)xlWorkSheet.Rows[rCnt + 1];
                        Line.Insert();
                        range.Cells[rCnt + 1, 6].Value2 = "Expense Report SaaS fee";
                        range.Cells[rCnt + 1, 7].Value2 = 0;
                    }

                    if (range.Cells[rCnt, 6].Value2.ToString().Trim() == "Email Monthly Fee")
                    {
                        //insert next new row
                        Excel.Range Line = (Excel.Range)xlWorkSheet.Rows[rCnt + 1];
                        Line.Insert();
                        range.Cells[rCnt + 1, 2].Value2 = "Single Sign-On Monthly Fee";
                        range.Cells[rCnt + 1, 6].Value2 = "Single Sign-On Monthly Fee";
                        range.Cells[rCnt + 1, 7].Value2 = 0;
                        range.Cells[rCnt + 1, 8].Value2 = "Specify Single Sign-On Monthly Fee";
                        xlWorkSheet.Range[range.Cells[rCnt + 1, 2], range.Cells[rCnt + 1, 3]].Merge();

                        //break;
                    }

                }

            }

            xlApp.DisplayAlerts = false;
            xlWorkBook.Save();

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
    }
}
