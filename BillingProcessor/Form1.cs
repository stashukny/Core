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
        public float Sum = 0;
        public int newTransactionFeeLine;
        public int newMonthlyFeeLine;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnAddScheduleA_Click(object sender, EventArgs e)
        {
            cycleThroughFiles("schedule");
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

        private void btnAddFormulas_Click(object sender, EventArgs e)
        {
            cycleThroughFiles("formulas");
        }

        private void btnListFiles_Click(object sender, EventArgs e)
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

        private void processData(string insertType, string filepath)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;


            int rCnt = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filepath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            if (xlWorkSheet.ProtectContents == true)
            {
                xlWorkSheet.Unprotect("c0r36o");
            }

            range = xlWorkSheet.UsedRange;

            if (insertType == "specific_formula")
            {
                updateSpecficFormula(ref xlWorkSheet, ref range, true, true);
            }

            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {

                if (insertType == "schedule")
                {
                    if (range.Cells[rCnt, 6].Value2 != null)
                    {
                        if (range.Cells[rCnt, 6].Value2.ToString().Trim() == "Single Sign-On Monthly Fee")
                        {
                            //insert next new row
                            Excel.Range Line = (Excel.Range)xlWorkSheet.Rows[rCnt + 1];
                            Line.Insert();
                            range.Cells[rCnt + 1, 6].Value2 = "PO Requisition Transactions";


                            Excel.Range Line2 = (Excel.Range)xlWorkSheet.Rows[rCnt + 2];
                            Line2.Insert();


                            range.Cells[rCnt + 2, 7].Value2 = 0.5;
                            range.Cells[rCnt + 2, 6].Value2 = "PO Requisition Transactions";
                            range.Cells[rCnt + 2, 5].Value2 = 0.5;
                            range.Cells[rCnt + 2, 2].Value2 = " PO Requisition Transaction Fee";
                            range.Cells[rCnt + 2, 1].Value2 = 1;
                            xlWorkSheet.Range[range.Cells[rCnt + 2, 2], range.Cells[rCnt + 2, 3]].Merge();

                            int i = 1;
                            for (i = 1; i < 9; i++ )
                            { 
                                range.Cells[rCnt + 1, i].Interior.ColorIndex = 1;
                                range.Cells[rCnt + 1, i].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            }
                        }

                    }
                }
                else if (insertType == "formulas")
                {
                    if (range.Cells[rCnt, 1].Value2 != null)
                    {
                        if (range.Cells[rCnt, 1].Value2.ToString().Trim() == "Subtotal of Charges: USD")
                        {
                            int Position = rCnt - 11;
                            updateFormulas(ref xlWorkSheet, ref range, Position, 25);                            

                            break;
                        }
                    }
                }
            }

            xlApp.DisplayAlerts = false;

            if (xlWorkSheet.ProtectContents == false)
            {
                xlWorkSheet.Protect("c0r36o");
            }
            
            xlWorkBook.Save();


            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void cycleThroughFiles(string insertType)
        {
            using (System.IO.StreamReader file = new System.IO.StreamReader(@"MasterClientList.txt"))
            {
                while (true)
                {
                    string path = file.ReadLine();
                    if (path == null) { break; }
                    processData(insertType, path);
                }
            }
        }

        private void btnRunAll_Click(object sender, EventArgs e)
        {
            cycleThroughFiles("schedule");
            cycleThroughFiles("formulas");
            cycleThroughFiles("specific_formula");
        }

        private void updateFormulas(ref Excel.Worksheet xlWorkSheet, ref Excel.Range range, int Position, int offset)
        {

            Excel.Range Line = (Excel.Range)xlWorkSheet.Rows[Position];
            Line.Insert();
            string initialFormula = range.Cells[Position - 1, 1].Formula;
            string updateFormula = initialFormula.Substring(initialFormula.Length - 10);
            int rowNum = Convert.ToInt32(updateFormula.Replace("), \" \")", "").Replace("F", ""));
            range.Cells[Position, 1].Formula = initialFormula.Replace((rowNum).ToString(), (rowNum + offset).ToString()).Replace((Position - 1).ToString(), Position.ToString());
            xlWorkSheet.Range[range.Cells[Position, 4], range.Cells[Position, 5]].Merge();
            range.Cells[Position, 4].Formula = range.Cells[Position - 1, 4].Formula.ToString().Replace(rowNum.ToString(), (rowNum + offset).ToString()).Replace("A" + (rowNum + offset).ToString() + "*", "");
            range.Cells[Position, 4].Formula = range.Cells[Position, 4].Formula.Replace("U", "Y");

            //more formulas
            range.Cells[Position, 3].Formula = range.Cells[Position, 4].Formula.Replace(range.Cells[Position, 4].Formula.ToString().Substring(range.Cells[Position, 4].Formula.ToString().IndexOf("),") + 3, 8), "").Replace(", 0", ",' '");
            range.Cells[Position, 2].Formula = range.Cells[Position, 4].Formula.Replace(range.Cells[Position, 4].Formula.ToString().Substring(range.Cells[Position, 4].Formula.ToString().IndexOf("*"), 6), "");

            newTransactionFeeLine = Position;
        }

        private void btnModifyOneFormula_Click(object sender, EventArgs e)
        {
            cycleThroughFiles("one_formula");
        }

        private void updateSpecficFormula(ref Excel.Worksheet xlWorkSheet, ref Excel.Range range, bool processMonthly, bool processTransaction)
        {
            if (processMonthly || processTransaction)
            {               
                int rCnt = 0;

                //Expense SaaS
                int iExpenseSaaSDesc = findElement(ref xlWorkSheet, ref range, "Expense Report SaaS fee", 6);
                int iExpenseSaaSLine = findElement(ref xlWorkSheet, ref range, "F" + iExpenseSaaSDesc.ToString(), 1);


                int iSSOSDesc = findElement(ref xlWorkSheet, ref range, "Single Sign-On Monthly Fee", 6);
                int iSSOLine = findElement(ref xlWorkSheet, ref range, "F" + iSSOSDesc.ToString(), 1);


                for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                {
                    if (range.Cells[rCnt, 1].Value2 != null)
                    {
                        if (processMonthly)
                        { 
                            if (range.Cells[rCnt, 1].Formula.ToString().Trim().IndexOf("Monthly Fees") > 0)
                            {
                                range.Cells[rCnt, 1].Formula = range.Cells[rCnt, 1].Formula.Replace(" =", "=").Replace("= ", "=");
                                range.Cells[rCnt, 1].Formula = range.Cells[rCnt, 1].Formula.ToString().Replace("A" + (iExpenseSaaSLine - 1).ToString() + "=\" \",", "A" + (iExpenseSaaSLine - 1).ToString() + "=\" \"," + " A" + (iExpenseSaaSLine).ToString() + "=\" \",");
                                range.Cells[rCnt, 1].Formula = range.Cells[rCnt, 1].Formula.ToString().Replace("A" + (iSSOLine - 1).ToString() + "=\" \")", "A" + (iSSOLine - 1).ToString() + "=\" \"," + " A" + (iSSOLine).ToString() + "=\" \")");
                                //range.Cells[rCnt, 1].Formula = range.Cells[rCnt, 1].Formula.ToString().Replace("A" + (iSSOLine - 1).ToString() + " = \" \",", "A" + (iSSOLine - 1).ToString() + " = \" \"," + " A" + (iSSOLine).ToString() + " = \" \",");

                                processMonthly = false;
                            }
                        }
                    
                        if (processTransaction)
                        { 
                            if (range.Cells[rCnt, 1].Formula.ToString().Trim().IndexOf("Transaction Fees") > 0)
                            {
                                range.Cells[rCnt, 1].Formula = range.Cells[rCnt, 1].Formula.Replace(" =", "=").Replace("= ", "=");
                                range.Cells[rCnt, 1].Formula = range.Cells[rCnt, 1].Formula.ToString().Replace("A" + (newTransactionFeeLine - 1).ToString() + "=\" \",", "A" + (newTransactionFeeLine - 1).ToString() + "=\" \"," + " A" + (newTransactionFeeLine).ToString() + "=\" \",");

                                processTransaction = false;
                            }
                        }

                        if (!(processMonthly || processTransaction))
                        {
                            break;
                        }
                    }

                }
            
            }
        }

        private int findElement (ref Excel.Worksheet xlWorkSheet, ref Excel.Range range, string findString, int col)
        {
            int rCnt = 0,
                iFound = 0;

            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {
                if (range.Cells[rCnt, col].Value2 != null)
                {

                    if (range.Cells[rCnt, col].Formula.ToString().Trim().IndexOf(findString) > -1)
                    {
                        iFound = rCnt;
                        break;
                    }
                }

            }
            return iFound;
        }
    }
}
