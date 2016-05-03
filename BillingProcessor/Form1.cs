﻿using System;
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

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
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

        private void button2_Click(object sender, EventArgs e)
        {
            cycleThroughFiles("formulas");
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

        private void processData(string insertType, string filepath)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;


            int rCnt = 0;

            //string filepath = @"C:\Users\sshatkin\Documents\Billing\Client Templates\Alaska Federal Credit Union\Cor360BillingTemplate - Alaska Federal Credit Union.xlsx";

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filepath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            if (xlWorkSheet.ProtectContents == true)
            {
                xlWorkSheet.Unprotect("c0r36o");
            }

            range = xlWorkSheet.UsedRange;

            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {

                if (insertType == "schedule")
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
                else
                {
                    if (range.Cells[rCnt, 1].Value2 != null)
                    {
                        if (range.Cells[rCnt, 1].Value2.ToString().Trim() == "Subtotal of Charges: USD")
                        {
                            int Position = rCnt - 5;
                            updateFormulas(ref xlWorkSheet, ref range, Position);                            

                            Position = rCnt - 2;
                            updateFormulas(ref xlWorkSheet, ref range, Position);      

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

        private void button4_Click(object sender, EventArgs e)
        {
            cycleThroughFiles("schedule");
            cycleThroughFiles("formulas");
        }

        private void updateFormulas(ref Excel.Worksheet xlWorkSheet, ref Excel.Range range, int Position)
        {
            Excel.Range Line = (Excel.Range)xlWorkSheet.Rows[Position];
            Line.Insert();
            string initialFormula = range.Cells[Position - 1, 1].Formula;
            string updateFormula = initialFormula.Substring(initialFormula.Length - 10);
            int rowNum = Convert.ToInt32(updateFormula.Replace("), \" \")", "").Replace("F", ""));
            range.Cells[Position, 1].Formula = initialFormula.Replace((rowNum).ToString(), (rowNum + 1).ToString()).Replace((Position - 1).ToString(), Position.ToString());
            xlWorkSheet.Range[range.Cells[Position, 4], range.Cells[Position, 5]].Merge();
            range.Cells[Position, 4].Formula = range.Cells[Position - 1, 4].Formula.ToString().Replace(rowNum.ToString(), (rowNum + 1).ToString()).Replace("A" + (rowNum + 1).ToString() + "*", "");         
        }

        private void button5_Click(object sender, EventArgs e)
        {
            using (System.IO.StreamReader file = new System.IO.StreamReader(@"MasterClientList.txt"))
            {
                while (true)
                {
                    string path = file.ReadLine();
                    if (path == null) { break; }
                    updateFormula(path);
                }
            }
            MessageBox.Show(Sum.ToString());
        }

        private void updateFormula(string filepath)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;


            int rCnt = 0;            

            //string filepath = @"C:\Users\sshatkin\Documents\Billing\Client Templates\Alaska Federal Credit Union\Cor360BillingTemplate - Alaska Federal Credit Union.xlsx";

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filepath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets["August-2015 Invoice"];
            //if (xlWorkSheet.ProtectContents == true)
            //{
            //    xlWorkSheet.Unprotect("c0r36o");
            //}

            range = xlWorkSheet.UsedRange;

            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {

                //// changed formula from old to new
                //if (range.Cells[rCnt, 4].Formula != null)
                //{
                //    if (range.Cells[rCnt, 4].Formula.ToString().Trim() == "=SUM(D16:D40)")
                //    {
                //        range.Cells[rCnt, 4].Formula = "=SUM(D16:D41)";

                //        break;
                //    }
                //}

                //// check where new formula doesn't exist
                //if (range.Cells[rCnt, 4].Formula.ToString().Trim() == "=SUM(D16:D41)")
                //{
                //    using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"FormulaExists.txt", true))
                //    {
                //        file.AutoFlush = true;
                //        file.WriteLine(filepath);
                //    }

                //    break;
                //}

                if (range.Cells[rCnt, 1].Value2 != null)
                {
                    if (range.Cells[rCnt, 1].Value2.ToString().Trim() == "INVOICE TOTAL: USD")
                    {
                        Sum += range.Cells[rCnt, 4].Value2;

                        break;
                    }
                }

            }
            
                //if ((range.Cells[151, 1].Value2.ToString().Trim() == "Billing Month"))
                //{
                //    using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"DataAligned.txt", true))
                //    {
                //        file.AutoFlush = true;
                //        file.WriteLine(filepath);
                //    }                    
                //}  

            xlApp.DisplayAlerts = false;

            //if (xlWorkSheet.ProtectContents == false)
            //{
            //    xlWorkSheet.Protect("c0r36o");
            //}

            //xlWorkBook.Save();


            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
    }
}