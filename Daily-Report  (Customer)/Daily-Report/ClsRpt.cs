using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Reflection;
using System.IO;

namespace Daily_Report
{
    class ClsRpt
    {
        public static string openfilePath = "";
        public static void HeaderRpt1(string path, string Selectpath1, string Selectpath2,string txtStartDate,string batchNum)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets.Add(xlWorkBook.Sheets[1], Type.Missing, Type.Missing, Type.Missing);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
           string[] date = txtStartDate.Split('/');
           xlWorkSheet.Name = "Daily Reports (" + date[1].ToString() + "-" + date[0].ToString() + "-" + date[2].ToString() + ")";
            //Daily Reports (28-03-2018)
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;
            //xlApp.Windows.Application.ActiveWindow.DisplayGridlines = false;

            // read excel file 1
            Excel.Application xlApp1 = new Excel.Application();
            Excel.Workbook xlWorkbook1 = xlApp1.Workbooks.Open(Selectpath1);
            Excel._Worksheet xlWorksheet1 = xlWorkbook1.Sheets[1];
            Excel.Range xlRange1 = xlWorksheet1.UsedRange;

            int rowCount1 = xlRange1.Rows.Count;
            //int colCount1 = xlRange1.Columns.Count;

            // read excel file 2
            Excel.Application xlApp2 = new Excel.Application();
            Excel.Workbook xlWorkbook2 = xlApp2.Workbooks.Open(Selectpath2);
            Excel._Worksheet xlWorksheet2 = xlWorkbook2.Sheets[1];
            Excel.Range xlRange2 = xlWorksheet2.UsedRange;

            int rowCount2 = xlRange2.Rows.Count;
            //int colCount2 = xlRange2.Columns.Count;
            //Add Header
            String[] ABC = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA","AB" };
            String[] Header_Name = { "BU", "Touchpoint", "Customer List Batch Number", "Dummy ID", "Client Number", "Owner Name", "Gender", "Phone Number", "Product Name", "Name of Rider", "Agent Name", "Chanel", "Interview Date", "Interview None", "Call Outcome", "Q2 New Purchase tNPS", "Q3 New Purchase Verbatim", "Q4 Agent tNPS", "Q5 Agent Verbatim", "Q6 Doc and info requirement", "Q7 Unreasonable requirement Verbatim", "Q8_Additional info submission", "Q9_Reason for purchase verbatim", "Q10_Area for Improvement verbatim", "Q11_Permit to Follow Up", "Q12_Request មេនូឡាយហ្វ៏ to call back", "Daily Flag Report", "C01 - Ke Siyan / C02 -  Uk Phearom / C03 - Norng SreyNin" };
            
            for (int i = 0; i <= ABC.Length - 1; i++)
            {
                xlWorkSheet.Cells[1, i + 1] = ABC[i];
                xlWorkSheet.Cells[1, i + 1].HorizontalAlignment = 3;
                xlWorkSheet.Cells[1, i + 1].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                xlWorkSheet.Cells[3, i + 1] = Header_Name[i];
                xlWorkSheet.Cells[3, i + 1].HorizontalAlignment = 3;
                xlWorkSheet.Cells[3, i + 1].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[3, i + 1].WrapText = true;
                xlWorkSheet.Cells[3, i + 1].VerticalAlignment = 2;
            }

            xlWorkSheet.Range[xlWorkSheet.Cells[2, 1], xlWorkSheet.Cells[2, 3]].Merge();
            xlWorkSheet.Cells[2, 1] = "IndoChina to create";
            xlWorkSheet.Cells[2,1].HorizontalAlignment = 3;
            //xlWorkSheet.Cells[2, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
            xlWorkSheet.Cells[2, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(146,208,80));
            for (int i = 1; i <= 3; i++)
            {
                //xlWorkSheet.Cells[3, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                xlWorkSheet.Cells[2, i].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            }

            xlWorkSheet.Range[xlWorkSheet.Cells[2, 4], xlWorkSheet.Cells[2, 12]].Merge();
            xlWorkSheet.Cells[2, 4] = "From Customer Data Set";
            xlWorkSheet.Cells[2,4].HorizontalAlignment = 3;
            xlWorkSheet.Cells[2, 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
            for (int i = 4; i <= 12; i++)
            {
                xlWorkSheet.Cells[3, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                xlWorkSheet.Cells[2, i].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            }

            xlWorkSheet.Range[xlWorkSheet.Cells[2, 13], xlWorkSheet.Cells[2, 15]].Merge();
            xlWorkSheet.Cells[2, 13] = "Official use";
            xlWorkSheet.Cells[2, 13].HorizontalAlignment = 3;
            xlWorkSheet.Cells[2, 13].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
            for (int i = 13; i <= 15; i++)
            {
                xlWorkSheet.Cells[3, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
                xlWorkSheet.Cells[2, i].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            }

            xlWorkSheet.Range[xlWorkSheet.Cells[2, 16], xlWorkSheet.Cells[2, 25]].Merge();
            xlWorkSheet.Cells[2, 16] = "tNPS Survey (response and coded Oes)";
            xlWorkSheet.Cells[2, 16].HorizontalAlignment = 3;
            //xlWorkSheet.Cells[2, 15].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
            xlWorkSheet.Cells[2, 16].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(146, 208, 80));
            for (int i = 16; i <= 25; i++)
            {
                //xlWorkSheet.Cells[3, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                xlWorkSheet.Cells[2, i].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            }

            xlWorkSheet.Range[xlWorkSheet.Cells[2, 26], xlWorkSheet.Cells[2, 27]].Merge();
            xlWorkSheet.Cells[2, 26] = "Official use";
            xlWorkSheet.Cells[2, 26].HorizontalAlignment = 3;
            xlWorkSheet.Cells[2, 26].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
            for (int i = 26; i <= 27; i++)
            {
                //xlWorkSheet.Cells[3, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
                xlWorkSheet.Cells[2, i].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            }

            //xlWorkSheet.Range[xlWorkSheet.Cells[2, 21], xlWorkSheet.Cells[2, 23]].Merge();
            xlWorkSheet.Cells[2, 28] = "Interviewer";
            xlWorkSheet.Cells[2, 28].HorizontalAlignment = 3;
            xlWorkSheet.Cells[2, 28].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            xlWorkSheet.Cells[2, 28].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(146, 208, 80));
            //for (int i = 22; i <= 23; i++)
            //{
                xlWorkSheet.Cells[3, 27].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(217, 217, 217));
                xlWorkSheet.Cells[2, 27].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //}

            xlWorkSheet.Cells[4, 30] = "Red";
            xlWorkSheet.Cells[5, 30] = "Green";
            xlWorkSheet.Cells[6, 30] = "Black";
            xlWorkSheet.Cells[4, 31] = "Q12=No , Q2 code 0-4";
            xlWorkSheet.Cells[5, 31] = "Q12=No , Q2 code 9 or 10";
            xlWorkSheet.Cells[6, 31] = "Q12= Yes";

            xlWorkSheet.get_Range("AD4", "AD6").Cells.Font.Size = 11;
            xlWorkSheet.get_Range("AD4", "AD6").Cells.Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White);
            xlWorkSheet.get_Range("AD4", "AD4").Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);
            xlWorkSheet.get_Range("AD5", "AD5").Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(0, 176, 80));
            xlWorkSheet.get_Range("AD6", "AD6").Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
            xlWorkSheet.get_Range("AD4", "AD6").Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            //===========================================
            //Read Data
            int CountData = 0;
            List<int> IDcode = new List<int>();
            List<string> ListDate = new List<string>();
            for (int rowidex = rowCount2; rowidex <= rowCount2; rowidex--)
            {
                if (rowidex == 1) { break; }
                string datevalue = xlRange2.Cells[rowidex, 29].Value2.ToString().Trim();
                double ddate = double.Parse(datevalue);
                DateTime d = DateTime.FromOADate(ddate);
                string getdate = d.ToString("M/d/yyyy");
                if (txtStartDate.ToString() == getdate)
                {
                    CountData += 1;
                    IDcode.Add(rowidex);
                    ListDate.Add(getdate);
                }
            }
            //Column in DB
            //int[] colidex = { 27, 28, 36, 37, 39, 40, 42, 43, 44, 45, 46, 47, 49, 4 };
              int[] colidex = { 28, 29,31, 37, 38, 40, 41, 43, 44, 45, 46, 47, 48, 50, 4 };
            //Column in Daily-Report
            //int[] reportidex = { 4, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 21 }; for old report
              int[] reportidex = { 4, 13,14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 27 };
            int rowCnt = IDcode.Count - 1; 
            for (int rowidex = IDcode.Count - 1; rowidex <= IDcode.Count - 1; rowidex--)
            {
                if (rowidex < 0)
                { break; }

                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 1] = "KH";
                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 2] = "New Purchase & Agent";
                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 3] = batchNum;
                //xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 8] = "Completed";
                if (xlRange2.Cells[IDcode[rowidex], 32].Value2.ToString().Trim() == "99. Other (ផ្សេងៗ​ ទៀត)")
                {
                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 14] = xlRange2.Cells[IDcode[rowidex], 33].Value2.ToString().Trim();
                }
                else 
                {
                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 14] = xlRange2.Cells[IDcode[rowidex], 32].Value2.ToString().Trim();
                }
                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 1].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 2].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 3].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 14].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                string[] IDout_come = xlRange2.Cells[IDcode[rowidex], 32].Value2.ToString().Trim().Split('.');
                //Get Data
                for (int i = 0; i <= colidex.Length - 1; i++)
                {
                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 26].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                    if (i == 0)
                    {
                        xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                        for (int getidx = 1; getidx <= rowCount1; getidx++)
                        {
                            try
                            {
                                if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == xlRange1.Cells[getidx, 1].Value2.ToString().Trim())
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 1].NumberFormat = "@";
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 1] = xlRange1.Cells[getidx, 2].Value2.ToString().Trim();//Convert To Text
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 2] = xlRange1.Cells[getidx, 3].Value2.ToString().Trim();
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 3] = xlRange1.Cells[getidx, 4].Value2.ToString().Trim();
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 4].NumberFormat = "@";
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 4] = xlRange1.Cells[getidx, 5].Value2.ToString().Trim();//Convert To Text
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 5] = xlRange1.Cells[getidx, 6].Value2.ToString().Trim();
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 6] = xlRange1.Cells[getidx, 7].Value2.ToString().Trim();
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 7] = xlRange1.Cells[getidx, 8].Value2.ToString().Trim();
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 8] = xlRange1.Cells[getidx, 9].Value2.ToString().Trim();
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 1].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 2].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 3].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 4].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 5].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 6].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 7].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 8].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                                    break;
                                }
                            }
                            catch { } 
                        }
                    }
                    else if (i == 2 || i == 4)
                    {
                        if (xlRange2.Cells[IDcode[rowidex], colidex[2]] != null && xlRange2.Cells[IDcode[rowidex], colidex[2]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[2]].Value2.ToString().Trim() != "")
                        {
                            if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "10 : extremely likely")
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = "10";
                            }
                            else if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "0 : not at all likely")
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = "0";
                            }
                            else
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                            }
                            xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]].NumberFormat = "0";
                        }
                    }
                    else if (i == 6)
                    {
                        if (xlRange2.Cells[IDcode[rowidex], colidex[2]] != null && xlRange2.Cells[IDcode[rowidex], colidex[2]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[2]].Value2.ToString().Trim() != "")
                        {
                            if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "10 : Totally agree")
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = "10";
                            }
                            else if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "0 : Totally disagree")
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = "0";
                            }
                            else
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                            }
                            xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]].NumberFormat = "0";
                        }
                    }
                    else if (i == colidex.Length - 1)
                    {
                        //if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "C01")
                        //{
                        //    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = 1;
                        //}
                        //else if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "C02")
                        //{
                        //    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 1] = 2;
                        //}
                        //else
                        //{
                        //    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 2] = 3;
                        //}
                        xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                        xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]].HorizontalAlignment = 3;
                        //xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 1].HorizontalAlignment = 3;
                        //xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 2].HorizontalAlignment = 3;
                        xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                        //xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 1].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                        //xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 2].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                    }
                    else
                    {
                        if (xlRange2.Cells[IDcode[rowidex], colidex[2]] != null && xlRange2.Cells[IDcode[rowidex], colidex[2]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[2]].Value2.ToString().Trim() != "")
                        {
                            if (xlRange2.Cells[IDcode[rowidex], colidex[i]] != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() != "")
                            {
                                if (i == 1)
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = ListDate[rowidex];
                                }
                                else { xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim(); }
                            }
                            else
                            {
                                if (i == 3)
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i] + 1].Value2.ToString().Trim();
                                }
                                else if (i == 5)
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i] + 1].Value2.ToString().Trim();
                                }
                                else
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = "";
                                }
                            }
                        }
                        else 
                        {
                            if (i == 1)
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = ListDate[rowidex];
                            }
                        }
                    }
                }
            }
            xlWorkSheet.Range["B:B"].ColumnWidth = 21.00;
            xlWorkSheet.Range["E:L"].ColumnWidth = 20.00;
            //xlWorkSheet.Range["F:F"].ColumnWidth = 20.00;
            //xlWorkSheet.Range["G:G"].ColumnWidth = 15.00;
            //xlWorkSheet.Range["H:H"].ColumnWidth = 19.00;
            //xlWorkSheet.Range["I:I"].ColumnWidth = 20.00;
            //xlWorkSheet.Range["J:J"].ColumnWidth = 20.00;
            xlWorkSheet.Range["M:M"].ColumnWidth = 11.00;
            xlWorkSheet.Range["N:N"].ColumnWidth = 50.00;
            xlWorkSheet.Range["AA:AA"].ColumnWidth = 32.00;
            xlWorkSheet.Range["P:P"].Columns.AutoFit();
            xlWorkSheet.Range["R:R"].Columns.AutoFit();
            xlWorkSheet.Range["V:V"].Columns.AutoFit();
            xlWorkSheet.Range["W:W"].Columns.AutoFit();
                //cleanup
                GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange1);
            Marshal.ReleaseComObject(xlWorksheet1);
            Marshal.ReleaseComObject(xlRange2);
            Marshal.ReleaseComObject(xlWorksheet2);

            //close and release
            xlWorkbook1.Close();
            Marshal.ReleaseComObject(xlWorkbook1);
            xlWorkbook2.Close();
            Marshal.ReleaseComObject(xlWorkbook2);

            //quit and release
            xlApp1.Quit();
            Marshal.ReleaseComObject(xlApp1);
            xlApp2.Quit();
            Marshal.ReleaseComObject(xlApp2);
            try
            {
                //xlWorkBook.CheckCompatibility = false;
                xlApp.DisplayAlerts = false;
                //xlWorkBook.DoNotPromptForConvert = true;
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                xlWorkBook.SaveAs(path, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            openfilePath = path;
            //MessageBox.Show("Daily-Report has been successful!!!.");
        }
    }
}
