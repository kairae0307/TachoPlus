using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Net.Sockets;
using System.IO;
using System.IO.Ports;
using System.Net;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Data.SqlClient;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data.OleDb;
using Microsoft.Win32;
using System.Management;
using System.Reflection;
using System.Net.Mail;
using System.Net.Mime;
using System.Net.NetworkInformation;
using System.Globalization;
using System.Security.Cryptography;
using System.Security.AccessControl;

namespace TachoPlus
{
    
    public partial class Cashier_DetailForm : Form
    {
        Form1 form1;
        private PrintableListView.PrintableListView m_list;
        public string m_Cashier_ID = "";
        public struct Total
        {
            public double Money;
            public double Distance;
            public double RealIncome;

        }
        public Total total;
        public Cashier_DetailForm(Form1 f)
        {
            form1 = f;
            m_list = new PrintableListView.PrintableListView();
            InitializeComponent();
        }

      public void Cashier_read(string Select_CashierID)
      {
          if (listView1.Items.Count > 0)
              listView1.Items.Clear();

          listView1.View = View.Details;
          listView1.GridLines = true;                   //   리스트 뷰 라인생성
          listView1.FullRowSelect = true;               // 라인 선택 */
           string DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + form1.mdbfilename;

                    string queryRead = "select * from TblTacho";

                    OleDbConnection conn = new OleDbConnection(@DBstring);
                    conn.Open();
                    OleDbCommand commRead = new OleDbCommand(queryRead, conn);
                    OleDbDataReader srRead = commRead.ExecuteReader();


                    conn = new OleDbConnection(@DBstring);
                    conn.Open();
                     commRead = new OleDbCommand(queryRead, conn);
                     srRead = commRead.ExecuteReader();

                    double TotalMoney = 0;
                    double TotalRealMoney = 0;
                    double TotalDist = 0;


                    while (srRead.Read())
                    {

                      
                        string CashierID = "";
                        if (srRead.IsDBNull(20) == false)
                        {
                            CashierID = srRead.GetString(20);
                            CashierID = CashierID.Trim();
                            m_Cashier_ID = CashierID;
                            Select_CashierID = Select_CashierID.Trim();
                            if (Select_CashierID == CashierID)
                            {
                                string tooltip_str = srRead.GetInt32(0).ToString();

                                int isOpenedDBNum = 0;
                                DateTime stTime = new DateTime(1, 1, 1, 0, 0, 0);
                                DateTime srTime = new DateTime(1, 1, 1, 0, 0, 0);
                                TimeSpan stSpan = new TimeSpan(0, 0, 0, 0);

                                uint tMoney = 0, tB = 0, tDB = 0, tDA = 0, tAB = 0, tAA = 0;
                                double tDistS = 0, tDistD = 0;
                                TimeSpan tOverD = new TimeSpan();

                                bool bIsMatched = false;

                                int nReadDBCnt = 0;         // 읽어들인 데이터 갯수 파악용

                                ListViewItem a = new ListViewItem(tooltip_str);       // ID


                                //  a.ToolTipText = " 영업 상세정보 \n ID 더블클릭!";

                                string strCarNo = srRead.GetString(1);  // 차량번호



                                a.SubItems.Add(strCarNo);
                                string strDriverNo = "";
                                if (srRead.IsDBNull(2) == false)                                            // 기사번호
                                {
                                    strDriverNo = srRead.GetString(2);
                                    a.SubItems.Add(strDriverNo);

                                }
                                else
                                {
                                    a.SubItems.Add(strDriverNo);
                                }



                                string strDriverName = "";
                                if (srRead.IsDBNull(3) == false)                                            // 기사이름
                                {
                                    strDriverName = srRead.GetString(3);
                                    a.SubItems.Add(strDriverName);

                                }
                                else
                                {
                                    a.SubItems.Add(strDriverName);
                                }

                                //a.SubItems.Add(srRead.GetString(2));                                

                                // string strDriverNo = srRead.GetString(2);


                                //  a.SubItems.Add(strDriverNo);                                           



                                a.SubItems.Add(srRead.GetDateTime(4).ToString("yyyy-MM-dd"));                       // 출고 날짜 
                                a.SubItems.Add(srRead.GetDateTime(4).ToString(" HH:mm:ss"));                       // 출고 시간  

                                //    a.SubItems.Add(srRead.GetDateTime(5).ToString("yyyy-MM-dd tt HH:mm:ss"));

                                a.SubItems.Add(srRead.GetDateTime(5).ToString("yyyy-MM-dd"));                       // 입고 날짜 
                                a.SubItems.Add(srRead.GetDateTime(5).ToString(" HH:mm:ss"));                       // 입고 시간  

                                if (srRead.IsDBNull(19) == false)                                            // 기사번호
                                {
                                    a.SubItems.Add(srRead.GetDateTime(19).ToString("yyyy-MM-dd"));                       // 입고 날짜 
                                    a.SubItems.Add(srRead.GetDateTime(19).ToString(" HH:mm"));                       // 입고 시간  
                                }
                                else
                                {
                                    a.SubItems.Add("");
                                    a.SubItems.Add("");
                                }

                                double dotsatus = srRead.GetDouble(10);



                                double uuu = (double)srRead.GetDouble(6);

                                string money;
                                if (dotsatus != 4 && dotsatus != 2 && dotsatus != 8)
                                {
                                    TotalMoney += uuu;
                                    int hhh = (int)uuu;
                                    money = string.Format("{0:D}", (int)hhh);
                                    a.SubItems.Add(money);
                                }
                                else
                                {
                                    if (dotsatus == 8)
                                    {
                                        TotalMoney += uuu;
                                        money = string.Format("{0:F3}", uuu);
                                        a.SubItems.Add(money);
                                    }
                                    else
                                    {
                                        TotalMoney += uuu;
                                        money = string.Format("{0:F2}", uuu);
                                        a.SubItems.Add(money);
                                    }
                                }// 미터수입
                                double ddd = 0;
                                if (srRead.IsDBNull(9) == false)                                            // 기사번호
                                {

                                    ddd = srRead.GetDouble(9); // 영업거리 
                                }


                                TotalDist += ddd;
                                string dist = string.Format("{0:N3} Km", ddd);
                                a.SubItems.Add(dist);


                                double GtandTotlaMoney = 0;
                                money = "0";
                                string GandMoenyStr = "0";
                                if (srRead.IsDBNull(7) == false)
                                {

                                    GtandTotlaMoney = srRead.GetDouble(7);// preview
                                    GtandTotlaMoney += srRead.GetDouble(6); ;  // Grnad Total

                                    if (dotsatus != 4 && dotsatus != 2 && dotsatus != 8)
                                    {
                                        money = string.Format("{0:D}", (int)srRead.GetDouble(7));  // preview total
                                        GandMoenyStr = string.Format("{0:D}", (int)GtandTotlaMoney);
                                    }
                                    else
                                    {
                                        if (dotsatus == 8)
                                        {
                                            money = string.Format("{0:F3}", srRead.GetDouble(7));
                                            GandMoenyStr = string.Format("{0:F3}", GtandTotlaMoney);
                                        }
                                        else
                                        {
                                            money = string.Format("{0:F2}", srRead.GetDouble(7));
                                            GandMoenyStr = string.Format("{0:F2}", GtandTotlaMoney);
                                        }
                                    }



                                    a.SubItems.Add(GandMoenyStr);
                                    a.SubItems.Add(money);
                                }
                                else
                                {
                                    a.SubItems.Add(GandMoenyStr);
                                    a.SubItems.Add(money);
                                }

                                if (srRead.IsDBNull(11) == false)  // real money
                                {

                                      uuu = (double)srRead.GetDouble(11);
                                    if (dotsatus != 4 && dotsatus != 2 && dotsatus != 8)
                                    {
                                        money = string.Format("{0:D}", (int)uuu);  // 
                                        TotalRealMoney += uuu;
                                    }
                                    else
                                    {
                                        if (dotsatus == 8)
                                        {
                                            money = string.Format("{0:F3}", uuu);
                                            TotalRealMoney += uuu;

                                        }
                                        else
                                        {
                                            money = string.Format("{0:F2}", uuu);
                                            TotalRealMoney += uuu;

                                        }
                                    }

                                    a.SubItems.Add(money);

                                }
                                else
                                {
                                    a.SubItems.Add("");

                                }

                                if (srRead.IsDBNull(20) == false)
                                {
                                    string Cashier = srRead.GetString(20);
                                    a.SubItems.Add(Cashier);

                                }
                                else
                                {
                                    a.SubItems.Add("");

                                }

                              

                                listView1.Items.Add(a);
                            }
                           
                          
                        }
                       
                         
                    }
                    total.Money = TotalMoney;
                    total.Distance = TotalDist;
                    total.RealIncome += TotalRealMoney;

                    int nTotalItems = listView1.Items.Count;

                    if (nTotalItems != 0)
                    {
                        ListViewItem b = new ListViewItem("SUM");
                        b.SubItems.Add("");                             ///   공백 다섯칸 만들기
                        b.SubItems.Add("");
                        b.SubItems.Add("");
                        b.SubItems.Add("");
                        b.SubItems.Add("");
                        b.SubItems.Add("");
                        b.SubItems.Add("");
                        b.SubItems.Add("");
                        b.SubItems.Add("");
                        int valtemp = (int)total.Money;
                        double vlatemp1 = (double)valtemp;

                        string temp = "";
                        if ((total.Money % vlatemp1) != 0)
                        {

                            temp = string.Format("{0:F2}", total.Money);
                            b.SubItems.Add(temp);
                        }
                        else
                        {
                            temp = string.Format("{0:D}", (int)total.Money);
                            b.SubItems.Add(temp);
                        }

                        temp = string.Format("{0:N3} Km", total.Distance);
                        b.SubItems.Add(temp);
                        b.BackColor = System.Drawing.Color.LightGray;

                        b.SubItems.Add("");
                        b.SubItems.Add("");
                      //  b.SubItems.Add("");
                    //    b.SubItems.Add("");  real

                         valtemp = (int)total.RealIncome;
                         vlatemp1 = (double)valtemp;

                         temp = "";
                        if ((total.Money % vlatemp1) != 0)
                        {

                            temp = string.Format("{0:F2}", total.RealIncome);
                            b.SubItems.Add(temp);
                        }
                        else
                        {
                            temp = string.Format("{0:D}", (int)total.RealIncome);
                            b.SubItems.Add(temp);
                        }


                        b.SubItems.Add("");
                        listView1.Items.Add(b);
                    }
                    FillList(this.m_list, listView1);
      
      }
      private void FillList(ListView list, ListView table)
      {

          list.SuspendLayout();

          // Clear list
          list.Items.Clear();
          list.Columns.Clear();

          // Columns
          int nCol = 0;


          int a = 0;

          int colcnt = 16;

       


          try
          {
              for (int i = 0; i < colcnt; i++)
              {


                  ColumnHeader[] col = new ColumnHeader[16];
                  ColumnHeader ch = new ColumnHeader();

                  col[i] = table.Columns[i];
                  ch.Text = col[i].Text;
                  ch.TextAlign = HorizontalAlignment.Right;
                  switch (nCol)
                  {
                      case 0: ch.Width = 40; break;       // id   


                      case 1:
                          ch.TextAlign = HorizontalAlignment.Left;    //trip no.
                          ch.Width = 90;
                          break;
                      case 2: ch.Width = 170; break;              // 입고
                      case 3: ch.Width = 120; break;               // 미터 수입

                      case 4: ch.Width = 120; break;               // 영업거리

                      case 5: ch.Width = 120; break;               // 

                      case 6: ch.Width = 120; break;
                      case 7: ch.Width = 120; break;               // speed
                      case 8: ch.Width = 120; break;              // sensor
                      case 9: ch.Width = 120; break;
                      case 10: ch.Width = 120; break;
                      case 11: ch.Width = 120; break;
                      case 12: ch.Width = 120; break;
                      case 13: ch.Width = 120; break;
                      case 14: ch.Width = 120; break;
                      case 15: ch.Width = 120; break;
                      case 16: ch.Width = 120; break;
                      case 17: ch.Width = 120; break;
                      case 18: ch.Width = 120; break;
                      case 19: ch.Width = 120; break;
                      case 20: ch.Width = 120; break;
                      case 21: ch.Width = 120; break;
                      case 22: ch.Width = 120; break;





                      //    case 11:
                      //   case 12:
                      //   case 13:
                      //   case 14:
                      //   case 15: ch.Width = 40; break;
                      //	case 16:
                      //	case 17: 
                      default:
                          ch.Width = 0;

                          break;
                  }
                  list.Columns.Add(ch);
                  nCol++;
              }

              // Rows
              for (int n = 0; n < table.Items.Count; n++)
              {
                  ListViewItem item = new ListViewItem();
                  //item.Text = row[0].ToString();
                  item.Text = table.Items[n].Text;

                  for (int i = 1; i < colcnt; i++)
                  {
                      item.SubItems.Add(table.Items[n].SubItems[i].Text);
                  }
                  list.Height = 100;
                  list.Items.Add(item);
              }



              list.ResumeLayout();
          }
          catch (Exception ex)
          {
              MessageBox.Show(ex.Message);
          }
      }
      private void buttonPrint_Click(object sender, EventArgs e)
      {
          m_list.Title = "CASHIER REPORT\n";

          m_list.PageSetup();
        

          m_list.FitToPage = true;
          m_list.PrintPreview();
      }

      private void button7_Click(object sender, EventArgs e)
      {
          Excel_print();
      }
      public void Excel_print()
      {
          // string filePath = "c:\\test.xlsx";
          Excel.ApplicationClass excel = new Excel.ApplicationClass();

          int colIndex = 0;
          int rowIndex = 5;
          Excel.Application excelApp = null;

          object missingType = Type.Missing;

          excelApp = new Microsoft.Office.Interop.Excel.Application();
          //  excel.Application.Workbooks.Add(true);
          Excel.Range oRng = null;
          Excel.Workbook excelBook = excel.Workbooks.Add(missingType);
          Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.Add(missingType, missingType, missingType, missingType);

          ((Excel.Range)excelWorksheet.get_Range("A:A", Missing.Value)).ColumnWidth = 8.33;
          ((Excel.Range)excelWorksheet.get_Range("B:B", Missing.Value)).ColumnWidth = 16.33;
          ((Excel.Range)excelWorksheet.get_Range("C:C", Missing.Value)).ColumnWidth = 21.33;
          ((Excel.Range)excelWorksheet.get_Range("D:D", Missing.Value)).ColumnWidth = 14.33;
          ((Excel.Range)excelWorksheet.get_Range("E:E", Missing.Value)).ColumnWidth = 14.33;
          ((Excel.Range)excelWorksheet.get_Range("F:F", Missing.Value)).ColumnWidth = 14.33;
          ((Excel.Range)excelWorksheet.get_Range("G:G", Missing.Value)).ColumnWidth = 14.33;
          ((Excel.Range)excelWorksheet.get_Range("H:H", Missing.Value)).ColumnWidth = 14.33;
          ((Excel.Range)excelWorksheet.get_Range("I:I", Missing.Value)).ColumnWidth = 16.33;
          ((Excel.Range)excelWorksheet.get_Range("J:J", Missing.Value)).ColumnWidth = 16.33;
          ((Excel.Range)excelWorksheet.get_Range("K:K", Missing.Value)).ColumnWidth = 16.33;
          ((Excel.Range)excelWorksheet.get_Range("L:L", Missing.Value)).ColumnWidth = 16.33;
          ((Excel.Range)excelWorksheet.get_Range("M:M", Missing.Value)).ColumnWidth = 10.33;
          ((Excel.Range)excelWorksheet.get_Range("N:N", Missing.Value)).ColumnWidth = 10.33;
          ((Excel.Range)excelWorksheet.get_Range("O:O", Missing.Value)).ColumnWidth = 10.33;
          ((Excel.Range)excelWorksheet.get_Range("P:P", Missing.Value)).ColumnWidth = 10.33;
       

          // BackColor 설정
          ((Excel.Range)excelWorksheet.get_Range("A5", "A5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
          ((Excel.Range)excelWorksheet.get_Range("B5", "B5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
          ((Excel.Range)excelWorksheet.get_Range("C5", "C5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
          ((Excel.Range)excelWorksheet.get_Range("D5", "D5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
          ((Excel.Range)excelWorksheet.get_Range("E5", "E5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
          ((Excel.Range)excelWorksheet.get_Range("F5", "F5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
          ((Excel.Range)excelWorksheet.get_Range("G5", "G5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
          ((Excel.Range)excelWorksheet.get_Range("H5", "H5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
          ((Excel.Range)excelWorksheet.get_Range("I5", "I5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
          ((Excel.Range)excelWorksheet.get_Range("J5", "J5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
          ((Excel.Range)excelWorksheet.get_Range("K5", "K5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
          ((Excel.Range)excelWorksheet.get_Range("L5", "L5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
          ((Excel.Range)excelWorksheet.get_Range("M5", "M5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
          ((Excel.Range)excelWorksheet.get_Range("N5", "N5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
          ((Excel.Range)excelWorksheet.get_Range("O5", "O5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
          ((Excel.Range)excelWorksheet.get_Range("P5", "P5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
        



          ((Excel.Range)excelWorksheet.get_Range("A:A, B:B, C:C, D:D, E:E,F:F,G:G,H:H,I:I,J:J,K:K,L:L,M:M,N:N,O:O,P:P", Missing.Value)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

          oRng = excel.get_Range("A2", "J2"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
        
       //   oRng.Value2 = label2.Text + "     " + label3.Text + "     " + label1.Text;  //문구 삽입
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //  oRng.Font.Color = 0xfffffff;    //폰트 컬러




          oRng = excel.get_Range("A5", "A5"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //  oRng.Font.Color = -16776961;    //폰트 컬러
          //테두리
          oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;

          ///////////////////////////////////////////////////////

          oRng = excel.get_Range("B5", "B5"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //    oRng.Font.Color = -16776961;    //폰트 컬러
          //테두리
          oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
          ///////////////////////////////////////////////////////

          oRng = excel.get_Range("C5", "C5"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //    oRng.Font.Color = -16776961;    //폰트 컬러
          //테두리
          oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
          ///////////////////////////////////////////////////////

          oRng = excel.get_Range("D5", "D5"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //    oRng.Font.Color = -16776961;    //폰트 컬러
          //테두리
          oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
          ///////////////////////////////////////////////////////

          oRng = excel.get_Range("E5", "E5"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //    oRng.Font.Color = -16776961;    //폰트 컬러
          //테두리
          oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
          ///////////////////////////////////////////////////////

          oRng = excel.get_Range("F5", "F5"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //    oRng.Font.Color = -16776961;    //폰트 컬러
          //테두리
          oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
          ///////////////////////////////////////////////////////

          oRng = excel.get_Range("G5", "G5"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //    oRng.Font.Color = -16776961;    //폰트 컬러
          //테두리
          oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
          ///////////////////////////////////////////////////////

          oRng = excel.get_Range("H5", "H5"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //    oRng.Font.Color = -16776961;    //폰트 컬러
          //테두리
          oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
          ///////////////////////////////////////////////////////

          oRng = excel.get_Range("I5", "I5"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //    oRng.Font.Color = -16776961;    //폰트 컬러
          //테두리
          oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


          oRng = excel.get_Range("J5", "J5"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //    oRng.Font.Color = -16776961;    //폰트 컬러
          //테두리
          oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


          oRng = excel.get_Range("K5", "K5"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //    oRng.Font.Color = -16776961;    //폰트 컬러
          //테두리
          oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


          oRng = excel.get_Range("L5", "L5"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //    oRng.Font.Color = -16776961;    //폰트 컬러
          //테두리
          oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


          oRng = excel.get_Range("M5", "M5"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //    oRng.Font.Color = -16776961;    //폰트 컬러
          //테두리
          oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


          oRng = excel.get_Range("N5", "N5"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //    oRng.Font.Color = -16776961;    //폰트 컬러
          //테두리
          oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


          oRng = excel.get_Range("O5", "O5"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //    oRng.Font.Color = -16776961;    //폰트 컬러
          //테두리
          oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


          oRng = excel.get_Range("P5", "P5"); //해당 범위의 셀 획득
          oRng.MergeCells = true; //머지
          oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
          oRng.Font.Bold = true;  //볼드
          //    oRng.Font.Color = -16776961;    //폰트 컬러
          //테두리
          oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
          oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


      


          for (int i = 0; i < this.listView1.Columns.Count; i++)
          {
              colIndex++;
              excel.Cells[5, colIndex] = this.listView1.Columns[i].Text;


          }

          for (int i = 0; i < this.listView1.Items.Count; i++)
          {
              rowIndex++;
              colIndex = 0;
              for (int j = 0; j < this.listView1.Items[i].SubItems.Count; j++)
              {
                  colIndex++;
                  excel.Cells[rowIndex, colIndex] = this.listView1.Items[i].SubItems[j].Text;
                  //     excel.Cells.AutoOutline();

              }
          }

          string Acell = "A";
          string Ncell = "P";
          Acell += (this.listView1.Items.Count + 5).ToString();
          Ncell += (this.listView1.Items.Count + 5).ToString();
          ((Excel.Range)excelWorksheet.get_Range(Acell, Ncell)).Interior.Color = ColorTranslator.ToOle(Color.LightGray);

          System.IO.Directory.CreateDirectory("c:\\Tacho2\\EXCEL");

          string canum = this.listView1.Items[0].SubItems[1].Text;
          //   string FileDate = (this.listView1.Items[0].SubItems[5].Text).Replace("-","") + "_" + (this.listView1.Items[0].SubItems[6].Text).Replace(":","");
          string FileDate = (this.listView1.Items[0].SubItems[5].Text);


          DateTime dt = Convert.ToDateTime(FileDate);

          FileDate = dt.Day.ToString() + "-" + dt.Month.ToString() + "-" + dt.Year.ToString();

          canum = canum.Trim();

          string savefile = "c:\\Tacho2\\EXCEL\\" +"Cashier"+ m_Cashier_ID + "-" + FileDate;

          // excelBook.SaveAs(savefile, Excel.XlFileFormat.xlExcel7);
          if (System.IO.File.Exists(savefile + ".xlsx"))  // information 같은 파일의이름이 존재 함
          {

              int num = 1;
              bool check = true;
              savefile += "_" + num.ToString();


              do
              {

                  if (!System.IO.File.Exists(savefile + ".xlsx"))  // information 같은 파일의이름이 존재 함
                  {
                      check = false;
                  }
                  num++;
                  if (check == true)
                  {
                      savefile = savefile.Remove(savefile.Length - 1);
                      savefile += num.ToString();
                  }


              } while (check);


          }
          excel.Visible = true;
          try
          {

              excelBook.SaveAs(savefile + ".xlsx");
          }
          catch (Exception ex)
          {

          }

          //excel.Save(filePath);
      }
    }
}
