using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Drawing.Printing;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace TachoPlus
{
    public partial class SearchdataForm : Form
    {
        Form1 form1;
        public Total total;
        public PrintableListView.PrintableListView m_list;
        public struct Total
        {
            public double Money;
            public double Distance;
            public double RealIncome;

        }
        public SearchdataForm(Form1 f)
        {
            form1 = f;
            InitializeComponent();
            ImageList dummyImageList = new ImageList();
            dummyImageList.ImageSize = new System.Drawing.Size(1, 18);
            listView1.SmallImageList = dummyImageList;
            listView1.View = View.Details;
            listView1.GridLines = true;
            listView1.FullRowSelect = true;
            m_list = new PrintableListView.PrintableListView();
        }
     
        public void SearchData(string CarNO_or_DirverID, DateTime Starttime, DateTime Endtime,int id)
        {

            // id ==0 차량번호 
            //      1 기사번호
            //      2 all
            //      3 cashier

            Starttime = new DateTime(Starttime.Year, Starttime.Month, Starttime.Day, 0, 0, 0);
            Endtime = new DateTime(Endtime.Year, Endtime.Month, Endtime.Day, 0, 0, 0);
            string Dirname = "";
            int nCnt = 1;
            total.Money = 0;
            total.Distance = 0;

            string year = Starttime.Year.ToString();
            string StartMonth =  Starttime.Month.ToString();
            string StartDay = Starttime.Day.ToString();

             string EndMonth =  Endtime.Month.ToString();
             string EndDay = Endtime.Day.ToString();
             DateTime mdbTime = new DateTime();


             if (form1.ViewerMode == true)
             {
                 Dirname = @"\\" + form1.ShareIP + "\\tacho2\\" + "\\TACHO\\" + year;  // 해당 연도 폴더를 지정한다. 
             }
             else
             {
                 Dirname = form1.TACHO2_path + "\\TACHO\\" + year;  // 해당 연도 폴더를 지정한다. 
             }
         //   Dirname = form1.TACHO2_path + "\\TACHO\\" + year;  // 해당 연도 폴더를 지정한다. 


            DirectoryInfo dirs = new DirectoryInfo(Dirname);
            DirectoryInfo[] DIRS = dirs.GetDirectories();
            int IdCount = 0;
            for(int i=0; i<DIRS.Length; i++)
            {

                int DirMonth = Int32.Parse(DIRS[i].ToString());


                if (Starttime.Month > DirMonth || Endtime.Month < DirMonth)
                {
                    continue;
                }
               
                    if (form1.ViewerMode == true)
                    {
                        Dirname = @"\\" + form1.ShareIP + "\\tacho2\\" + "\\TACHO\\" + year + "\\" + DIRS[i].ToString();  // 해당 연도 폴더를 지정한다. 
                    }
                    else
                    {
                        Dirname = form1.TACHO2_path + "\\TACHO\\" + year + "\\" + DIRS[i].ToString();  // Month 폴더를 지정한다.
                    }
               //     Dirname = form1.TACHO2_path + "\\TACHO\\" + year + "\\" + StartMonth;  // Month 폴더를 지정한다.
                    DirectoryInfo mdbPath = new DirectoryInfo(Dirname);
                    FileInfo[] files = mdbPath.GetFiles();   // 파일 가져온다.
                    string[] file_str = new string[files.Length];
                    char[] trimChars = { '.', 'm', 'd', 'b' };
                    int cnt = 0;
                    for (int a = 0; a < files.Length; a++)   // mdb 파일 목록을 가져온다.
                    {

                        if (files[a].Extension != ".ldb")
                        {

                            file_str[a] = files[a].ToString();
                            file_str[a] = file_str[a].TrimEnd(trimChars);

                        }
                        string DBstring = "";
                        if (file_str[a] == null)
                        {
                            continue;
                        }
                        DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + Dirname + "\\" + file_str[a] + ".mdb";

                        string queryRead = "select * from TblTacho";

                        OleDbConnection conn = new OleDbConnection(@DBstring);
                        conn.Open();
                        OleDbCommand commRead = new OleDbCommand(queryRead, conn);
                        OleDbDataReader srRead = commRead.ExecuteReader();


                        double TotalMoney = 0;
                        double TotalDist = 0;
                        double Total_RealMoney = 0;
                     

                        while (srRead.Read())
                        {
                            
                            string tooltip_str = srRead.GetInt32(0).ToString();

                        
                            DateTime stTime = new DateTime(1, 1, 1, 0, 0, 0);
                            DateTime srTime = new DateTime(1, 1, 1, 0, 0, 0);
                            TimeSpan stSpan = new TimeSpan(0, 0, 0, 0);





                            string strCarNo = srRead.GetString(1);  // 차량번호

                            strCarNo = strCarNo.TrimEnd();
                            mdbTime = srRead.GetDateTime(4); // begin Time
                            string strDriverNo = "";
                            string CashierID = "";
                           if (srRead.IsDBNull(2) == false)
                            {
                                strDriverNo = srRead.GetString(2);
                                strDriverNo = strDriverNo.Trim();
                                CarNO_or_DirverID = CarNO_or_DirverID.Trim();
                              
                            }

                           if (srRead.IsDBNull(20) == false)
                           {
                               CashierID = srRead.GetString(20);
                               CashierID = CashierID.Trim();
                               CarNO_or_DirverID = CarNO_or_DirverID.Trim(); //  검색한  cashier ID
                           }
                           
                            mdbTime = new DateTime(mdbTime.Year, mdbTime.Month, mdbTime.Day, 0, 0, 0);

                         
                            if (mdbTime < Starttime || mdbTime > Endtime)
                                continue;

                          
                                if (id == 0)
                                {
                                    if (strCarNo != CarNO_or_DirverID) continue;  // 차량 번호 걸러내기 
                                }
                                else if (id == 1)
                                {
                                    if (strDriverNo != CarNO_or_DirverID) continue; // 기사 번호 걸러내기 
                                }
                                else if (id == 3)
                                {
                                    if (CashierID != CarNO_or_DirverID) continue; // cashier 번호 걸러내기 
                                }
                         


                            IdCount++;
                            ListViewItem list = new ListViewItem(IdCount.ToString());       // ID
                            //  a.ToolTipText = " 영업 상세정보 \n ID 더블클릭!";

                           
                            list.SubItems.Add(strCarNo);                                               // 차량번호
                             strDriverNo = "";
                            if (srRead.IsDBNull(2) == false)
                            {
                                strDriverNo = srRead.GetString(2);
                                list.SubItems.Add(strDriverNo);

                            }
                            else
                            {
                                list.SubItems.Add(strDriverNo);
                            }

                            //a.SubItems.Add(srRead.GetString(2));                                  // 기사번호

                            // string strDriverNo = srRead.GetString(2);


                            //  a.SubItems.Add(strDriverNo);                                            // 기사번호

                            string strDriverName = "";


                            list.SubItems.Add(srRead.GetDateTime(4).ToString("yyyy-MM-dd tt HH:mm:ss"));                       // 출고시간
                            list.SubItems.Add(srRead.GetDateTime(5).ToString("yyyy-MM-dd tt HH:mm:ss"));


                            double dotsatus = srRead.GetDouble(10);



                            double uuu = (double)srRead.GetDouble(6);

                            string money;
                            if (dotsatus != 4 && dotsatus != 2 && dotsatus !=8)
                            {
                                TotalMoney += uuu;
                                int hhh = (int)uuu;
                                money = string.Format("{0}", (int)hhh);
                                list.SubItems.Add(money);
                            }
                            else
                            {
                                if (dotsatus == 8)
                                {
                                    TotalMoney += uuu;
                                    money = string.Format("{0:F3}", uuu);
                                    list.SubItems.Add(money);
                                }
                                else
                                {
                                    TotalMoney += uuu;
                                    money = string.Format("{0:F2}", uuu);
                                    list.SubItems.Add(money);
                                }
                            }// 미터수입



                            double ddd = srRead.GetDouble(8);



                            TotalDist += ddd;
                            string dist = string.Format("{0:N3} Km", ddd);
                            list.SubItems.Add(dist);


                            double GtandTotlaMoney = 0;
                            money = "0";
                            string GandMoenyStr = "0";
                            if (srRead.IsDBNull(7) == false)
                            {

                                GtandTotlaMoney = srRead.GetDouble(7);// preview
                                GtandTotlaMoney += srRead.GetDouble(6); ;  // Grnad Total

                                if (dotsatus != 4 && dotsatus != 2 && dotsatus != 8)
                                {
                                    money = string.Format("{0}", srRead.GetDouble(7));  // preview total
                                    GandMoenyStr = string.Format("{0}", GtandTotlaMoney);
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

                             



                                list.SubItems.Add(GandMoenyStr);
                                list.SubItems.Add(money);
                            }
                            else
                            {
                                list.SubItems.Add(GandMoenyStr);
                                list.SubItems.Add(money);
                            }


                            if (srRead.IsDBNull(11) == false)
                            {
                                uuu = (double)srRead.GetDouble(11);


                                if (dotsatus != 4 && dotsatus != 2 && dotsatus != 8)
                                {
                                    Total_RealMoney += uuu;
                                    int hhh = (int)uuu;
                                    money = string.Format("{0}", (int)hhh);
                                    list.SubItems.Add(money);
                                }
                                else
                                {
                                    if (dotsatus == 8)
                                    {
                                        Total_RealMoney += uuu;
                                        money = string.Format("{0:F3}", uuu);
                                        list.SubItems.Add(money);
                                    }
                                    else
                                    {
                                        Total_RealMoney += uuu;
                                        money = string.Format("{0:F2}", uuu);
                                        list.SubItems.Add(money);
                                    }
                                }// 미터수입
                            }
                            else
                            {
                                int hhh =0;
                                money = string.Format("{0}", (int)hhh);
                                list.SubItems.Add(money);
                            }

                            list.SubItems.Add(CashierID);
                            listView1.Items.Add(list);
                        }
                        total.Money += TotalMoney;
                        total.Distance += TotalDist;
                        total.RealIncome += Total_RealMoney;

                        conn.Close();
                    }
                  
              
             
             
            }
            int nTotalItems = listView1.Items.Count;

            if (nTotalItems != 0)
            {
                ListViewItem b = new ListViewItem("SUM");
                b.SubItems.Add("");                             ///   공백 다섯칸 만들기
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
                    temp = string.Format("{0}", total.Money);
                    b.SubItems.Add(temp);
                }

                temp = string.Format("{0:N3} Km", total.Distance);
                b.SubItems.Add(temp);
                b.BackColor = System.Drawing.Color.LightGray;

                b.SubItems.Add("");
                b.SubItems.Add("");
                 valtemp = (int)total.RealIncome;
                 vlatemp1 = (double)valtemp;

                 temp = "";
                if ((total.RealIncome % vlatemp1) != 0)
                {


                    temp = string.Format("{0:F2}", total.RealIncome);
                    b.SubItems.Add(temp);
                }
                else
                {
                    temp = string.Format("{0}", total.RealIncome);
                    b.SubItems.Add(temp);
                }
              //  b.SubItems.Add("");
                b.SubItems.Add("");
                listView1.Items.Add(b);
            }


            FillList(m_list, listView1);
         
        }
        public void PrintPreview_FormData()
        {
            m_list.Title = "TACHO REPORT\n";


            //m_list.FitToPage = m_cbFitToPage.Checked;

            //	m_list.PageSetup();

            m_list.FitToPage = true;

            m_list.PrintPreview();
        }
        private void buttonPrint_Click(object sender, EventArgs e)
        {
            PrintPreview_FormData();
        }
        public void FillList(ListView list, ListView table)
        {
            list.SuspendLayout();

            // Clear list
            list.Items.Clear();
            list.Columns.Clear();

            // Columns
            int nCol = 0;


            int a = 0;


            try
            {
                for (int i = 0; i < 11; i++)
                {


                    ColumnHeader[] col = new ColumnHeader[11];
                    ColumnHeader ch = new ColumnHeader();

                    col[i] = table.Columns[i];
                    ch.Text = col[i].Text;
                    ch.TextAlign = HorizontalAlignment.Right;
                    switch (nCol)
                    {
                        case 0: ch.Width = 40; break;       // id   

                        case 1: ch.Width = 80; break;       // 차량번호

                        case 2: ch.Width = 80; break;       // 기사 번호

                        case 3:
                            ch.TextAlign = HorizontalAlignment.Left;    // 출고 
                            ch.Width = 170;
                            break;
                        case 4: ch.Width = 170; break;              // 입고
                        case 5: ch.Width = 120; break;               // 미터 수입

                        case 6: ch.Width = 120; break;               // 영업거리

                        case 7: ch.Width = 120; break;               // 

                        case 8: ch.Width = 120; break;

                        case 9: ch.Width = 120; break;
                        case 10: ch.Width = 120; break;


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

                    for (int i = 1; i < table.Columns.Count; i++)
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

        private void button7_Click(object sender, EventArgs e)
        {
            Excel_print();
        }
        public void Excel_print()
        {

            string filePath = "c:\\Tacho_Report.xlsx";
            Excel.ApplicationClass excel = new Excel.ApplicationClass();

            int colIndex = 0;
            int rowIndex = 3;
            Excel.Application excelApp = null;
            excel.UserName = form1.formname + ".xlsx";


            object missingType = Type.Missing;
            object fileName = form1.formname + ".xlsx";


            excelApp = new Microsoft.Office.Interop.Excel.Application();
            //  excel.Application.Workbooks.Add(true);
            Excel.Range oRng = null;
            Excel.Workbook excelBook = excel.Workbooks.Add(missingType);



            //   excelBook = excelApp.Workbooks.Open((string)fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //    Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing);



            //    excelBook.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //  Excel.XlSaveAsAccessMode.xlShared, Excel.XlSaveConflictResolution.xlLocalSessionChanges,
            //  Type.Missing, Type.Missing, Type.Missing, Type.Missing);


            Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.Add(missingType, missingType, missingType, missingType);
            excelWorksheet.Name = form1.formname;
            ((Excel.Range)excelWorksheet.get_Range("A:A", Missing.Value)).ColumnWidth = 8.33;
            ((Excel.Range)excelWorksheet.get_Range("B:B", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("C:C", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("D:D", Missing.Value)).ColumnWidth = 25.33;
            ((Excel.Range)excelWorksheet.get_Range("E:E", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("F:F", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("G:G", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("H:H", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("I:I", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("J:J", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("K:K", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("L:L", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("M:M", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("N:N", Missing.Value)).ColumnWidth = 18.33;
            ((Excel.Range)excelWorksheet.get_Range("O:O", Missing.Value)).ColumnWidth = 18.33;
            ((Excel.Range)excelWorksheet.get_Range("P:P", Missing.Value)).ColumnWidth = 14.33;


            // BackColor 설정
            ((Excel.Range)excelWorksheet.get_Range("A3", "A3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("B3", "B3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("C3", "C3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("D3", "D3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("E3", "E3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("F3", "F3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("G3", "G3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("H3", "H3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("I3", "I3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("J3", "J3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("K3", "K3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("L3", "L3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("M3", "M3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("N3", "N3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("O3", "O3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("P3", "P3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);



            ((Excel.Range)excelWorksheet.get_Range("A:A, B:B, C:C, D:D, E:E,F:F,G:G,H:H,I:I,J:J,K:K,L:L,M:M,N:N,O:O,P:P", Missing.Value)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            oRng = excel.get_Range("A1", "N1"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.Value2 = " SHIFT REPORT ";  //문구 삽입
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Size = 16;
            oRng.Font.Bold = true;  //볼드
            //  oRng.Font.Color = 0xfffffff;    //폰트 컬러




            oRng = excel.get_Range("A3", "A3"); //해당 범위의 셀 획득
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

            oRng = excel.get_Range("B3", "B3"); //해당 범위의 셀 획득
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

            oRng = excel.get_Range("C3", "C3"); //해당 범위의 셀 획득
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

            oRng = excel.get_Range("D3", "D3"); //해당 범위의 셀 획득
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

            oRng = excel.get_Range("E3", "E3"); //해당 범위의 셀 획득
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

            oRng = excel.get_Range("F3", "F3"); //해당 범위의 셀 획득
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

            oRng = excel.get_Range("G3", "G3"); //해당 범위의 셀 획득
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

            oRng = excel.get_Range("H3", "H3"); //해당 범위의 셀 획득
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

            oRng = excel.get_Range("I3", "I3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;



            oRng = excel.get_Range("J3", "J3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("K3", "K3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("L3", "L3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("M3", "M3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("N3", "N3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("O3", "O3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("P3", "P3"); //해당 범위의 셀 획득
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
                if (form1.iButtonMode == false)
                {
                    if (i == 3 || i == 8 || i == 9 || i == 14 || i == 15)
                    {
                        continue;
                    }
                }
                colIndex++;
                excel.Cells[3, colIndex] = this.listView1.Columns[i].Text;

            }

            for (int i = 0; i < this.listView1.Items.Count; i++)
            {
                rowIndex++;
                colIndex = 0;
                for (int j = 0; j < this.listView1.Items[i].SubItems.Count; j++)
                {
                    if (form1.iButtonMode == false)
                    {
                        if (j == 3 || j == 8 || j == 9 || j == 14 || j == 15)
                        {
                            continue;
                        }
                    }
                    colIndex++;

                    excel.Cells[rowIndex, colIndex] = this.listView1.Items[i].SubItems[j].Text;
                    //     excel.Cells.AutoOutline();

                }
            }
            string Acell = "A";
            string Jcell = "P";
            Acell += (this.listView1.Items.Count + 3).ToString();
            Jcell += (this.listView1.Items.Count + 3).ToString();
            ((Excel.Range)excelWorksheet.get_Range(Acell, Jcell)).Interior.Color = ColorTranslator.ToOle(Color.LightGray);

            System.IO.Directory.CreateDirectory("c:\\Tacho2\\EXCEL");
            string savefile = "c:\\Tacho2\\EXCEL\\" + form1.formname + ".xlsx";

            savefile = "c:\\Tacho2\\EXCEL\\Report_20" + form1.formname[0] + form1.formname[1] + "-" + form1.formname[2] + form1.formname[3] + "-" + form1.formname[4] + form1.formname[5];

            // excelBook.SaveAs(savefile, Excel.XlFileFormat.xlExcel7);

            if (System.IO.File.Exists(savefile + ".xlsx"))  // information 같은 파일의이름이 존재 함
            {

                //    MessageBox.Show("같은 파일 존재");
                //   return;
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
            excelBook.SaveAs(savefile);

            /*    if (excelApp != null) // 작업후 프로세스가 남는 경우를 방지하기 위해서...
                {
                    System.Diagnostics.Process[] pProcess;
                    pProcess = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                    pProcess[0].Kill();
                }*/




            //   Marshal.ReleaseComObject(excelWorksheet);
            //   Marshal.ReleaseComObject(excelBook);
            //   Marshal.ReleaseComObject(excelApp);
            // excel.Visible = true;

            /*
            string filePath = "c:\\test.xlsx";
            Excel.ApplicationClass excel = new Excel.ApplicationClass();
            int colIndex = 0;
            int rowIndex = 1;
            excel.Application.Workbooks.Add(true);

            for (int i = 0; i < listView1.Columns.Count; i++)
            {
                colIndex++;
                excel.Cells[1, colIndex] = listView1.Columns[i].Text;
            }
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                rowIndex++;
                colIndex = 0;
                for (int j = 0; j < listView1.Items[i].SubItems.Count; j++)
                {
                    colIndex++;
                    excel.Cells[rowIndex, colIndex] = listView1.Items[i].SubItems[j].Text;
                    //     excel.Cells.AutoOutline();

                }
            }
            excel.Visible = true;*/


            //excel.Save(filePath);
        }
    }
}