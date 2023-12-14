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
    public partial class CashierForm : Form
    {
        Form1 form1;
        public PrintableListView.PrintableListView m_list;
        public double AllTotal_Income = 0;
        public struct CashierLIST
        {
            public string CashierID;         
            public double TotalIncome;
            public double dot_status;

        }

        public List<string> IDList = new List<string>();

        public CashierForm(Form1 f)
        {
            form1 = f;
            m_list = new PrintableListView.PrintableListView();
            InitializeComponent();
        }

        public void  Cashie_Total()
        {
        
             string Dirname = "";
            int ID_Num = 1;       
            int IdCount = 0;

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


                    int CashierCnt = 0;
                    while (srRead.Read())
                    {
                        if (srRead.IsDBNull(20) == false)
                        {
                            IDList.Add(srRead.GetString(20));
                            CashierCnt++;

                        }
                    }
                    conn.Close();

                    List<string> IDNewList = new List<string>();// 중복 제거를 위해 새로 만든다. IDNewList

                    for (int i = 0; i < IDList.Count; i++)
                    {
                        if (IDNewList.Count == 0)
                        {
                            IDNewList.Add(IDList[i]);
                        }
                        else
                        {
                            bool check = false;
                            for (int j = 0; j < IDNewList.Count; j++)
                            {
                                if(IDList[i].Contains(IDNewList[j]))
                                {
                                    check = true;
                                }
                            }
                            if (check == false)
                            {
                                IDNewList.Add(IDList[i]);
                            }

                        }
                    }


                    CashierLIST[] cashierList = new CashierLIST[IDNewList.Count];

                    for (int i = 0; i < IDNewList.Count; i++)
                    {
                        cashierList[i].CashierID = IDNewList[i];
                    }


                    conn = new OleDbConnection(@DBstring);
                    conn.Open();
                     commRead = new OleDbCommand(queryRead, conn);
                     srRead = commRead.ExecuteReader();

                    double TotalMoney = 0;
                    double TotalDist = 0;


                    while (srRead.Read())
                    {

                          string CashierID = "";
                        if (srRead.IsDBNull(20) == false)
                        {
                            CashierID = srRead.GetString(20);
                            CashierID = CashierID.Trim();
                         
                          
                        }
                        int index = 0;
                        double dotsatus=0;
                        for (int i = 0; i < cashierList.Length; i++)        // id index위치를 찾는다. 
                        {
                            if (cashierList[i].CashierID == CashierID)
                            {
                                index = i;
                                 dotsatus = srRead.GetDouble(10);
                                 cashierList[i].dot_status = dotsatus;


                                 double uuu = (double)srRead.GetDouble(6);

                                 string money;
                                 if (dotsatus != 4 && dotsatus != 2 && dotsatus != 8)
                                 {
                                    
                                     cashierList[i].TotalIncome += uuu;
                                     int hhh = (int)uuu;
                                     money = string.Format("{0}", (int)cashierList[i].TotalIncome);
                                     //    list.SubItems.Add(money);
                                 }
                                 else
                                 {
                                     if (dotsatus == 8)
                                     {
                                       
                                         cashierList[i].TotalIncome += uuu;
                                         money = string.Format("{0:F3}", cashierList[i].TotalIncome);
                                         //  list.SubItems.Add(money);
                                     }
                                     else
                                     {
                                       
                                         cashierList[i].TotalIncome += uuu;
                                         money = string.Format("{0:F2}", cashierList[i].TotalIncome);
                                         //    list.SubItems.Add(money);
                                     }
                                 }// 미터수입
                            }
                        }
                      
  
                       // list.SubItems.Add(CashierID);
                       // listView1.Items.Add(list);
                    }
                    IdCount = 1;

                    AllTotal_Income = 0;
                    for (int i = 0; i < cashierList.Length; i++)
                    {
                        ListViewItem list = new ListViewItem(IdCount.ToString());       // ID
                        list.SubItems.Add(cashierList[i].CashierID);
                        if (cashierList[i].dot_status != 4 && cashierList[i].dot_status != 2 && cashierList[i].dot_status != 8)
                        {
                            string money = string.Format("{0}", (int)cashierList[i].TotalIncome);
                                list.SubItems.Add(money);
                        }
                        else
                        {
                            if (cashierList[i].dot_status == 8)
                            {
                                string  money = string.Format("{0:F3}", cashierList[i].TotalIncome);
                                  list.SubItems.Add(money);
                            }
                            else
                            {
                                string money = string.Format("{0:F2}", cashierList[i].TotalIncome);
                                   list.SubItems.Add(money);
                            }
                        }// 미터수입
                        AllTotal_Income += cashierList[i].TotalIncome;
                        IdCount++;

                      
                        listView1.Items.Add(list);
                    }
                   

                    conn.Close();



                     

           
            int nTotalItems = listView1.Items.Count;
            if (nTotalItems != 0)
            {
                ListViewItem b = new ListViewItem("SUM");
                b.SubItems.Add("");
                int valtemp = (int)AllTotal_Income;
                double vlatemp1 = (double)valtemp;

                string temp = "";
                if ((AllTotal_Income % vlatemp1) != 0)
                {


                    temp = string.Format("{0:F2}", AllTotal_Income);
                    b.SubItems.Add(temp);
                }
                else
                {
                    temp = string.Format("{0}", AllTotal_Income);
                    b.SubItems.Add(temp);
                }
                b.BackColor = System.Drawing.Color.LightGray;
                listView1.Items.Add(b);
            }

      


           FillList(m_list, listView1);
        }
        public void FillList(ListView list, ListView table)
        {
        //    list.SuspendLayout();

            // Clear list
            list.Items.Clear();
            list.Columns.Clear();

            // Columns
            int nCol = 0;


            int a = 0;


            try
            {
                for (int i = 0; i < 3; i++)
                {


                    ColumnHeader[] col = new ColumnHeader[3];
                    ColumnHeader ch = new ColumnHeader();

                    col[i] = table.Columns[i];
                    ch.Text = col[i].Text;
                    ch.TextAlign = HorizontalAlignment.Right;
                   
                    switch (nCol)
                    {
                        case 0: ch.Width = 100; break;       // id   

                        case 1: ch.Width = 200; break;       // 차량번호

                        case 2: ch.Width = 400; break;       // 기사 번호
                          
                     
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
                  //  list.Height = 100;
                    list.Items.Add(item);
                }



                list.ResumeLayout();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CashierForm_Load(object sender, EventArgs e)
        {
           
        }

        private void listView1_SubItemClicked(object sender, ListViewEx.SubItemEventArgs e)
        {
            try
            {
                // if (e.SubItem == 0 || e.SubItem == 1 || e.SubItem == 2 || e.SubItem == 3 || e.SubItem == 4 || e.SubItem == 5 || e.SubItem == 6 || e.SubItem == 7)
                if (e.SubItem >= 0 && e.SubItem <= 2)
                {
                    if (listView1.SelectedItems.Count == 0)
                    {

                        return;
                    }
                    ListView.SelectedListViewItemCollection items = listView1.SelectedItems;

                    foreach (ListViewItem item in items)
                    {
                        if ((item.SubItems[0].Text.CompareTo("SUM") == 0))
                        {
                         
                            return;
                        }
                        else
                        {
                         
                        }
                    }

                    Cashier_DetailForm Cashier_detail = new Cashier_DetailForm(form1);
                //    Cashier_detail.Read_Transaction(nOpenedindex);
                    Cashier_detail.MdiParent = this.ParentForm;

                    //  transactionform.MdiParent = this;
                    Cashier_detail.BringToFront();
                    LayoutMdi(MdiLayout.TileHorizontal);

                    string str_id = listView1.SelectedItems[0].SubItems[1].Text;

                    Cashier_detail.Cashier_read(str_id);
                    Cashier_detail.Focus();
                    Cashier_detail.Show();

                  
                  
                }

            }
            catch (Exception excep)
            {
                MessageBox.Show(excep.ToString());
            }
        }

        private void buttonPrint_Click(object sender, EventArgs e)
        {
            m_list.Title = "CASHIER LIST\n";


            //m_list.FitToPage = m_cbFitToPage.Checked;

            //	m_list.PageSetup();

            m_list.FitToPage = true;

            m_list.PrintPreview();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string filePath = "c:\\Tacho_Report.xlsx";
            Excel.ApplicationClass excel = new Excel.ApplicationClass();

            int colIndex = 0;
            int rowIndex = 3;
            Excel.Application excelApp = null;
            excel.UserName ="Cashier_List.xlsx";


            object missingType = Type.Missing;
            object fileName =   "Cashier_List.xlsx";


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
            excelWorksheet.Name = "Cashier_List";
            ((Excel.Range)excelWorksheet.get_Range("A:A", Missing.Value)).ColumnWidth = 8.33;
            ((Excel.Range)excelWorksheet.get_Range("B:B", Missing.Value)).ColumnWidth = 25.33;
            ((Excel.Range)excelWorksheet.get_Range("C:C", Missing.Value)).ColumnWidth = 25.33;
         


            // BackColor 설정
            ((Excel.Range)excelWorksheet.get_Range("A3", "A3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("B3", "B3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("C3", "C3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
         



            ((Excel.Range)excelWorksheet.get_Range("A:A, B:B, C:C", Missing.Value)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            oRng = excel.get_Range("A1", "N1"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.Value2 = " CASHIER LIST ";  //문구 삽입
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
            string Jcell = "C";
            Acell += (this.listView1.Items.Count + 3).ToString();
            Jcell += (this.listView1.Items.Count + 3).ToString();
            ((Excel.Range)excelWorksheet.get_Range(Acell, Jcell)).Interior.Color = ColorTranslator.ToOle(Color.LightGray);

            System.IO.Directory.CreateDirectory("c:\\Tacho2\\EXCEL");
            string savefile = "c:\\Tacho2\\EXCEL\\Cashier_List.xlsx";

            savefile = "c:\\Tacho2\\EXCEL\\Cashier_List_20" + form1.formname[0] + form1.formname[1] + "-" + form1.formname[2] + form1.formname[3] + "-" + form1.formname[4] + form1.formname[5];


            // 같은파일 이름이 있는지 검사후 존해 한다면 이름을 변경 한다. 

            // excelBook.SaveAs(savefile, Excel.XlFileFormat.xlExcel7);

            excel.Visible = true;


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
            excelBook.SaveAs(savefile); 
        }



    }
}
