using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Data.OleDb;
using PrintableListView;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace TachoPlus
{
  
   

    public partial class TransactionForm : Form
    {

        public string TACHO2_path = "";
        Form1 form1;
        private iniClass inicls = new iniClass();
        private OleDbConnection conn;
        private OpenedDBInfo openDBInfo;
        byte[] Sales_Temp;
        public Total total;
        public DateTime dt;
        public TransactionForm(Form1 f)
        {

            InitializeComponent();
            form1 = f;
            string path1 = Application.StartupPath + "\\WinTacho.ini";
            TACHO2_path = inicls.GetIniValue("Tacho Init", "path", path1); // 타코 루트
            total = new Total();
            m_list = new PrintableListView.PrintableListView();
            if (form1.iButtonMode == false)
            {
                label1.Visible = false;
                label3.Visible = false;
            }
        }


        public struct Total
        {
            public double tMoney;
            public double tDistS;
            public double tDisteD;
            public double tVacantDist;
            public double tTotalDist;

        }
        private struct OpenedDBInfo
        {
            public string CarNo;
            public string DriverNo;
            //public string DriverName;
            public DateTime OutTime;
            public DateTime InTime;
            public string Title;
        }
        private int BcdToDecimalByLsb(byte[] arr, int cnt)
        {
            int rtValue = 0, mulValue = 0;

            for (int i = 0; i < cnt; i++)
            {
                if (i == 0) mulValue = 1;
                else if (i == 1) mulValue = 100;
                else if (i == 2) mulValue = 10000;
                else if (i == 3) mulValue = 1000000;
                else mulValue = 0;

                rtValue += (((arr[i] >> 4) * 10) + (arr[i] & 0x0F)) * mulValue;
            }

            return rtValue;
        }

        private int BcdToDecimal(byte bTemp)
        {
            return (((bTemp >> 4) * 10) + (bTemp & 0x0F));
        }

        int BinToBcd8(int n)
        {
            return ((int)((n / 10) << 4) + (n % 10));
        }

        public byte[] BinToBcd24P(byte[] arr, int n)
        {

            int rulTmp;

            rulTmp = n % 1000000;
            arr[2] = (byte)BinToBcd8(rulTmp / 10000);
            rulTmp = n % 10000;
            arr[1] = (byte)BinToBcd8(rulTmp / 100);
            arr[0] = (byte)BinToBcd8(rulTmp % 100);
            return arr;
        }



     
        public void  Read_Transaction(int id)
        {
             if (listView1.Items.Count > 0)
                    listView1.Items.Clear();

                listView1.View = View.Details;
                listView1.GridLines = true;                   //   리스트 뷰 라인생성
                listView1.FullRowSelect = true;               // 라인 선택 */

                DateTime PaidTime = new DateTime(2015, 1, 1, 1, 1, 1);
                string DBstring = "";
                string NameDB = "";
                total.tDisteD = 0;
                total.tDistS = 0;
                total.tMoney = 0;

                NameDB = TACHO2_path + form1.mdbfilename;             
                DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + NameDB;


                conn = new OleDbConnection(@DBstring);
                conn.Open();


                string queryRead1 = "SELECT * FROM TblTacho WHERE ID=" + id.ToString();
                OleDbCommand commRead1 = new OleDbCommand(queryRead1, conn);
                OleDbDataReader srRead1 = commRead1.ExecuteReader();

                int SalesLength = 0;
                string strCarNo = "";
             
                double dot_Stauts = 0;
               
                string strDriverNo = "";
                string DriverName ="";
                double MoneyCheck = 0;
                string CashierID = "";
              
                while (srRead1.Read())
                {
                   // srRead1.GetBytes(27, 0, Sales_Temp, 0, 125000);  // 상세영업
                    MoneyCheck = srRead1.GetDouble(6);

                    if (MoneyCheck == 0)
                    {
                        MessageBox.Show("There is no sales history.");
                        form1.TransactionForm_Run = true;
                        return;
                    }
                     strCarNo = srRead1.GetString(1);  // 차량번호 if (!srRead1.IsDBNull(27))
                   //    strDriverNo = srRead1.GetString(2);
                    if (srRead1.IsDBNull(2) == false)
                    {
                        strDriverNo = srRead1.GetString(2);


                    }

                    if (srRead1.IsDBNull(3) == false)
                    {
                        DriverName = srRead1.GetString(3);


                    }

                    if (srRead1.IsDBNull(20) == false)
                    {
                        CashierID =  srRead1.GetString(20);

                   }
                   

                     if (!srRead1.IsDBNull(18))
                     {
                         SalesLength = srRead1.GetInt32(18);
                     }
                     Sales_Temp = new byte[SalesLength];
                        if (!srRead1.IsDBNull(27))
                        {
                            srRead1.GetBytes(27, 0, Sales_Temp, 0, SalesLength);
                        }

                        dot_Stauts = srRead1.GetDouble(10);

                        if (srRead1.IsDBNull(19) == false)                                            // 
                        {
                        //    a.SubItems.Add(srRead1.GetDateTime(19).ToString("yyyy-MM-dd"));                       //  날짜 
                         //   a.SubItems.Add(srRead1.GetDateTime(19).ToString(" HH:mm:ss"));                       //  시간  
                            PaidTime = srRead1.GetDateTime(19);
                        }
                        else
                        {
                          PaidTime =new DateTime(2099, 1, 1, 1, 1, 1);
                        }
                }
                label1.Text = "DRIVER NAME :   " + DriverName;
                label2.Text = "TAXI ID :   " + strCarNo;
                label3.Text = "DRIVER ID :   " + strDriverNo;
                int loopcnt = SalesLength / 64;
                int idcnt = 1;
                int sales_cnt = 0;

                if (form1.iButtonMode == false)
                {
                    /*    this.listView1.Columns[7].Text ="Reserved";
                        this.listView1.Columns[8].Text = "Reserved";
                        this.listView1.Columns[13].Text = "Reserved";
                        this.listView1.Columns[14].Text = "Reserved";
                        this.listView1.Columns[15].Text = "Reserved";
                        this.listView1.Columns[21].Text = "Reserved";
                        this.listView1.Columns[22].Text = "Reserved";*/

                    //ColumnHeader oldch = null;

                    this.listView1.Columns[0].Text = "ID";
                    this.listView1.Columns[1].Text = "Trip No.";
                    this.listView1.Columns[2].Text = "Trip Start Date";
                    this.listView1.Columns[3].Text = "Trip Start Time";
                    this.listView1.Columns[4].Text = "Trip End Date";
                    this.listView1.Columns[5].Text = "Trip End Time";
                    this.listView1.Columns[6].Text = "Income";
                    this.listView1.Columns[7].Text = "Hired Distances";
                    this.listView1.Columns[8].Text = "Vacant Distances";
                    this.listView1.Columns[9].Text = "Total Distances";
                    this.listView1.Columns[10].Text = "Extra";
                    this.listView1.Columns[11].Text = "Call";
                    this.listView1.Columns[12].Text = "A/P";
                    this.listView1.Columns[13].Text = "Luggage";
                    this.listView1.Columns[14].Text = "Toll";

                    /*  ColumnHeader oldch = null;

                      for (int i = 15; i < 19; i++)
                      {
                          if (listView1.Columns[i].DisplayIndex == i)
                          {
                              oldch = listView1.Columns[i];
                              this.listView1.Columns.Remove(oldch);
                          }
                      }*/



                    this.listView1.Columns[15].Text = "Reserved";
                    this.listView1.Columns[15].Width = 0;
                    this.listView1.Columns[16].Text = "Reserved";
                    this.listView1.Columns[16].Width = 0;
                    this.listView1.Columns[17].Text = "Reserved";
                    this.listView1.Columns[17].Width = 0;
                    this.listView1.Columns[18].Text = "Reserved";
                    this.listView1.Columns[18].Width = 0;
                    this.listView1.Columns[19].Text = "Reserved";
                    this.listView1.Columns[19].Width = 0;
                    this.listView1.Columns[20].Text = "Reserved";
                    this.listView1.Columns[20].Width = 0;
                    this.listView1.Columns[21].Text = "Reserved";
                    this.listView1.Columns[21].Width = 0;
                    this.listView1.Columns[22].Text = "Reserved";
                    this.listView1.Columns[22].Width = 0;


                }
                else  // iButton Mode
                {
                    if (form1.MALAYSIA_Set == true)
                    {
                        this.listView1.Columns[19].Text = "Fare";
                        this.listView1.Columns[17].Text = "Night";
                        this.listView1.Columns[15].Text = "Toll_Cnt";
                    }
                }
                
                   while (loopcnt!=0)
                   {
                        ListViewItem a = new ListViewItem(idcnt.ToString());

                     //   a.SubItems.Add(strCarNo);
                       // a.SubItems.Add(strDriverNo);
                      //  a.SubItems.Add(DriverName);
                         string money = "";
                         double Intmoney = 0;
                         double VacantDist = 0;
                         double Dist = 0;
                         double TotalDist = 0;
                        if (form1.iButtonMode == false)
                        {
                            try
                            {
                                if (Sales_Temp[sales_cnt + 0] == 0xAA)
                                {
                                 //   listView1.Items.Add(a);
                                    loopcnt--;
                                    idcnt++;
                                    sales_cnt += 64;
                                    continue;
                                }

                                int tripNo = BcdToDecimal(Sales_Temp[sales_cnt + 0]) + (BcdToDecimal(Sales_Temp[sales_cnt + 1]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 2]) * 10000);
                             
                                a.SubItems.Add(tripNo.ToString());  // trip no

                              //  a.SubItems.Add(DriverName);  // driver Name

                                DateTime INTime = new DateTime(BcdToDecimal(Sales_Temp[sales_cnt + 3]) + 2000, BcdToDecimal(Sales_Temp[sales_cnt + 4]), BcdToDecimal(Sales_Temp[sales_cnt + 5]), BcdToDecimal(Sales_Temp[sales_cnt + 6]),
                                                     BcdToDecimal(Sales_Temp[sales_cnt + 7]), BcdToDecimal(Sales_Temp[sales_cnt + 8]));
                                string instr = INTime.ToString("yyyy-MM-dd");
                                a.SubItems.Add(instr);

                                instr = INTime.ToString("HH:mm");
                                a.SubItems.Add(instr);


                                DateTime OUTTime = new DateTime(BcdToDecimal(Sales_Temp[sales_cnt + 9]) + 2000, BcdToDecimal(Sales_Temp[sales_cnt + 10]), BcdToDecimal(Sales_Temp[sales_cnt + 11]), BcdToDecimal(Sales_Temp[sales_cnt + 12]),
                                                    BcdToDecimal(Sales_Temp[sales_cnt + 13]), BcdToDecimal(Sales_Temp[sales_cnt + 14]));
                                string outstr = OUTTime.ToString("yyyy-MM-dd ");
                                a.SubItems.Add(outstr);

                                outstr = OUTTime.ToString("HH:mm");
                                a.SubItems.Add(outstr);


                            }
                            catch
                            {

                                DateTime INTime = new DateTime(2015, 1, 1, 1, 1, 1);
                                string instr = INTime.ToString("yyyy-MM-dd");
                                a.SubItems.Add(instr);

                                instr = INTime.ToString("HH:mm");
                                a.SubItems.Add(instr);


                                DateTime OUTTime = new DateTime(2015, 1, 1, 1, 1, 1);
                                string outstr = OUTTime.ToString("yyyy-MM-dd ");
                                a.SubItems.Add(outstr);

                                outstr = OUTTime.ToString("HH:mm");
                                a.SubItems.Add(outstr);

                            }

                            if (PaidTime.Year == 2099)
                            {
                             //   a.SubItems.Add("");
                             //   a.SubItems.Add("");
                            }
                            else
                            {


                                string paid = PaidTime.ToString("yyyy-MM-dd");
                               // a.SubItems.Add(paid);

                                paid = PaidTime.ToString("HH:mm");
                              //  a.SubItems.Add(paid);
                            }





                            money = "";
                            Intmoney = BcdToDecimal(Sales_Temp[sales_cnt + 55]) + (BcdToDecimal(Sales_Temp[sales_cnt + 56]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 57]) * 10000) +
                                               (BcdToDecimal(Sales_Temp[sales_cnt + 58]) * 1000000) + (BcdToDecimal(Sales_Temp[sales_cnt + 59]) * 100000000);



                            if (dot_Stauts == 4)
                            {
                                Intmoney = Intmoney / 100;
                                money = string.Format("{0:F2}", Intmoney);
                            }
                            else if (dot_Stauts == 8)
                            {
                                Intmoney = Intmoney / 1000;
                                money = string.Format("{0:F3}", Intmoney);
                            }
                            else
                            {
                                money = string.Format("{0:D}", (int)Intmoney);
                            }

                            a.SubItems.Add(money);
                            // hired dist
                            Dist = BcdToDecimal(Sales_Temp[sales_cnt + 60]) + (BcdToDecimal(Sales_Temp[sales_cnt + 61]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 62]) * 10000);
                            Dist = Dist / 1000;

                            string dist = string.Format("{0:N3} Km", Dist);
                            a.SubItems.Add(dist);

                            // vacant dist
                            VacantDist = BcdToDecimal(Sales_Temp[sales_cnt + 45]) + (BcdToDecimal(Sales_Temp[sales_cnt + 46]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 47]) * 10000);
                            VacantDist = VacantDist / 1000;

                            dist = string.Format("{0:N3} Km", VacantDist);
                            a.SubItems.Add(dist);

                            // total dist
                          //  TotalDist = BcdToDecimal(Sales_Temp[sales_cnt + 48]) + (BcdToDecimal(Sales_Temp[sales_cnt + 49]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 50]) * 10000);
                            TotalDist = Dist + VacantDist;  // 계산으로 변경 
                          //  TotalDist = TotalDist / 1000;

                            dist = string.Format("{0:N3} Km", TotalDist);
                            a.SubItems.Add(dist);
                            int intspeed = 0;
                            intspeed = (Sales_Temp[sales_cnt + 52] << 8) + Sales_Temp[sales_cnt + 51];
                            //     double Maxspeed =Sales_Temp[sales_cnt + 51] +(Sales_Temp[sales_cnt + 52] * 100);
                            double Maxspeed = (double)intspeed / 10;

                            string speedstr = string.Format("{0:N1} Km/h", Maxspeed);
                          //  a.SubItems.Add(speedstr);

                            int sensor = (Sales_Temp[sales_cnt + 54] & 0x0F);               // sensor + PowerCheck
                            if (form1.iButtonMode == false)
                            {
                                sensor = 0;
                            }
                           // a.SubItems.Add(sensor.ToString());  // sensor


                            int PowerCheck = (Sales_Temp[sales_cnt + 54] & 0xF0);

                            if (PowerCheck > 0)
                            {
                                PowerCheck = 1;
                            }


                            int TollCnt = Sales_Temp[sales_cnt + 44];
                            if (form1.iButtonMode == false)
                            {
                                TollCnt = 0;
                            }
                          //  a.SubItems.Add(TollCnt.ToString()); // salik

                            ////////////////////////  call
                            string strcall = "";
                            double Intcall = BcdToDecimal(Sales_Temp[sales_cnt + 15]) + (BcdToDecimal(Sales_Temp[sales_cnt + 16]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 17]) * 10000) +
                                                (BcdToDecimal(Sales_Temp[sales_cnt + 18]) * 1000000) + (BcdToDecimal(Sales_Temp[sales_cnt + 19]) * 100000000);



                            if (dot_Stauts == 4)
                            {
                                Intcall = Intcall / 100;
                                strcall = string.Format("{0:F2}", Intcall);
                            }
                            else if (dot_Stauts == 8)
                            {
                                Intcall = Intcall / 1000;
                                strcall = string.Format("{0:F3}", Intcall);
                            }
                            else
                            {
                                strcall = string.Format("{0:D}", (int)Intcall);
                            }

                            a.SubItems.Add(strcall);
                            ////////////////////////

                            ////////////////////////  Lugg
                            string strlugg = "";
                            double Intlugg = BcdToDecimal(Sales_Temp[sales_cnt + 20]) + (BcdToDecimal(Sales_Temp[sales_cnt + 21]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 22]) * 10000) +
                                                (BcdToDecimal(Sales_Temp[sales_cnt + 23]) * 1000000) + (BcdToDecimal(Sales_Temp[sales_cnt + 24]) * 100000000);



                            if (dot_Stauts == 4)
                            {
                                Intlugg = Intlugg / 100;
                                strlugg = string.Format("{0:F2}", Intlugg);
                            }
                            else if (dot_Stauts == 8)
                            {
                                Intlugg = Intlugg / 1000;
                                strlugg = string.Format("{0:F3}", Intlugg);
                            }
                            else
                            {
                                strlugg = string.Format("{0:D}", (int)Intlugg);
                            }

                            a.SubItems.Add(strlugg);
                            ////////////////////////


                            //////////////////////// ap
                            string strap = "";
                            double Intap = BcdToDecimal(Sales_Temp[sales_cnt + 25]) + (BcdToDecimal(Sales_Temp[sales_cnt + 26]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 27]) * 10000) +
                                                (BcdToDecimal(Sales_Temp[sales_cnt + 28]) * 1000000) + (BcdToDecimal(Sales_Temp[sales_cnt + 29]) * 100000000);



                            if (dot_Stauts == 4)
                            {
                                Intap = Intap / 100;
                                strap = string.Format("{0:F2}", Intap);
                            }
                            else if (dot_Stauts == 8)
                            {
                                Intap = Intap / 1000;
                                strap = string.Format("{0:F3}", Intap);
                            }
                            else
                            {
                                strap = string.Format("{0:D}", (int)Intap);
                            }

                            a.SubItems.Add(strap);
                            ////////////////////////


                            //////////////////////// extra
                            string strextra = "";
                            double Intextra = BcdToDecimal(Sales_Temp[sales_cnt + 30]) + (BcdToDecimal(Sales_Temp[sales_cnt + 31]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 32]) * 10000) +
                                                (BcdToDecimal(Sales_Temp[sales_cnt + 33]) * 1000000) + (BcdToDecimal(Sales_Temp[sales_cnt + 34]) * 100000000);



                            if (dot_Stauts == 4)
                            {
                                Intextra = Intextra / 100;
                                strextra = string.Format("{0:F2}", Intextra);
                            }
                            else if (dot_Stauts == 8)
                            {
                                Intextra = Intextra / 1000;
                                strextra = string.Format("{0:F3}", Intextra);
                            }
                            else
                            {
                                strextra = string.Format("{0:D}", (int)Intextra);
                            }

                            a.SubItems.Add(strextra);
                            ////////////////////////



                            //////////////////////// toll
                            string strtoll = "";
                            double Inttoll = BcdToDecimal(Sales_Temp[sales_cnt + 35]) + (BcdToDecimal(Sales_Temp[sales_cnt + 36]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 37]) * 10000) +
                                                (BcdToDecimal(Sales_Temp[sales_cnt + 38]) * 1000000) + (BcdToDecimal(Sales_Temp[sales_cnt + 39]) * 100000000);



                            if (dot_Stauts == 4)
                            {
                                Inttoll = Inttoll / 100;
                                strtoll = string.Format("{0:F2}", Inttoll);
                            }
                            else if (dot_Stauts == 8)
                            {
                                Inttoll = Inttoll / 1000;
                                strtoll = string.Format("{0:F3}", Inttoll);
                            }
                            else
                            {
                                strtoll = string.Format("{0:D}", (int)Inttoll);
                            }

                            a.SubItems.Add(strtoll);
                            ////////////////////////
                            if (form1.iButtonMode == false)
                            {
                                PowerCheck = 0;
                                CashierID = "";
                            }
                            a.SubItems.Add(""); // driver name
                            a.SubItems.Add("");   // paid time
                            a.SubItems.Add("");
                            a.SubItems.Add("");
                            a.SubItems.Add("");  // sensor
                            a.SubItems.Add(""); // salik
                            a.SubItems.Add(""); // driver Name
                            a.SubItems.Add("");
                            a.SubItems.Add("");
                        }
                        else
                        {

                            try
                            {


                                int tripNo = BcdToDecimal(Sales_Temp[sales_cnt + 0]) + (BcdToDecimal(Sales_Temp[sales_cnt + 1]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 2]) * 10000);
                                a.SubItems.Add(strCarNo);
                                a.SubItems.Add(strDriverNo);
                                a.SubItems.Add(tripNo.ToString());  // trip no

                                a.SubItems.Add(DriverName);  // driver Name

                                DateTime INTime = new DateTime(BcdToDecimal(Sales_Temp[sales_cnt + 3]) + 2000, BcdToDecimal(Sales_Temp[sales_cnt + 4]), BcdToDecimal(Sales_Temp[sales_cnt + 5]), BcdToDecimal(Sales_Temp[sales_cnt + 6]),
                                                     BcdToDecimal(Sales_Temp[sales_cnt + 7]), BcdToDecimal(Sales_Temp[sales_cnt + 8]));
                                string instr = INTime.ToString("yyyy-MM-dd");
                                a.SubItems.Add(instr);

                                instr = INTime.ToString("HH:mm");
                                a.SubItems.Add(instr);


                                DateTime OUTTime = new DateTime(BcdToDecimal(Sales_Temp[sales_cnt + 9]) + 2000, BcdToDecimal(Sales_Temp[sales_cnt + 10]), BcdToDecimal(Sales_Temp[sales_cnt + 11]), BcdToDecimal(Sales_Temp[sales_cnt + 12]),
                                                    BcdToDecimal(Sales_Temp[sales_cnt + 13]), BcdToDecimal(Sales_Temp[sales_cnt + 14]));
                                string outstr = OUTTime.ToString("yyyy-MM-dd ");
                                a.SubItems.Add(outstr);

                                outstr = OUTTime.ToString("HH:mm");
                                a.SubItems.Add(outstr);


                            }
                            catch
                            {

                                DateTime INTime = new DateTime(2015, 1, 1, 1, 1, 1);
                                string instr = INTime.ToString("yyyy-MM-dd");
                                a.SubItems.Add(instr);

                                instr = INTime.ToString("HH:mm");
                                a.SubItems.Add(instr);


                                DateTime OUTTime = new DateTime(2015, 1, 1, 1, 1, 1);
                                string outstr = OUTTime.ToString("yyyy-MM-dd ");
                                a.SubItems.Add(outstr);

                                outstr = OUTTime.ToString("HH:mm");
                                a.SubItems.Add(outstr);

                            }

                            if (PaidTime.Year == 2099)
                            {
                                a.SubItems.Add("");
                                a.SubItems.Add("");
                            }
                            else
                            {


                                string paid = PaidTime.ToString("yyyy-MM-dd");
                                a.SubItems.Add(paid);

                                paid = PaidTime.ToString("HH:mm");
                                a.SubItems.Add(paid);
                            }





                             money = "";
                             Intmoney = BcdToDecimal(Sales_Temp[sales_cnt + 55]) + (BcdToDecimal(Sales_Temp[sales_cnt + 56]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 57]) * 10000) +
                                                (BcdToDecimal(Sales_Temp[sales_cnt + 58]) * 1000000) + (BcdToDecimal(Sales_Temp[sales_cnt + 59]) * 100000000);



                            if (dot_Stauts == 4)
                            {
                                Intmoney = Intmoney / 100;
                                money = string.Format("{0:F2}", Intmoney);
                            }
                            else if (dot_Stauts == 8)
                            {
                                Intmoney = Intmoney / 1000;
                                money = string.Format("{0:F3}", Intmoney);
                            }
                            else
                            {
                                money = string.Format("{0:D}", (int)Intmoney);
                            }

                            a.SubItems.Add(money);
                            // hired dist
                             Dist = BcdToDecimal(Sales_Temp[sales_cnt + 60]) + (BcdToDecimal(Sales_Temp[sales_cnt + 61]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 62]) * 10000);
                            Dist = Dist / 1000;

                            string dist = string.Format("{0:N3} Km", Dist);
                            a.SubItems.Add(dist);

                            // vacant dist
                             VacantDist = BcdToDecimal(Sales_Temp[sales_cnt + 45]) + (BcdToDecimal(Sales_Temp[sales_cnt + 46]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 47]) * 10000);
                            VacantDist = VacantDist / 1000;

                            dist = string.Format("{0:N3} Km", VacantDist);
                            a.SubItems.Add(dist);

                            // total dist
                             TotalDist = BcdToDecimal(Sales_Temp[sales_cnt + 48]) + (BcdToDecimal(Sales_Temp[sales_cnt + 49]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 50]) * 10000);
                            TotalDist = TotalDist / 1000;

                            dist = string.Format("{0:N3} Km", TotalDist);
                            a.SubItems.Add(dist);
                            int intspeed = 0;
                            intspeed = (Sales_Temp[sales_cnt + 52] << 8) + Sales_Temp[sales_cnt + 51];
                            //     double Maxspeed =Sales_Temp[sales_cnt + 51] +(Sales_Temp[sales_cnt + 52] * 100);
                            double Maxspeed = (double)intspeed / 10;

                            string speedstr = string.Format("{0:N1} Km/h", Maxspeed);
                            a.SubItems.Add(speedstr);

                            int sensor = (Sales_Temp[sales_cnt + 54] & 0x0F);               // sensor + PowerCheck
                            if (form1.iButtonMode == false)
                            {
                                sensor = 0;
                            }
                            a.SubItems.Add(sensor.ToString());  // sensor


                            int PowerCheck = (Sales_Temp[sales_cnt + 54] & 0xF0);

                            if (PowerCheck > 0)
                            {
                                PowerCheck = 1;
                            }


                            int TollCnt = Sales_Temp[sales_cnt + 44];
                            if (form1.iButtonMode == false)
                            {
                                TollCnt = 0;
                            }
                          
                                a.SubItems.Add(TollCnt.ToString()); // salik
                           

                            ////////////////////////  call
                            string strcall = "";
                            double Intcall = BcdToDecimal(Sales_Temp[sales_cnt + 15]) + (BcdToDecimal(Sales_Temp[sales_cnt + 16]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 17]) * 10000) +
                                                (BcdToDecimal(Sales_Temp[sales_cnt + 18]) * 1000000) + (BcdToDecimal(Sales_Temp[sales_cnt + 19]) * 100000000);



                            if (dot_Stauts == 4)
                            {
                                Intcall = Intcall / 100;
                                strcall = string.Format("{0:F2}", Intcall);
                            }
                            else if (dot_Stauts == 8)
                            {
                                Intcall = Intcall / 1000;
                                strcall = string.Format("{0:F3}", Intcall);
                            }
                            else
                            {
                                strcall = string.Format("{0:D}", (int)Intcall);
                            }

                            a.SubItems.Add(strcall);
                            ////////////////////////

                            ////////////////////////  Lugg
                            string strlugg = "";
                            double Intlugg = BcdToDecimal(Sales_Temp[sales_cnt + 20]) + (BcdToDecimal(Sales_Temp[sales_cnt + 21]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 22]) * 10000) +
                                                (BcdToDecimal(Sales_Temp[sales_cnt + 23]) * 1000000) + (BcdToDecimal(Sales_Temp[sales_cnt + 24]) * 100000000);



                            if (dot_Stauts == 4)
                            {
                                Intlugg = Intlugg / 100;
                                strlugg = string.Format("{0:F2}", Intlugg);
                            }
                            else if (dot_Stauts == 8)
                            {
                                Intlugg = Intlugg / 1000;
                                strlugg = string.Format("{0:F3}", Intlugg);
                            }
                            else
                            {
                                strlugg = string.Format("{0:D}", (int)Intlugg);
                            }

                            a.SubItems.Add(strlugg);
                            ////////////////////////


                            //////////////////////// ap
                            string strap = "";
                            double Intap = BcdToDecimal(Sales_Temp[sales_cnt + 25]) + (BcdToDecimal(Sales_Temp[sales_cnt + 26]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 27]) * 10000) +
                                                (BcdToDecimal(Sales_Temp[sales_cnt + 28]) * 1000000) + (BcdToDecimal(Sales_Temp[sales_cnt + 29]) * 100000000);



                            if (dot_Stauts == 4)
                            {
                                Intap = Intap / 100;
                                strap = string.Format("{0:F2}", Intap);
                            }
                            else if (dot_Stauts == 8)
                            {
                                Intap = Intap / 1000;
                                strap = string.Format("{0:F3}", Intap);
                            }
                            else
                            {
                                strap = string.Format("{0:D}", (int)Intap);
                            }

                            a.SubItems.Add(strap);
                            ////////////////////////


                            //////////////////////// extra
                            string strextra = "";
                            double Intextra = BcdToDecimal(Sales_Temp[sales_cnt + 30]) + (BcdToDecimal(Sales_Temp[sales_cnt + 31]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 32]) * 10000) +
                                                (BcdToDecimal(Sales_Temp[sales_cnt + 33]) * 1000000) + (BcdToDecimal(Sales_Temp[sales_cnt + 34]) * 100000000);



                            if (dot_Stauts == 4)
                            {
                                Intextra = Intextra / 100;
                                strextra = string.Format("{0:F2}", Intextra);
                            }
                            else if (dot_Stauts == 8)
                            {
                                Intextra = Intextra / 1000;
                                strextra = string.Format("{0:F3}", Intextra);
                            }
                            else
                            {
                                strextra = string.Format("{0:D}", (int)Intextra);
                            }

                            a.SubItems.Add(strextra);
                            ////////////////////////



                            //////////////////////// toll
                            string strtoll = "";
                            double Inttoll = BcdToDecimal(Sales_Temp[sales_cnt + 35]) + (BcdToDecimal(Sales_Temp[sales_cnt + 36]) * 100) + (BcdToDecimal(Sales_Temp[sales_cnt + 37]) * 10000) +
                                                (BcdToDecimal(Sales_Temp[sales_cnt + 38]) * 1000000) + (BcdToDecimal(Sales_Temp[sales_cnt + 39]) * 100000000);



                            if (dot_Stauts == 4)
                            {
                                Inttoll = Inttoll / 100;
                                strtoll = string.Format("{0:F2}", Inttoll);
                            }
                            else if (dot_Stauts == 8)
                            {
                                Inttoll = Inttoll / 1000;
                                strtoll = string.Format("{0:F3}", Inttoll);
                            }
                            else
                            {
                                strtoll = string.Format("{0:D}", (int)Inttoll);
                            }

                            a.SubItems.Add(strtoll);
                            ////////////////////////
                            if (form1.iButtonMode == false)
                            {
                                PowerCheck = 0;
                                CashierID = "";
                            }
                            a.SubItems.Add(PowerCheck.ToString());
                            a.SubItems.Add(CashierID);
                        }


                        total.tMoney += Intmoney;

                        total.tDistS += Dist;   // hired

                        total.tVacantDist += VacantDist;  // VacantDist dist

                        total.tTotalDist += TotalDist;




                        listView1.Items.Add(a);
                        loopcnt--;
                        idcnt++;
                        sales_cnt += 64;
                   }

                   ListViewItem b = new ListViewItem("SUM");
                //   b.SubItems.Add("");                             ///   공백 만들기
                //   b.SubItems.Add("");
               //    b.SubItems.Add("");
                   if (form1.iButtonMode == false)
                   {
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                    //   b.SubItems.Add("");
                    //   b.SubItems.Add("");
                    //   b.SubItems.Add("");
                       string aa = "";
                       if (dot_Stauts == 4)
                       {
                           aa = string.Format("{0:F2}", total.tMoney);     // 미터수입
                       }
                       else if (dot_Stauts == 8)
                       {
                           aa = string.Format("{0:F3}", total.tMoney);     // 미터수입
                       }
                       else
                       {
                           aa = string.Format("{0:D}", (int)total.tMoney);     // 미터수입
                       }
                       b.SubItems.Add(aa);   // 미터 수입

                       aa = string.Format("{0:N3} Km", total.tDistS);  // 영업거리
                       b.SubItems.Add(aa);// 영업거리

                       aa = string.Format("{0:N3} Km", total.tVacantDist);  //빈차거리
                       b.SubItems.Add(aa);// 빈차거리

                       aa = string.Format("{0:N3} Km", total.tTotalDist);  //토탈거리
                       b.SubItems.Add(aa);// 토탈거리

                       //  b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                   }
                   else
                   {
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       string aa = "";
                       if (dot_Stauts == 4)
                       {
                           aa = string.Format("{0:F2}", total.tMoney);     // 미터수입
                       }
                       else if (dot_Stauts == 8)
                       {
                           aa = string.Format("{0:F3}", total.tMoney);     // 미터수입
                       }
                       else
                       {
                           aa = string.Format("{0:D}", (int)total.tMoney);     // 미터수입
                       }
                       b.SubItems.Add(aa);   // 미터 수입

                       aa = string.Format("{0:N3} Km", total.tDistS);  // 영업거리
                       b.SubItems.Add(aa);// 영업거리

                       aa = string.Format("{0:N3} Km", total.tVacantDist);  //빈차거리
                       b.SubItems.Add(aa);// 빈차거리

                       aa = string.Format("{0:N3} Km", total.tTotalDist);  //토탈거리
                       b.SubItems.Add(aa);// 토탈거리

                       //  b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");
                       b.SubItems.Add("");

                   }
                   b.BackColor = System.Drawing.Color.LightGray;
                   listView1.Items.Add(b);
                   conn.Close();
                   srRead1.Close();

                   FillList(this.m_list, listView1);

        }

        private void buttonPrint_Click(object sender, EventArgs e)
        {
            PrintPreview_FormInfo();
        }
        public void PrintPreview_FormInfo()
        {
           // m_list.Title = "[(" + openDBInfo.OutTime.ToString() + ") ~ (" + openDBInfo.InTime.ToString() + ")]"
            //            + "\r\n차량: [" + openDBInfo.CarNo + "], 기사번호: [" + openDBInfo.DriverNo + "] 의 (" + openDBInfo.Title + ") 정보";

            //   m_list.Title = "[(" + openDBInfo.OutTime.ToString() + ") ~ (" + openDBInfo.InTime.ToString() + ")]"
            //              + "\r\n차량: [" + formData.CarArea + " " + formData.CarSign + " " + openDBInfo.CarNo + "], 기사번호: [" + openDBInfo.DriverNo + "] 의 (" + openDBInfo.Title + ") 정보";
            //m_list.FitToPage = m_cbFitToPage.Checked;
            m_list.PageSetup();
            if (form1.iButtonMode == false)
            {
                m_list.Title = "[ SHIFT REPORT ]\r\n\r\n";
            }
            else
            {
                m_list.Title = "[ SHIFT REPORT ]\r\n" + label2.Text + "    " + label3.Text + "    " + label1.Text + "\r\n\r\n";
            }
          
            m_list.FitToPage = true;
            m_list.PrintPreview();
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

            int colcnt = 0;

                if(form1.iButtonMode ==false)
                {
                    colcnt =14;
                }
                else
                {
                    colcnt =23;
                }


            try
            {
                for (int i = 0; i < colcnt; i++)
                {


                    ColumnHeader[] col = new ColumnHeader[23];
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

        private void button7_Click(object sender, EventArgs e)
        {
            Excel_print();
           // Excel_Print1();
        }
        public void Excel_Print1()
        {


            string filePath = "c:\\test.xlsx";
                            Excel.Application excelApp = null;

                            object missingType = Type.Missing;

                            excelApp = new Microsoft.Office.Interop.Excel.Application();
                            Excel.Workbook excelBook = excelApp.Workbooks.Add(missingType);
                            Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.Add(missingType, missingType, missingType, missingType);

                            excelApp.Visible = false;


                             /*************************************************************************************************************************************************
                              * 디자인 변경
                              * **********************************************************************************************************************************************/
                             // 컬럼 Width 설정
                             ((Excel.Range)excelWorksheet.get_Range("A:A", Missing.Value)).ColumnWidth = 16.33; //
                             ((Excel.Range)excelWorksheet.get_Range("B:B", Missing.Value)).ColumnWidth = 16.33; // 
                             ((Excel.Range)excelWorksheet.get_Range("C:C", Missing.Value)).ColumnWidth = 16.33; // 
                             ((Excel.Range)excelWorksheet.get_Range("D:D", Missing.Value)).ColumnWidth = 16.33;    // 

                             // 컬럼 정렬 설정
                          //   ((Excel.Range)excelWorksheet.get_Range("B:B, C:C, E:E, F:F, I:I", Missing.Value)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                          //   ((Excel.Range)excelWorksheet.get_Range("B:B, C:C, E:E, I:I", Missing.Value)).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                          //   ((Excel.Range)excelWorksheet.get_Range("A2", "I2")).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                             // 행 높이 설정
                             ((Excel.Range)excelWorksheet.get_Range("1:1", Missing.Value)).RowHeight = 20;

                             // Merge
                          //   ((Excel.Range)excelWorksheet.get_Range("A1", "A1")).Merge(Missing.Value);

                             // 머릿글 설정
                          //   ((Excel.Range)excelWorksheet.get_Range("A1", "A1")).set_Value(Missing.Value, DateTime.Now.ToString("yyyy년 M월 d일") + "   ");
                             ((Excel.Range)excelWorksheet.get_Range("A1", "A1")).set_Value(Missing.Value, DateTime.Now.ToString(this.listView1.Columns[0].Text) );
                             ((Excel.Range)excelWorksheet.get_Range("B1", "B1")).set_Value(Missing.Value, DateTime.Now.ToString(this.listView1.Columns[1].Text) );
                             ((Excel.Range)excelWorksheet.get_Range("C1", "C1")).set_Value(Missing.Value, DateTime.Now.ToString(this.listView1.Columns[2].Text) );
                             ((Excel.Range)excelWorksheet.get_Range("D1", "D1")).set_Value(Missing.Value, DateTime.Now.ToString(this.listView1.Columns[3].Text));
                            
                             // Font Bold 설정
                          //   ((Excel.Range)excelWorksheet.get_Range("A1", "A1")).Font.Bold = true;
                           //  ((Excel.Range)excelWorksheet.get_Range("A2", "J2")).Font.Bold = true;

                             // BackColor 설정
                            // ((Excel.Range)excelWorksheet.get_Range("A1", "J1")).Interior.Color = ColorTranslator.ToOle(Color.Navy);
                            // ((Excel.Range)excelWorksheet.get_Range("A2", "J2")).Interior.Color = ColorTranslator.ToOle(Color.RoyalBlue);

                             // Font Color 설정
                           //  ((Excel.Range)excelWorksheet.get_Range("A1", "J1")).Font.Color = ColorTranslator.ToOle(Color.White);
                           //  ((Excel.Range)excelWorksheet.get_Range("A2", "J2")).Font.Color = ColorTranslator.ToOle(Color.White);

                             // Page Setup
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.LeftHeader = "";
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.CenterHeader = "";
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.RightHeader = "";
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.LeftFooter = "";
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.CenterFooter = "";
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.RightFooter = "";
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.LeftMargin = excelApp.Application.InchesToPoints(0.75);
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.RightMargin = excelApp.Application.InchesToPoints(0.75);
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.TopMargin = excelApp.Application.InchesToPoints(1);
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.BottomMargin = excelApp.Application.InchesToPoints(1);
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.HeaderMargin = excelApp.Application.InchesToPoints(0.5);
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.FooterMargin = excelApp.Application.InchesToPoints(0.5);
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.PrintHeadings = false;
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.PrintGridlines = true;
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.PrintComments = Excel.XlPrintLocation.xlPrintNoComments;
                             //((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.PrintQuality = 600;
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.CenterHorizontally = false;
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.CenterVertically = false;
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.Draft = false;
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
                             //((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.FirstPageNumber = xlAutomatic;
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.Order = Excel.XlOrder.xlDownThenOver;
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.BlackAndWhite = false;
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.Zoom = 60;
                             ((Excel.Worksheet)excelBook.ActiveSheet).PageSetup.PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed;

                      //       excelBook.SaveAs(@saveFileDialog1.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, missingType, missingType, missingType, missingType,
                       //        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missingType, missingType, missingType, missingType, missingType);

                             excelApp.Visible = true;
                            
                           
                             System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                           
      




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
            ((Excel.Range)excelWorksheet.get_Range("Q:Q", Missing.Value)).ColumnWidth = 10.33;
            ((Excel.Range)excelWorksheet.get_Range("R:R", Missing.Value)).ColumnWidth = 10.33;
            ((Excel.Range)excelWorksheet.get_Range("S:S", Missing.Value)).ColumnWidth = 10.33;
            ((Excel.Range)excelWorksheet.get_Range("T:T", Missing.Value)).ColumnWidth = 10.33;
            ((Excel.Range)excelWorksheet.get_Range("U:U", Missing.Value)).ColumnWidth = 10.33;
            ((Excel.Range)excelWorksheet.get_Range("V:V", Missing.Value)).ColumnWidth = 10.33;
            ((Excel.Range)excelWorksheet.get_Range("W:W", Missing.Value)).ColumnWidth = 10.33;
            ((Excel.Range)excelWorksheet.get_Range("X:X", Missing.Value)).ColumnWidth = 10.33;
            ((Excel.Range)excelWorksheet.get_Range("Y:Y", Missing.Value)).ColumnWidth = 10.33;

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
            ((Excel.Range)excelWorksheet.get_Range("Q5", "Q5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("R5", "R5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("S5", "S5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("T5", "T5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("U5", "U5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("V5", "V5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("W5", "W5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("X5", "X5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ((Excel.Range)excelWorksheet.get_Range("Y5", "Y5")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);




            ((Excel.Range)excelWorksheet.get_Range("A:A, B:B, C:C, D:D, E:E,F:F,G:G,H:H,I:I,J:J,K:K,L:L,M:M,N:N,O:O,P:P,Q:Q,R:R,S:S,T:T,U:U,V:V,W:W,X:X,Y:Y", Missing.Value)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            oRng = excel.get_Range("A2", "J2"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.Value2 = label2.Text + "     " + label3.Text + "     " + label1.Text;  //문구 삽입
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


            oRng = excel.get_Range("Q5", "Q5"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("R5", "R5"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("S5", "S5"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("T5", "T5"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("U5", "U5"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("V5", "V5"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("W5", "W5"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("X5", "X5"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;



            oRng = excel.get_Range("Y5", "Y5"); //해당 범위의 셀 획득
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
            string Ncell = "Y";
            Acell += (this.listView1.Items.Count + 5).ToString();
            Ncell += (this.listView1.Items.Count + 5).ToString();
            ((Excel.Range)excelWorksheet.get_Range(Acell, Ncell)).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
          
            System.IO.Directory.CreateDirectory("c:\\Tacho2\\EXCEL");

            string canum = this.listView1.Items[0].SubItems[1].Text;
         //   string FileDate = (this.listView1.Items[0].SubItems[5].Text).Replace("-","") + "_" + (this.listView1.Items[0].SubItems[6].Text).Replace(":","");
            string FileDate = (this.listView1.Items[0].SubItems[5].Text) ;


             dt = Convert.ToDateTime(FileDate);

             FileDate = dt.Day.ToString() + "-" + dt.Month.ToString() + "-" + dt.Year.ToString();

            canum = canum.Trim();

            string savefile = "c:\\Tacho2\\EXCEL\\" + canum + "-" + FileDate;


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

            // excelBook.SaveAs(savefile, Excel.XlFileFormat.xlExcel7);

          excel.Visible = true;
            try
            {

                excelBook.SaveAs(savefile+".xlsx");
            }
            catch (Exception ex)
            {

            }
            excel.Visible = true;
            //excel.Save(filePath);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PrintPreview_FormInfo();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel_print();
        }

    }
}