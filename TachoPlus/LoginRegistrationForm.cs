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

namespace TachoPlus
{
    public partial class LoginRegistrationForm : Form
    {
        public string id = "";
        public string password = "";

        Form1 form1;
        public LoginRegistrationForm(Form1 f)
        {
            form1 = f;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (textBox2.Text != textBox3.Text)
            {
                MessageBox.Show(" The password is incorrect.");
                return;
            }
            string DBstring = "";
            if (form1.CashierMode == true)
            {

                return;
            }
            else
            {
               
                string NameDB = Application.StartupPath + "\\Information.mdb";
                //DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + NameDB;
                DBstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +"c:\\tacho2" + "\\Information.mdb;Jet OLEDB:Database Password=1111";
            }
          
            //					 Db_backup = false;


            OleDbConnection conn = new OleDbConnection(@DBstring);
            conn.Open();
            string queryRead = "";
       

             if (form1.UserLogin == true)
             {
                 queryRead = "select  * from User_Login";
             }
             else if (form1.AdminLogin == true)
             {
                 if (form1.UserIdCreate == true)
                 {
                     queryRead = "select  * from User_Login";
                 }
                 else
                 {
                     queryRead = "select  * from Admin_Login";
                 }
             }

            OleDbCommand commRead = new OleDbCommand(queryRead, conn);
            OleDbDataReader srRead = commRead.ExecuteReader();

            //panel1.BackColor = Color.Wheat;

        
            while (srRead.Read())
            {
                if (srRead.IsDBNull(1) == false)
                {

                    id = srRead.GetString(1);

                    if (id == textBox1.Text)
                    {
                        MessageBox.Show("The ID is joined.");
                        return;
                    }
                }

            }

            string query = "";

            id = textBox1.Text;
            password = textBox2.Text;

            try
            {
                OleDbCommand comm;

                // Fill DB - TblTacho
                if (form1.UserLogin == true)
                {
                    query = "Insert into User_Login ( [ID],[Password]"
                                        + ") values(?,?)";
                   
                }
                else if (form1.AdminLogin == true)
                {
                    if (form1.UserIdCreate == true)
                    {
                        query = "Insert into User_Login ( [ID],[Password]"
                                       + ") values(?,?)";
                    }
                    else
                    {
                        query = "Insert into Admin_Login ( [ID],[Password]"
                                          + ") values(?,?)";
                    }
                }

                comm = new OleDbCommand(query, conn);
                comm.Parameters.Add("ID", OleDbType.Char).Value = id;
                comm.Parameters.Add("Password", OleDbType.Char).Value = password;

                comm.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
               
            
                MessageBox.Show(ex.Message);

            }

            conn.Close();

           
            MessageBox.Show("Successfully registered!", "Result");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox2.Text != textBox3.Text)
            {
                MessageBox.Show(" The password is incorrect.");
                return;
            }
            bool Id_Search = false;
            string DBstring = "";


            if (form1.CashierMode == true)
            {

                return;
            }
            else
            {

                string NameDB = Application.StartupPath + "\\Information.mdb";
                //DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + NameDB;
                DBstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "C:\\tacho2\\" + "Information.mdb;Jet OLEDB:Database Password=1111";
            }
          
            //					 Db_backup = false;
            int DelID = 0;

            OleDbConnection conn = new OleDbConnection(@DBstring);
            conn.Open();
            string queryRead = "";


            if (form1.UserLogin == true)
            {
                queryRead = "select  * from User_Login";
            }
            else if (form1.AdminLogin == true)
            {
                if (form1.UserIdCreate == true)
                {
                    queryRead = "select  * from User_Login";
                }
                else
                {

                    queryRead = "select  * from Admin_Login";
                }
            }

            string cmdStr = "";
            
            try
            {
                if (form1.UserLogin == true)
                {
                    cmdStr = "select count(*) from User_Login where ID='" + textBox1.Text + "'";
                }
                else if (form1.AdminLogin == true)
                {
                    if (form1.UserIdCreate == true)
                    {
                        cmdStr = "select count(*) from User_Login where ID='" + textBox1.Text + "'";
                    }
                    else
                    {
                        cmdStr = "select count(*) from Admin_Login where ID='" + textBox1.Text + "'";
                    }
                }
                
                OleDbCommand Checkuser = new OleDbCommand(cmdStr, conn);
                int temp = Convert.ToInt32(Checkuser.ExecuteScalar().ToString());
                if (temp == 1)
                {
                    string cmdStr2 = "";
                    if (form1.UserLogin == true)
                    {
                        cmdStr2 = "Select Password from User_Login where Password='" + textBox2.Text + "'";
                    }
                    else if (form1.AdminLogin == true)
                    {
                        if (form1.UserIdCreate == true)
                        {
                            cmdStr2 = "Select Password from User_Login where Password='" + textBox2.Text + "'";
                        }
                        else
                        {
                            cmdStr2 = "Select Password from Admin_Login where Password='" + textBox2.Text + "'";
                        }
                    }
                     
                    OleDbCommand pass = new OleDbCommand(cmdStr2, conn);
                    string password = pass.ExecuteScalar().ToString();
                 

                    if (password == textBox2.Text)
                    {
                        Id_Search = true;
                   
                    }
                    else
                    {
                      
                        MessageBox.Show(" The password is incorrect."); 
                    }
                }
                else
                {
                   
                    MessageBox.Show(" The id is incorrect.");
                }
            }
            catch (Exception exx)
            {
              
                MessageBox.Show(" The password is incorrect."); 
                    
            }



             string queryDel ="";
             id = textBox1.Text;
            if (Id_Search == true)
            {
                if (form1.UserLogin == true)
                {
                    queryDel = " DELETE FROM User_Login WHERE [ID] =?";
                }
                else if (form1.AdminLogin == true)
                {
                    if (form1.UserIdCreate == true)
                    {
                        queryDel = " DELETE FROM User_Login WHERE [ID] =?";
                    }
                    else
                    {
                        queryDel = " DELETE FROM Admin_Login WHERE [ID] =?";
                    }
                }



                OleDbCommand My_Command = new OleDbCommand(queryDel, conn);
                    My_Command.Parameters.Add("@ID", textBox1.Text);
                    My_Command.ExecuteNonQuery();

                    conn.Close();
               
           
                MessageBox.Show("Successfully deleted!", "Result");
            }
           

        }
    }
}
