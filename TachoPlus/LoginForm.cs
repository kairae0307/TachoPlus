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
    public partial class LoginForm : Form
    {
        public string ID = "";
        public string Password = "";
        Form1 form1;
        public LoginForm(Form1 f)
        {
            form1 = f;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            string formname = "";
            int dotval = 0;
            int Cnt = 0;
            if (checkBox1.Checked == true)
            {
                form1.AdminLogin = true;
                form1.UserLogin = false;
            }
            else
            {
                form1.AdminLogin = false;
                form1.UserLogin = true;
            }

          
            try
            {
                string DBstring = "";

                if (form1.CashierMode == true)
                {
                  
                   // string NameDB = @"\\" + form1.ShareIP + "\\" + Application.StartupPath+ "\\Information.mdb";
                    //DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + NameDB;
                    DBstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"\\" + form1.ShareIP + "\\tacho2" + "\\Information.mdb;Jet OLEDB:Database Password=1111";
                }
                else
                {
                   // string NameDB = Application.StartupPath + "\\Information.mdb";
                    //DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + NameDB;
                    DBstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + form1.TACHO2_path + "\\Information.mdb;Jet OLEDB:Database Password=1111";
                }

             
                //					 Db_backup = false;

                OleDbConnection con = new OleDbConnection(@DBstring);
                //  OleDbConnection con = new OleDbConnection(ConfigurationManager.ConnectionStrings["CWsystem.mdb"].ConnectionString);
                con.Open();
                string cmdStr = "";
                if (form1.UserLogin == true)
                {
                    cmdStr = "select count(*) from User_Login where ID='" + textBox1.Text + "'";
                }
                else if (form1.AdminLogin == true)
                {
                    cmdStr = "select count(*) from Admin_Login where ID='" + textBox1.Text + "'";
                }
                
                OleDbCommand Checkuser = new OleDbCommand(cmdStr, con);
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
                        cmdStr2 = "Select Password from Admin_Login where Password='" + textBox2.Text + "'";
                    }
                     
                    OleDbCommand pass = new OleDbCommand(cmdStr2, con);
                    string password = pass.ExecuteScalar().ToString();
                    con.Close();

                    if (password == textBox2.Text)
                    {
                        form1.LoginOK = true;
                        form1.label26.Text = textBox1.Text;
                       if( form1.iButtonMode ==true)
                        {
                           form1.CashierID = textBox1.Text;
                          form1.Visible = true;
                          form1.Ibuttonthread = new Thread(new ThreadStart(form1.Run_Ibutton));
                          form1.Ibuttonthread.IsBackground = true;
                          Thread.Sleep(100);
                          form1.Ibuttonthread.Start();
                       
                        }

                         this.DialogResult = DialogResult.OK;
                         return;
                    }
                    else
                    {
                        form1.Visible = false;
                        MessageBox.Show(" The password is incorrect."); 
                    }
                }
                else
                {
                    form1.Visible = false;
                    MessageBox.Show(" The id is incorrect.");
                }
            }
            catch (Exception ex)
            {
                form1.Visible = false;
                MessageBox.Show(" The password is incorrect."); 
                    
            }
        
       
        }

        private void LoginForm_Load(object sender, EventArgs e)
        {
           
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                form1.AdminLogin = true;
                form1.UserLogin = false;
            }
            else
            {
                form1.AdminLogin = false;
                form1.UserLogin = true;
            }
        }

        private void LoginForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                button1_Click(sender, e);

        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                button1_Click(sender, e);
        }

        private void LoginForm_FormClosed(object sender, FormClosedEventArgs e)
        {

            if (form1.LoginOK == false)
            {
                Application.Exit();
            }
        
        }

     

       
    }
}
