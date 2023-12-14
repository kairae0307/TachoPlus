using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace TachoPlus
{
    public partial class SearchForm : Form
    {
        Form1 form1;
        SearchdataForm searchdataform;
        IncomeForm incomeForm;
        public SearchForm(Form1 f)
        {
            form1 = f;
            InitializeComponent();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            //TimeSpan diff = dateTimePickerEnd.Value - dateTimePickerStart.Value;
                 form1.searchCarNum =textBox1.Text;
                 form1.searchDreverID = textBox3.Text;

                 form1.dtSearchStartDay = dateTimePickerStart.Value;
                 form1.dtSearchEndDay = dateTimePickerEnd.Value;

                 form1.searchCashierID = textBox4.Text;

           
            
                 if (textBox2.Text != "")
                 {
                     form1.SearchIncome = Double.Parse(textBox2.Text);
                 }  
                 else
                 {
                     form1.SearchIncome = 0;
                 }

        

                 if (radioButton1.Checked == true)  // PlateNumber
                 {
                     searchdataform = new SearchdataForm(form1);
                     searchdataform.MdiParent = this.ParentForm;
                     searchdataform.BringToFront();
                     searchdataform.Show();
                     searchdataform.SearchData(form1.searchCarNum, form1.dtSearchStartDay, form1.dtSearchEndDay, 0);
                 }
                 else if (radioButton2.Checked==true)  // DriverID
                 {
                     searchdataform = new SearchdataForm(form1);
                     searchdataform.MdiParent = this.ParentForm;
                     searchdataform.BringToFront();
                     searchdataform.Show();
                     searchdataform.SearchData(form1.searchDreverID, form1.dtSearchStartDay, form1.dtSearchEndDay, 1);
                 }
                 else if (radioButton3.Checked == true) // All 
                 {
                     searchdataform = new SearchdataForm(form1);
                     searchdataform.MdiParent = this.ParentForm;
                     searchdataform.BringToFront();
                     searchdataform.Show();
                     searchdataform.SearchData(form1.searchDreverID, form1.dtSearchStartDay, form1.dtSearchEndDay, 2);
                 }
                 else if (radioButton4.Checked == true) // Income 
                 {
                     incomeForm = new IncomeForm(form1);
                     incomeForm.MdiParent = this.ParentForm;
                     incomeForm.BringToFront();
                     incomeForm.Show();

                     incomeForm.Income_SearchData(form1.SearchIncome, form1.dtSearchStartDay, form1.dtSearchEndDay);

                 


                 }
                 else if (radioButton5.Checked == true) //  cashier
                 {
                     searchdataform = new SearchdataForm(form1);
                     searchdataform.MdiParent = this.ParentForm;
                     searchdataform.BringToFront();
                     searchdataform.Show();
                     searchdataform.SearchData(form1.searchCashierID, form1.dtSearchStartDay, form1.dtSearchEndDay, 3);
                 }
                 else
                 {
                     return;
                 }
        }
    }
}