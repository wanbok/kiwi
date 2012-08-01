using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace KIWI
{
    public partial class FormUserInfo : Form
    {
        public enum FORMS {FORMUSERINFO = 0, FORMUSERINPUT, FORMUSEROUTPUT, FORMSIMULATEINPUT, FORMSIMULATE, FORMANALYSIS };

        private FormUserInput mFormUserInput;
        private FormUserOutput mFormUserOutput;
        private FormUserSimulateInput mFormUserSimulateInput;
        private FormUserSimulate mFormUserSimulate;
        private FormAdmin mFormAdmin;
        private FormUserAnalysis mFormUserAnalysis;

        private System.Data.OleDb.OleDbConnection OLEcon;
        private System.Data.OleDb.OleDbDataAdapter Adot;
        private DataSet ds = new DataSet();
        private DataTable ExcelTable = new DataTable();

        public FormUserInfo()
        {
            InitializeComponent();
            /*
            Adot = new System.Data.OleDb.OleDbDataAdapter("Select * From [" + ExcelTable.Rows[0][2].ToString().Trim() + "]", OLEcon); 
            try { 
                Adot.Fill(ds); 
            } 
            catch (Exception ep) { 
                MessageBox.Show(ep.Message, "DataSet Error"); 
                return; 
            }
            */

        }

        private void FormUserInfo_Load(object sender, EventArgs e)
        {
            
            mFormUserOutput = new FormUserOutput(this);
            mFormUserSimulateInput = new FormUserSimulateInput(this);
            mFormUserSimulate = new FormUserSimulate(this);
            //mFormAdmin = new FormAdmin(this);
            mFormUserAnalysis = new FormUserAnalysis(this);
        }
        
        private void buttonStart_Click(object sender, EventArgs e)
        {
            Navigation((int)FORMS.FORMUSERINPUT);
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        public void Navigation(int childform)
        {
            switch (childform)
            {
                case (int)FORMS.FORMUSERINFO:
                    break;
                case (int)FORMS.FORMUSERINPUT:
                    
                    
                    mFormUserInput.Show();
                    
                    break;
                case (int)FORMS.FORMUSEROUTPUT:
                    mFormUserInput.Close();
                    
                    mFormUserOutput.Show();
                    
                    break;
                case (int)FORMS.FORMSIMULATEINPUT:
                    mFormUserOutput.Close();
                    mFormUserInput.Close();
                    mFormUserSimulateInput.Show();
                    break;
                case (int)FORMS.FORMSIMULATE:
                    mFormUserSimulate.ShowDialog();
                    break;
                case (int)FORMS.FORMANALYSIS:
                    mFormUserAnalysis.ShowDialog();
                    break;
                case 6:
                    break;
                case 7:
                    break;
                case 8:
                    break;
                case 9:
                    break;

            }
        }

    }
}
