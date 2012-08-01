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
    public partial class FormOutput : Form
    {
        

        public FormOutput()
        {
            InitializeComponent();
            chart1.Visible = true;
            chart2.Visible = true;
            chart3.Visible = true;
            chart4.Visible = true;
            chart5.Visible = true;
            chart6.Visible = true;
            chart7.Visible = false;
            chart8.Visible = false;
    
        }

        public FormOutput(CInputData data)
        {
            InitializeComponent();
            chart3.Series[0].Points[0].SetValueY(data.getData1());
            chart3.Series[0].Points[1].SetValueY(data.getData2());
            chart3.Series[0].Points[2].SetValueY(data.getData3());
            chart3.Series[0].Points[3].SetValueY(data.getData4());
            chart3.Series[0].Points[4].SetValueY(data.getData5());
            chart1.Visible = true;
            chart2.Visible = true;
            chart3.Visible = true;
            chart4.Visible = true;
            chart5.Visible = true;
            chart6.Visible = true;
            chart7.Visible = false;
            chart8.Visible = false;
        }

        //전체 차트1을 더블클릭하여
        //확대된 차트와 데이터를 보여준다
        private void chart1_DoubleClick(object sender, EventArgs e)
        {
            FormPopupChart form = new FormPopupChart();
            form.Show();
        }

        private void chart3_DoubleClick(object sender, EventArgs e)
        {
            FormPopupChart form = new FormPopupChart();
            form.Show();
        }

        private void chart5_DoubleClick(object sender, EventArgs e)
        {
            FormPopupChart form = new FormPopupChart();
            form.Show();
        }

        private void chart2_DoubleClick(object sender, EventArgs e)
        {
            FormPopupChart form = new FormPopupChart();
            form.Show();
        }

        private void chart4_DoubleClick(object sender, EventArgs e)
        {
            FormPopupChart form = new FormPopupChart();
            form.Show();

        }

        private void chart6_DoubleClick(object sender, EventArgs e)
        {
            FormPopupChart form = new FormPopupChart();
            form.Show();

        }

        //check 전체
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }





        private void chart7_DoubleClick(object sender, EventArgs e)
        {
            FormPopupChartBar form = new FormPopupChartBar();
            form.Show();
        }

        private void chart8_DoubleClick(object sender, EventArgs e)
        {
            FormPopupChartBar2 form = new FormPopupChartBar2();
            form.Show();

        }

        private void checkBox18_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox18.Checked == true)
            {
                //bar
                chart1.Visible = false;
                chart2.Visible = false;
                chart3.Visible = false;
                chart4.Visible = false;
                chart5.Visible = false;
                chart6.Visible = false;
                chart7.Visible = true;
                chart8.Visible = true;

                checkBox17.Checked = false;
            }
        }

        private void checkBox17_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox17.Checked == true)
            {
                //pie
                chart1.Visible = true;
                chart2.Visible = true;
                chart3.Visible = true;
                chart4.Visible = true;
                chart5.Visible = true;
                chart6.Visible = true;
                chart7.Visible = false;
                chart8.Visible = false;

                checkBox18.Checked = false;
            }
        }


        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox1.Checked == false )
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
            }
            else if (checkBox1.Checked == true)
            {
                checkBox2.Checked = true;
                checkBox3.Checked = true;
            }
        }


        //도매
        private void checkBox2_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox1.Checked == false && checkBox2.Checked == true)
            {
                checkBox3.Checked = false;
                FormPopupWholeSale form = new FormPopupWholeSale();
                form.Show();
            }
        }

        //소매
        private void checkBox3_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox1.Checked == false && checkBox3.Checked == true)
            {
                checkBox2.Checked = false;
                FormPopupRetail form = new FormPopupRetail();
                form.Show();

            }
        }

        private void ToolStripMenuItemFTPSetting_Click(object sender, EventArgs e)
        {
//            FormFTPSetting form = new FormFTPSetting();
//            form.Show();
        }




    }
}
