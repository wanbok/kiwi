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
    public partial class FormPopupChartSample : Form
    {
        public FormPopupChartSample()
        {
            InitializeComponent();



        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
            {
                e.Handled = true;
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
            {
                e.Handled = true;
            }

        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
            {
                e.Handled = true;
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
            {
                e.Handled = true;
            }
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
            {
                e.Handled = true;
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
            {
                e.Handled = true;
            }
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
            {
                e.Handled = true;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            chart1.Series[0].Points.Clear();
            chart1.Series[0].Points.Add(Convert.ToInt32(textBox1.Text));
            chart1.Series[0].Points.Add(Convert.ToInt32(textBox2.Text));

            chart1.Series[0].Points[0].LegendText = "월평균 업무취급수수료";
            chart1.Series[0].Points[1].LegendText = "직영매장 판매수익";

            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            chart1.Series[0].Points.Clear();
            chart1.Series[0].Points.Add(Convert.ToInt32(textBox1.Text));
            chart1.Series[0].Points.Add(Convert.ToInt32(textBox2.Text));
            chart1.Series[0].Points[0].LegendText = "월평균 업무취급수수료";
            chart1.Series[0].Points[1].LegendText = "직영매장 판매수익";
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            chart2.Series[0].Points.Clear();
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox3.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox4.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox5.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox6.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox7.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox8.Text));

            chart2.Series[0].Points[0].LegendText = "직원급여(간부급)";
            chart2.Series[0].Points[1].LegendText = "직원급여(평사원)";
            chart2.Series[0].Points[2].LegendText = "지급 임차료";
            chart2.Series[0].Points[3].LegendText = "지급 수수료";
            chart2.Series[0].Points[4].LegendText = "판매촉진비";
            chart2.Series[0].Points[5].LegendText = "건물관리비";

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            chart2.Series[0].Points.Clear();
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox3.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox4.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox5.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox6.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox7.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox8.Text));

            chart2.Series[0].Points[0].LegendText = "직원급여(간부급)";
            chart2.Series[0].Points[1].LegendText = "직원급여(평사원)";
            chart2.Series[0].Points[2].LegendText = "지급 임차료";
            chart2.Series[0].Points[3].LegendText = "지급 수수료";
            chart2.Series[0].Points[4].LegendText = "판매촉진비";
            chart2.Series[0].Points[5].LegendText = "건물관리비";
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            chart2.Series[0].Points.Clear();
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox3.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox4.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox5.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox6.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox7.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox8.Text));

            chart2.Series[0].Points[0].LegendText = "직원급여(간부급)";
            chart2.Series[0].Points[1].LegendText = "직원급여(평사원)";
            chart2.Series[0].Points[2].LegendText = "지급 임차료";
            chart2.Series[0].Points[3].LegendText = "지급 수수료";
            chart2.Series[0].Points[4].LegendText = "판매촉진비";
            chart2.Series[0].Points[5].LegendText = "건물관리비";
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            chart2.Series[0].Points.Clear();
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox3.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox4.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox5.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox6.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox7.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox8.Text));

            chart2.Series[0].Points[0].LegendText = "직원급여(간부급)";
            chart2.Series[0].Points[1].LegendText = "직원급여(평사원)";
            chart2.Series[0].Points[2].LegendText = "지급 임차료";
            chart2.Series[0].Points[3].LegendText = "지급 수수료";
            chart2.Series[0].Points[4].LegendText = "판매촉진비";
            chart2.Series[0].Points[5].LegendText = "건물관리비";
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            chart2.Series[0].Points.Clear();
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox3.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox4.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox5.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox6.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox7.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox8.Text));

            chart2.Series[0].Points[0].LegendText = "직원급여(간부급)";
            chart2.Series[0].Points[1].LegendText = "직원급여(평사원)";
            chart2.Series[0].Points[2].LegendText = "지급 임차료";
            chart2.Series[0].Points[3].LegendText = "지급 수수료";
            chart2.Series[0].Points[4].LegendText = "판매촉진비";
            chart2.Series[0].Points[5].LegendText = "건물관리비";
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            chart2.Series[0].Points.Clear();
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox3.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox4.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox5.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox6.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox7.Text));
            chart2.Series[0].Points.Add(Convert.ToInt32(textBox8.Text));

            chart2.Series[0].Points[0].LegendText = "직원급여(간부급)";
            chart2.Series[0].Points[1].LegendText = "직원급여(평사원)";
            chart2.Series[0].Points[2].LegendText = "지급 임차료";
            chart2.Series[0].Points[3].LegendText = "지급 수수료";
            chart2.Series[0].Points[4].LegendText = "판매촉진비";
            chart2.Series[0].Points[5].LegendText = "건물관리비";
        }
    }
}
