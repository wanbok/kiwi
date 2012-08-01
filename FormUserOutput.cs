using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms.DataVisualization.Charting;
namespace KIWI
{
    public partial class FormUserOutput : Form
    {
        private FormUserOutput mFormUserOutput;


        private TextBox[] txtOut = null;     //전체
        private TextBox[] txtWOut = null;    //도매
        private TextBox[] txtROut = null;    //소매

        //월평균 판매대수 도매
        private string txtInput4 = "0";
        //월평균 판매대수 계 엑셀 셀 번호
        string ColumnNameInput4 = "F10";
        //월평균 판매대수 소매
        private string txtInput16 = "0";
        //월평균 판매대수 계 엑셀 셀 번호
        string ColumnNameInput16 = "G10";
        //월평균 판매대수 계
        private string txtInput25 = "0";
        //월평균 판매대수 계 엑셀 셀 번호
        string ColumnName_Input25 = "H10";


        public FormUserOutput()
        {
            InitializeComponent();

            txtOut = new TextBox[96] { txtOut1, txtOut2, txtOut3, txtOut4, txtOut5, txtOut6, txtOut7, txtOut8, txtOut9, txtOut10,
            txtOut11, txtOut12, txtOut13, txtOut14, txtOut15, txtOut16, txtOut17, txtOut18, txtOut19, txtOut20,
            txtOut21, txtOut22, txtOut23, txtOut24, txtOut25, txtOut26, txtOut27, txtOut28, txtOut29, txtOut30,
            txtOut31, txtOut32, txtOut33, txtOut34, txtOut35, txtOut36, txtOut37, txtOut38, txtOut39, txtOut40,
            txtOut41, txtOut42, txtOut43, txtOut44, txtOut45, txtOut46, txtOut47, txtOut48, txtOut49, txtOut50,
            txtOut51, txtOut52, txtOut53, txtOut54, txtOut55, txtOut56, txtOut57, txtOut58, txtOut59, txtOut60,
            txtOut61, txtOut62, txtOut63, txtOut64, txtOut65, txtOut66, txtOut67, txtOut68, txtOut69, txtOut70,
            txtOut71, txtOut72, txtOut73, txtOut74, txtOut75, txtOut76, txtOut77, txtOut78, txtOut79, txtOut80,
            txtOut81, txtOut82, txtOut83, txtOut84, txtOut85, txtOut86, txtOut87, txtOut88, txtOut89, txtOut90,
            txtOut91, txtOut92, txtOut93, txtOut94, txtOut95, txtOut96
            };
            txtWOut = new TextBox[84] { txtWOut1, txtWOut2, txtWOut3, txtWOut4, txtWOut5, txtWOut6, txtWOut7, txtWOut8, txtWOut9, txtWOut10,
            txtWOut11, txtWOut12, txtWOut13, txtWOut14, txtWOut15, txtWOut16, txtWOut17, txtWOut18, txtWOut19, txtWOut20,
            txtWOut21, txtWOut22, txtWOut23, txtWOut24, txtWOut25, txtWOut26, txtWOut27, txtWOut28, txtWOut29, txtWOut30,
            txtWOut31, txtWOut32, txtWOut33, txtWOut34, txtWOut35, txtWOut36, txtWOut37, txtWOut38, txtWOut39, txtWOut40,
            txtWOut41, txtWOut42, txtWOut43, txtWOut44, txtWOut45, txtWOut46, txtWOut47, txtWOut48, txtWOut49, txtWOut50,
            txtWOut51, txtWOut52, txtWOut53, txtWOut54, txtWOut55, txtWOut56, txtWOut57, txtWOut58, txtWOut59, txtWOut60,
            txtWOut61, txtWOut62, txtWOut63, txtWOut64, txtWOut65, txtWOut66, txtWOut67, txtWOut68, txtWOut69, txtWOut70,
            txtWOut71, txtWOut72, txtWOut73, txtWOut74, txtWOut75, txtWOut76, txtWOut77, txtWOut78, txtWOut79, txtWOut80,
            txtWOut81, txtWOut82, txtWOut83, txtWOut84
            };
            txtROut = new TextBox[72] { txtROut1, txtROut2, txtROut3, txtROut4, txtROut5, txtROut6, txtROut7, txtROut8, txtROut9, txtROut10,
            txtROut11, txtROut12, txtROut13, txtROut14, txtROut15, txtROut16, txtROut17, txtROut18, txtROut19, txtROut20,
            txtROut21, txtROut22, txtROut23, txtROut24, txtROut25, txtROut26, txtROut27, txtROut28, txtROut29, txtROut30,
            txtROut31, txtROut32, txtROut33, txtROut34, txtROut35, txtROut36, txtROut37, txtROut38, txtROut39, txtROut40,
            txtROut41, txtROut42, txtROut43, txtROut44, txtROut45, txtROut46, txtROut47, txtROut48, txtROut49, txtROut50,
            txtROut51, txtROut52, txtROut53, txtROut54, txtROut55, txtROut56, txtROut57, txtROut58, txtROut59, txtROut60,
            txtROut61, txtROut62, txtROut63, txtROut64, txtROut65, txtROut66, txtROut67, txtROut68, txtROut69, txtROut70,
            txtROut71, txtROut72
            };
            // ReadOnly설정
            for (int i = 0; i < txtOut.Length; i++)
            {

                txtOut[i].ReadOnly = true;
                txtOut[i].BackColor = Color.Silver;

                if (i < txtWOut.Length)
                {
                    txtWOut[i].ReadOnly = true;
                    txtWOut[i].BackColor = Color.Silver;
                }
                if (i < txtROut.Length)
                {
                    txtROut[i].ReadOnly = true;
                    txtROut[i].BackColor = Color.Silver;
                }
            }
            pnlChart.Visible = false;
            //기본입력의 월평균 판매대수 취득
            if (CommonUtil.openAsName != null)
            {
                excel.Worksheet _WorkSheet = CommonUtil.GetExcelWorksheet(CommonUtil.openAsName, 1);
                txtInput4 = CommonUtil.NullToString0(_WorkSheet.get_Range(ColumnNameInput4, Type.Missing).Value2);
                txtInput16 = CommonUtil.NullToString0(_WorkSheet.get_Range(ColumnNameInput16, Type.Missing).Value2);
                txtInput25 = CommonUtil.NullToString0(_WorkSheet.get_Range(ColumnName_Input25, Type.Missing).Value2);
                
                // 결과 취득 클래스

                //전체, 단위당
                setOut(CDataControl.g_ResultBusinessTotal.getArrayOutput전체(), CDataControl.g_ResultBusiness.getArrayOutput전체());
                setWOut(CDataControl.g_ResultStoreTotal.getArrayOutput전체(), CDataControl.g_ResultStore.getArrayOutput전체());
                setROut(CDataControl.g_ResultFutureTotal.getArrayOutput전체(), CDataControl.g_ResultFuture.getArrayOutput전체());

                List<long[]> 전체 = new List<long[]>();
 
                전체.Add(CDataControl.g_ResultBusiness.getArrayOutput전체());
                전체.Add(CDataControl.g_ResultStore.getArrayOutput전체());
                전체.Add(CDataControl.g_ResultFuture.getArrayOutput전체());
                
                OpenChart(chart1, 전체);
                OpenChart(chart2, 전체);
                OpenChart(chart3, 전체);
            }

        }

        public FormUserOutput(FormUserOutput formUserOutput)
        {
            InitializeComponent();
            //더블 버퍼
            this.SetStyle(ControlStyles.DoubleBuffer, true);
            this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            this.SetStyle(ControlStyles.UserPaint, true);

            mFormUserOutput = formUserOutput;
        }

        private void setOut(long[] arrTotal, long[] arr)
        {   
            for (int i = 0; i < 15; i++)
            {   
                txtOut[i].Text = CommonUtil.NullToString0(arrTotal[i]);
                txtOut[i + 16].Text = CommonUtil.NullToString0(arr[i]);
            }
        }

        private void setWOut(long[] arrTotal, long[] arr)
        {
            for (int i = 16; i < 29; i++)
            {   
                txtWOut[i].Text = CommonUtil.NullToString0(arr[i]);
                txtOut[i + 14].Text = CommonUtil.NullToString0(arr[i]);
            }
        }

        private void setROut(long[] arrTotal, long[] arr)
        {
            for (int i = 30; i < 41; i++)
            {
                txtROut[i].Text = CommonUtil.NullToString0(arr[i]);
                txtOut[i + 12].Text = CommonUtil.NullToString0(arr[i]);
            }
        }



        private void OpenChart(Chart chart, List<long[]> lists)
        {
            long[] yValues = null;
            long[] yValues2 = null;
            long[] yValues3 = null;

            string[] xValues = null;

            chart.Series[0].Name = "업계평균";
            chart.Series[1].Name = "당대리점(현재수익)";
            chart.Series[2].Name = "당대리점(미래수익)";
            
            if (chart.Name == "chart1")
            {
                xValues = new string[6] { "누적가입자 수수료", "CS관리수수료", "월단위 업무취급 수수료", "사업자모델 매입에 따른 추가수익", "유통모델 매입에 따른 추가수익(현금+Volume)", "직영매장 판매수익" };

                yValues = new long[6] {lists[0].ToArray()[0], lists[0].ToArray()[1], lists[0].ToArray()[2], lists[0].ToArray()[3], lists[0].ToArray()[4], lists[0].ToArray()[5]};
                yValues2 = new long[6] {lists[1].ToArray()[0], lists[1].ToArray()[1], lists[1].ToArray()[2], lists[1].ToArray()[3], lists[1].ToArray()[4], lists[1].ToArray()[5]};
                yValues3 = new long[6] {lists[2].ToArray()[0], lists[2].ToArray()[1], lists[2].ToArray()[2], lists[2].ToArray()[3], lists[2].ToArray()[4], lists[2].ToArray()[5]};

                chart.Series[0].Points.DataBindXY(xValues, yValues);
                chart.Series[1].Points.DataBindXY(xValues, yValues2);
                chart.Series[2].Points.DataBindXY(xValues, yValues3);
            }
            else if (chart.Name == "chart2")
            {
                xValues = new string[4] { "누적가입자 수수료", "CS관리수수료", "사업자모델 매입에 따른 추가수익", "유통모델 매입에 따른 추가수익(현금+Volume)" };

                yValues = new long[4] { lists[0].ToArray()[16], lists[0].ToArray()[17], lists[0].ToArray()[18], lists[0].ToArray()[19]};
                yValues2 = new long[4] { lists[1].ToArray()[16], lists[1].ToArray()[17], lists[1].ToArray()[18], lists[1].ToArray()[19]};
                yValues3 = new long[4] { lists[2].ToArray()[16], lists[2].ToArray()[17], lists[2].ToArray()[18], lists[2].ToArray()[19]};

                chart.Series[0].Points.DataBindXY(xValues, yValues);
                chart.Series[1].Points.DataBindXY(xValues, yValues2);
                chart.Series[2].Points.DataBindXY(xValues, yValues3);
            }
            else if (chart.Name == "chart3")
            {
                xValues = new string[2] { "월단위 업무취급 수수료", "직영매장 판매수익" };

                yValues = new long[2] { lists[0].ToArray()[30], lists[0].ToArray()[31] };
                yValues2 = new long[2] { lists[1].ToArray()[30], lists[1].ToArray()[31] };
                yValues3 = new long[2] { lists[2].ToArray()[30], lists[2].ToArray()[31] };
                
                chart.Series[0].Points.DataBindXY(xValues, yValues);
                chart.Series[1].Points.DataBindXY(xValues, yValues2);
                chart.Series[2].Points.DataBindXY(xValues, yValues3);
            }
            else if (chart.Name == "chart4")
            {
                xValues = new string[6] { "누적가입자 수수료", "CS관리수수료", "월단위 업무취급 수수료", "사업자모델 매입에 따른 추가수익", "유통모델 매입에 따른 추가수익(현금+Volume)", "직영매장 판매수익" };

                yValues = new long[6] { lists[0].ToArray()[0], lists[0].ToArray()[1], lists[0].ToArray()[2], lists[0].ToArray()[3], lists[0].ToArray()[4], lists[0].ToArray()[5] };
                yValues2 = new long[6] { lists[1].ToArray()[0], lists[1].ToArray()[1], lists[1].ToArray()[2], lists[1].ToArray()[3], lists[1].ToArray()[4], lists[1].ToArray()[5] };
                yValues3 = new long[6] { lists[2].ToArray()[0], lists[2].ToArray()[1], lists[2].ToArray()[2], lists[2].ToArray()[3], lists[2].ToArray()[4], lists[2].ToArray()[5] };

                chart.Series[0].Points.DataBindXY(xValues, yValues);
                chart.Series[1].Points.DataBindXY(xValues, yValues2);
                chart.Series[2].Points.DataBindXY(xValues, yValues3);
            }
            else if (chart.Name == "chart5")
            {
                xValues = new string[4] { "누적가입자 수수료", "CS관리수수료", "사업자모델 매입에 따른 추가수익", "유통모델 매입에 따른 추가수익(현금+Volume)" };

                yValues = new long[4] { lists[0].ToArray()[16], lists[0].ToArray()[17], lists[0].ToArray()[18], lists[0].ToArray()[19] };
                yValues2 = new long[4] { lists[1].ToArray()[16], lists[1].ToArray()[17], lists[1].ToArray()[18], lists[1].ToArray()[19] };
                yValues3 = new long[4] { lists[2].ToArray()[16], lists[2].ToArray()[17], lists[2].ToArray()[18], lists[2].ToArray()[19] };

                chart.Series[0].Points.DataBindXY(xValues, yValues);
                chart.Series[1].Points.DataBindXY(xValues, yValues2);
                chart.Series[2].Points.DataBindXY(xValues, yValues3);
            }
            else if (chart.Name == "chart6")
            {
                xValues = new string[2] { "월단위 업무취급 수수료", "직영매장 판매수익" };

                yValues = new long[2] { lists[0].ToArray()[30], lists[0].ToArray()[31] };
                yValues2 = new long[2] { lists[1].ToArray()[30], lists[1].ToArray()[31] };
                yValues3 = new long[2] { lists[2].ToArray()[30], lists[2].ToArray()[31] };

                chart.Series[0].Points.DataBindXY(xValues, yValues);
                chart.Series[1].Points.DataBindXY(xValues, yValues2);
                chart.Series[2].Points.DataBindXY(xValues, yValues3);
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            pnlChart.Visible = true;
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            pnlChart.Visible = false;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            pnlChart2.Visible = true;
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            pnlChart2.Visible = false;
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            pnlChart3.Visible = true;
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            pnlChart3.Visible = false;
        }


    }
}
