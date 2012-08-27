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
using System.IO;

namespace KIWI
{
    public partial class FormUserSimulateOutput : Form
    {
        private TextBox[] txtOut = null;     //전체
        private TextBox[] txtWOut = null;    //도매
        private TextBox[] txtROut = null;    //소매

        private Label[] lblTitle = null;

        private Double[] existingData = null;
        private Double[] existingWData = null;
        private Double[] existingRData = null;

        private Double[] simulData = null;
        private Double[] simulWData = null;
        private Double[] simulRData = null;

        private Label[] labelWon = null;
        private Label[] labelWon2 = null;
        private Label[] labelWon3 = null;

        private string[] names = new string[6] { "업계평균", "당대리점(현재수익)", "당대리점(미래수익)", "시뮬레이션-업계평균", "시뮬레이션-당대리점(현재수익)", "시뮬레이션-당대리점(미래수익)" };
        private List<string[]> selectedData = new List<string[]>();

        private int currentIndex = 0;

        public FormUserSimulateOutput()
        {
            InitializeComponent();

            labelWon = new Label[6] { label88, label89, label91, label92, label93, label96 };
            labelWon2 = new Label[6] { label106, label101, label100, label99, label98, label97 };
            labelWon3 = new Label[6] { label114, label113, label112, label107, label108, label111 };

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
                //txtOut[i].BackColor = Color.White;
                txtOut[i].BorderStyle = BorderStyle.None;
                txtOut[i].TextChanged += new System.EventHandler(addComma_TextChanged);

                if (i < txtWOut.Length)
                {
                    txtWOut[i].ReadOnly = true;
                    //txtWOut[i].BackColor = Color.White;
                    txtWOut[i].BorderStyle = BorderStyle.None;
                    txtWOut[i].TextChanged += new System.EventHandler(addComma_TextChanged);
                }
                if (i < txtROut.Length)
                {
                    txtROut[i].ReadOnly = true;
                    //txtROut[i].BackColor = Color.White;
                    txtROut[i].BorderStyle = BorderStyle.None;
                    txtROut[i].TextChanged += new System.EventHandler(addComma_TextChanged);
                }
            }

            lblTitle = new Label[] { lblTitle1, lblTitle2, lblTitle3, lblTitle4, lblTitle5, lblTitle6, lblTitle7, lblTitle8, lblTitle9 };

            checkBox1.CheckedChanged += new EventHandler(checkboxes);
            checkBox2.CheckedChanged += new EventHandler(checkboxes);
            checkBox3.CheckedChanged += new EventHandler(checkboxes);
            checkBox4.CheckedChanged += new EventHandler(checkboxes);
            checkBox5.CheckedChanged += new EventHandler(checkboxes);
            checkBox6.CheckedChanged += new EventHandler(checkboxes);

            existingData = new Double[96];
            existingWData = new Double[84];
            existingRData = new Double[72];

            simulData = new Double[96];
            simulWData = new Double[84];
            simulRData = new Double[72];

            pnlChart.Visible = false;
            if (CommonUtil.isLoadedDataFromFile)
            {
                radio기존업계평균.Enabled = true;
                radio기존업계평균.Checked = CommonUtil.isSelectExistData;
                applyData(CommonUtil.isSelectExistData);
            }
            else
            {
                radio기존업계평균.Enabled = false;
                radio기존업계평균.Checked = false;
                applyData(false);
            }

        }

        public void applyData(bool isFile) {
            // 정보적용
            setOut(isFile);
            
            // 시뮬레이션 정보적용
            setSimulatorOut();

            checkboxes(null, null); // 데이터를 적용시키기위한 트리거 용도
        }

        private void setOut(bool isFile)
        {
            if (CDataControl.g_ResultBusinessTotal == null ||
                CDataControl.g_ResultBusiness == null ||
                CDataControl.g_ResultStoreTotal == null ||
                CDataControl.g_ResultStore == null ||
                CDataControl.g_ResultFutureTotal == null ||
                CDataControl.g_ResultFuture == null) return;

            if (isFile)
            {
                // 전체
                for (int i = 0, n = 16; i < n; i++)
                {
                    existingData[i] = CDataControl.g_FileResultBusinessTotal.getArr전체_수익_비용_및_계산포함()[i];
                    existingData[i + n] = CDataControl.g_FileResultBusiness.getArr전체_수익_비용_및_계산포함()[i];
                    existingData[i + n * 2] = CDataControl.g_FileResultStoreTotal.getArr전체_수익_비용_및_계산포함()[i];
                    existingData[i + n * 3] = CDataControl.g_FileResultStore.getArr전체_수익_비용_및_계산포함()[i];
                    existingData[i + n * 4] = CDataControl.g_FileResultFutureTotal.getArr전체_수익_비용_및_계산포함()[i];
                    existingData[i + n * 5] = CDataControl.g_FileResultFuture.getArr전체_수익_비용_및_계산포함()[i];
                }
                // 도매
                for (int i = 0, n = 14; i < n; i++)
                {
                    existingWData[i] = CDataControl.g_FileResultBusinessTotal.getArr도매_수익_비용_및_계산포함()[i];
                    existingWData[i + n] = CDataControl.g_FileResultBusiness.getArr도매_수익_비용_및_계산포함()[i];
                    existingWData[i + n * 2] = CDataControl.g_FileResultStoreTotal.getArr도매_수익_비용_및_계산포함()[i];
                    existingWData[i + n * 3] = CDataControl.g_FileResultStore.getArr도매_수익_비용_및_계산포함()[i];
                    existingWData[i + n * 4] = CDataControl.g_FileResultFutureTotal.getArr도매_수익_비용_및_계산포함()[i];
                    existingWData[i + n * 5] = CDataControl.g_FileResultFuture.getArr도매_수익_비용_및_계산포함()[i];
                }
                // 소매
                for (int i = 0, n = 12; i < n; i++)
                {
                    existingRData[i] = CDataControl.g_FileResultBusinessTotal.getArr소매_수익_비용_및_계산포함()[i];
                    existingRData[i + n] = CDataControl.g_FileResultBusiness.getArr소매_수익_비용_및_계산포함()[i];
                    existingRData[i + n * 2] = CDataControl.g_FileResultStoreTotal.getArr소매_수익_비용_및_계산포함()[i];
                    existingRData[i + n * 3] = CDataControl.g_FileResultStore.getArr소매_수익_비용_및_계산포함()[i];
                    existingRData[i + n * 4] = CDataControl.g_FileResultFutureTotal.getArr소매_수익_비용_및_계산포함()[i];
                    existingRData[i + n * 5] = CDataControl.g_FileResultFuture.getArr소매_수익_비용_및_계산포함()[i];
                }
            }
            else
            {
                // 전체
                for (int i = 0, n = 16; i < n; i++)
                {
                    existingData[i] = CDataControl.g_ResultBusinessTotal.getArr전체_수익_비용_및_계산포함()[i];
                    existingData[i + n] = CDataControl.g_ResultBusiness.getArr전체_수익_비용_및_계산포함()[i];
                    existingData[i + n * 2] = CDataControl.g_ResultStoreTotal.getArr전체_수익_비용_및_계산포함()[i];
                    existingData[i + n * 3] = CDataControl.g_ResultStore.getArr전체_수익_비용_및_계산포함()[i];
                    existingData[i + n * 4] = CDataControl.g_ResultFutureTotal.getArr전체_수익_비용_및_계산포함()[i];
                    existingData[i + n * 5] = CDataControl.g_ResultFuture.getArr전체_수익_비용_및_계산포함()[i];
                }
                // 도매
                for (int i = 0, n = 14; i < n; i++)
                {
                    existingWData[i] = CDataControl.g_ResultBusinessTotal.getArr도매_수익_비용_및_계산포함()[i];
                    existingWData[i + n] = CDataControl.g_ResultBusiness.getArr도매_수익_비용_및_계산포함()[i];
                    existingWData[i + n * 2] = CDataControl.g_ResultStoreTotal.getArr도매_수익_비용_및_계산포함()[i];
                    existingWData[i + n * 3] = CDataControl.g_ResultStore.getArr도매_수익_비용_및_계산포함()[i];
                    existingWData[i + n * 4] = CDataControl.g_ResultFutureTotal.getArr도매_수익_비용_및_계산포함()[i];
                    existingWData[i + n * 5] = CDataControl.g_ResultFuture.getArr도매_수익_비용_및_계산포함()[i];
                }
                // 소매
                for (int i = 0, n = 12; i < n; i++)
                {
                    existingRData[i] = CDataControl.g_ResultBusinessTotal.getArr소매_수익_비용_및_계산포함()[i];
                    existingRData[i + n] = CDataControl.g_ResultBusiness.getArr소매_수익_비용_및_계산포함()[i];
                    existingRData[i + n * 2] = CDataControl.g_ResultStoreTotal.getArr소매_수익_비용_및_계산포함()[i];
                    existingRData[i + n * 3] = CDataControl.g_ResultStore.getArr소매_수익_비용_및_계산포함()[i];
                    existingRData[i + n * 4] = CDataControl.g_ResultFutureTotal.getArr소매_수익_비용_및_계산포함()[i];
                    existingRData[i + n * 5] = CDataControl.g_ResultFuture.getArr소매_수익_비용_및_계산포함()[i];
                }
            }
        }

        private void setSimulatorOut()
        {
            if (CDataControl.g_SimResultBusinessTotal == null ||
                CDataControl.g_SimResultBusiness == null ||
                CDataControl.g_SimResultStoreTotal == null ||
                CDataControl.g_SimResultStore == null ||
                CDataControl.g_SimResultFutureTotal == null ||
                CDataControl.g_SimResultFuture == null) return;

            // 전체
            for (int i = 0, n = 16; i < n; i++)
            {
                simulData[i] = CDataControl.g_SimResultBusinessTotal.getArr전체_수익_비용_및_계산포함()[i];
                simulData[i + n] = CDataControl.g_SimResultBusiness.getArr전체_수익_비용_및_계산포함()[i];
                simulData[i + n * 2] = CDataControl.g_SimResultStoreTotal.getArr전체_수익_비용_및_계산포함()[i];
                simulData[i + n * 3] = CDataControl.g_SimResultStore.getArr전체_수익_비용_및_계산포함()[i];
                simulData[i + n * 4] = CDataControl.g_SimResultFutureTotal.getArr전체_수익_비용_및_계산포함()[i];
                simulData[i + n * 5] = CDataControl.g_SimResultFuture.getArr전체_수익_비용_및_계산포함()[i];
            }
            // 도매
            for (int i = 0, n = 14; i < n; i++)
            {
                simulWData[i] = CDataControl.g_SimResultBusinessTotal.getArr도매_수익_비용_및_계산포함()[i];
                simulWData[i + n] = CDataControl.g_SimResultBusiness.getArr도매_수익_비용_및_계산포함()[i];
                simulWData[i + n * 2] = CDataControl.g_SimResultStoreTotal.getArr도매_수익_비용_및_계산포함()[i];
                simulWData[i + n * 3] = CDataControl.g_SimResultStore.getArr도매_수익_비용_및_계산포함()[i];
                simulWData[i + n * 4] = CDataControl.g_SimResultFutureTotal.getArr도매_수익_비용_및_계산포함()[i];
                simulWData[i + n * 5] = CDataControl.g_SimResultFuture.getArr도매_수익_비용_및_계산포함()[i];
            }
            // 소매
            for (int i = 0, n = 12; i < n; i++)
            {
                simulRData[i] = CDataControl.g_SimResultBusinessTotal.getArr소매_수익_비용_및_계산포함()[i];
                simulRData[i + n] = CDataControl.g_SimResultBusiness.getArr소매_수익_비용_및_계산포함()[i];
                simulRData[i + n * 2] = CDataControl.g_SimResultStoreTotal.getArr소매_수익_비용_및_계산포함()[i];
                simulRData[i + n * 3] = CDataControl.g_SimResultStore.getArr소매_수익_비용_및_계산포함()[i];
                simulRData[i + n * 4] = CDataControl.g_SimResultFutureTotal.getArr소매_수익_비용_및_계산포함()[i];
                simulRData[i + n * 5] = CDataControl.g_SimResultFuture.getArr소매_수익_비용_및_계산포함()[i];
            }
        }

        private void OpenChart(int[] indexes)
        {
            Chart[] charts = new Chart[] {chart1, chart2, chart3, chart4, chart5, chart6};
            double[] yValues = null;
            double[] yValues2 = null;
            double[] yValues3 = null;

            string[] xValues = null;

            for (int j = 0; j < 3; j++)
            {
                for (int i = 0; i < 2; i++)
                {
                    Chart chart = charts[j + i * 3];
                    chart.Series[0].Name = " ";
                    chart.Series[1].Name = "  ";
                    chart.Series[2].Name = "   ";
                    chart.Series[0].Name = indexes[0] < 0 ? "    " : names[indexes[0]];
                    chart.Series[1].Name = indexes[1] < 0 ? "     " : names[indexes[1]];
                    chart.Series[2].Name = indexes[2] < 0 ? "      " : names[indexes[2]];

                    if (chart.Name == "chart1")
                    {
                        xValues = new string[6] { "누적가입자 수수료", "CS관리수수료", "월단위 업무취급 수수료", "사업자모델 매입에 따른 추가수익", "유통모델 매입에 따른 추가수익(현금+Volume)", "직영매장 판매수익" };

                        yValues = new double[6]{ Convert.ToDouble(txtOut17.Text), Convert.ToDouble(txtOut18.Text), Convert.ToDouble(txtOut19.Text), 
                            Convert.ToDouble(txtOut20.Text), Convert.ToDouble(txtOut21.Text), Convert.ToDouble(txtOut22.Text) };

                        yValues2 = new double[6]{ Convert.ToDouble(txtOut49.Text), Convert.ToDouble(txtOut50.Text), Convert.ToDouble(txtOut51.Text), 
                            Convert.ToDouble(txtOut52.Text), Convert.ToDouble(txtOut53.Text), Convert.ToDouble(txtOut54.Text) };

                        yValues3 = new double[6]{ Convert.ToDouble(txtOut81.Text), Convert.ToDouble(txtOut82.Text), Convert.ToDouble(txtOut83.Text), 
                            Convert.ToDouble(txtOut84.Text), Convert.ToDouble(txtOut85.Text), Convert.ToDouble(txtOut86.Text) };

                        chart.Series[0].Points.DataBindXY(xValues, yValues);
                        chart.Series[1].Points.DataBindXY(xValues, yValues2);
                        chart.Series[2].Points.DataBindXY(xValues, yValues3);
                    }
                    else if (chart.Name == "chart2")
                    {
                        xValues = new string[4] { "누적가입자 수수료", "CS관리수수료", "사업자모델 매입에 따른 추가수익", "유통모델 매입에 따른 추가수익(현금+Volume)" };

                        yValues = new double[4]{ Convert.ToDouble(txtWOut15.Text), Convert.ToDouble(txtWOut16.Text), Convert.ToDouble(txtWOut17.Text), 
                            Convert.ToDouble(txtWOut18.Text) };

                        yValues2 = new double[4]{ Convert.ToDouble(txtWOut43.Text), Convert.ToDouble(txtWOut44.Text), Convert.ToDouble(txtWOut45.Text), 
                            Convert.ToDouble(txtWOut46.Text) };

                        yValues3 = new double[4]{ Convert.ToDouble(txtWOut71.Text), Convert.ToDouble(txtWOut72.Text), Convert.ToDouble(txtWOut73.Text), 
                            Convert.ToDouble(txtWOut74.Text) };

                        chart.Series[0].Points.DataBindXY(xValues, yValues);
                        chart.Series[1].Points.DataBindXY(xValues, yValues2);
                        chart.Series[2].Points.DataBindXY(xValues, yValues3);
                    }
                    else if (chart.Name == "chart3")
                    {
                        xValues = new string[2] { "월단위 업무취급 수수료", "직영매장 판매수익" };

                        yValues = new double[2] { Convert.ToDouble(txtROut13.Text), Convert.ToDouble(txtROut14.Text) };

                        yValues2 = new double[2] { Convert.ToDouble(txtROut37.Text), Convert.ToDouble(txtROut38.Text) };

                        yValues3 = new double[2] { Convert.ToDouble(txtROut61.Text), Convert.ToDouble(txtROut62.Text) };

                        chart.Series[0].Points.DataBindXY(xValues, yValues);
                        chart.Series[1].Points.DataBindXY(xValues, yValues2);
                        chart.Series[2].Points.DataBindXY(xValues, yValues3);
                    }
                    else if (chart.Name == "chart4")
                    {
                        xValues = new string[7] { "대리점 투자비용", "인건비(급여,복리후생비)", "임차료", "이자비용", "부가세", "법인세", "기타관리비용" };

                        yValues = new double[7]{ Convert.ToDouble(txtOut24.Text), Convert.ToDouble(txtOut25.Text), Convert.ToDouble(txtOut26.Text), 
                            Convert.ToDouble(txtOut27.Text), Convert.ToDouble(txtOut28.Text), Convert.ToDouble(txtOut29.Text), Convert.ToDouble(txtOut30.Text) };

                        yValues2 = new double[7]{ Convert.ToDouble(txtOut56.Text), Convert.ToDouble(txtOut57.Text), Convert.ToDouble(txtOut58.Text), 
                            Convert.ToDouble(txtOut59.Text), Convert.ToDouble(txtOut60.Text), Convert.ToDouble(txtOut61.Text), Convert.ToDouble(txtOut62.Text) };

                        yValues3 = new double[7]{ Convert.ToDouble(txtOut88.Text), Convert.ToDouble(txtOut89.Text), Convert.ToDouble(txtOut90.Text), 
                            Convert.ToDouble(txtOut91.Text), Convert.ToDouble(txtOut92.Text), Convert.ToDouble(txtOut93.Text), Convert.ToDouble(txtOut94.Text) };

                        chart.Series[0].Points.DataBindXY(xValues, yValues);
                        chart.Series[1].Points.DataBindXY(xValues, yValues2);
                        chart.Series[2].Points.DataBindXY(xValues, yValues3);
                    }
                    else if (chart.Name == "chart5")
                    {
                        xValues = new string[7] { "대리점 투자비용", "인건비(급여,복리후생비)", "임차료", "이자비용", "부가세", "법인세", "기타관리비용" };

                        yValues = new double[7]{ Convert.ToDouble(txtWOut20.Text), Convert.ToDouble(txtWOut21.Text), Convert.ToDouble(txtWOut22.Text),
                        Convert.ToDouble(txtWOut23.Text), Convert.ToDouble(txtWOut24.Text), Convert.ToDouble(txtWOut25.Text), Convert.ToDouble(txtWOut26.Text) };

                        yValues2 = new double[7]{ Convert.ToDouble(txtWOut48.Text), Convert.ToDouble(txtWOut49.Text), Convert.ToDouble(txtWOut50.Text), 
                            Convert.ToDouble(txtWOut51.Text), Convert.ToDouble(txtWOut52.Text), Convert.ToDouble(txtWOut53.Text), Convert.ToDouble(txtWOut54.Text) };

                        yValues3 = new double[7]{ Convert.ToDouble(txtWOut76.Text), Convert.ToDouble(txtWOut77.Text), Convert.ToDouble(txtWOut78.Text), 
                            Convert.ToDouble(txtWOut79.Text), Convert.ToDouble(txtWOut80.Text), Convert.ToDouble(txtWOut81.Text), Convert.ToDouble(txtWOut82.Text) };

                        chart.Series[0].Points.DataBindXY(xValues, yValues);
                        chart.Series[1].Points.DataBindXY(xValues, yValues2);
                        chart.Series[2].Points.DataBindXY(xValues, yValues3);
                    }
                    else if (chart.Name == "chart6")
                    {
                        xValues = new string[6] { "인건비(급여,복리후생비)", "임차료", "이자비용", "부가세", "법인세", "기타관리비용" };

                        yValues = new double[6] { Convert.ToDouble(txtROut16.Text), Convert.ToDouble(txtROut17.Text), Convert.ToDouble(txtROut18.Text),
                        Convert.ToDouble(txtROut19.Text), Convert.ToDouble(txtROut20.Text), Convert.ToDouble(txtROut21.Text) };

                        yValues2 = new double[6] { Convert.ToDouble(txtROut40.Text), Convert.ToDouble(txtROut41.Text), Convert.ToDouble(txtROut42.Text),
                        Convert.ToDouble(txtROut43.Text), Convert.ToDouble(txtROut44.Text), Convert.ToDouble(txtROut45.Text)};

                        yValues3 = new double[6] { Convert.ToDouble(txtROut64.Text), Convert.ToDouble(txtROut65.Text), Convert.ToDouble(txtROut66.Text),
                        Convert.ToDouble(txtROut67.Text), Convert.ToDouble(txtROut68.Text), Convert.ToDouble(txtROut69.Text)};

                        chart.Series[0].Points.DataBindXY(xValues, yValues);
                        chart.Series[1].Points.DataBindXY(xValues, yValues2);
                        chart.Series[2].Points.DataBindXY(xValues, yValues3);
                    }
                    chart.ChartAreas[0].AxisX.LabelStyle.Angle = 0;
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            FormUserSimulateInput frm = new FormUserSimulateInput(this);
            frm.ShowDialog();
        }

        private void checkboxes(object sender, EventArgs e)
        {
            CheckBox[] chks = new CheckBox[6] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6 };

            int cnt = 0;
            int[] indexes = new int[]{-1, -1, -1};
            int indexesCounter = 0;

            for (int i = 0; i < chks.Length; i++)
            {
                if (chks[i].Checked)
                {
                    cnt++;
                }
            }

            if (cnt > 3)
            {
                MessageBox.Show("체크는 3개만 됩니다.");
                if (sender != null) (sender as CheckBox).Checked = false;
            }

            currentIndex = 0;
            for (int i = 0; i < chks.Length; i++)
            {
                if (chks[i].Checked)
                {
                    setDataAtIndex(i);
                    indexes[indexesCounter++] = i;
                }
            }
            for (int i = currentIndex; i < 3; i++) {
                setDataToEmpty();
                indexes[indexesCounter++] = -1;
            }
            
            OpenChart(indexes);
        }

        private void setDataAtIndex(int index)
        {
            lblTitle[currentIndex].Text = names[index];
            lblTitle[currentIndex+3].Text = names[index];
            lblTitle[currentIndex+6].Text = names[index];
            if (index < 3)          // 기본 데이터
            {
                for (int i = 0; i < 32; i++)
                {
                    txtOut[i + currentIndex * 32].Enabled = true;
                    txtOut[i + currentIndex * 32].Text = existingData[i + index * 32].ToString();
                }
                for (int i = 0; i < 28; i++)
                {
                    txtWOut[i + currentIndex * 28].Enabled = true;
                    txtWOut[i + currentIndex * 28].Text = existingWData[i + index * 28].ToString();
                }
                for (int i = 0; i < 24; i++)
                {
                    txtROut[i + currentIndex * 24].Enabled = true;
                    txtROut[i + currentIndex * 24].Text = existingRData[i + index * 24].ToString();
                }
            }
            else                    // 시뮬레이션 데이터
            {
                if (index == 3 && radio기존업계평균.Checked)  //  index 3: 시뮬레이션 업계평균
                {
                    for (int i = 0; i < 32; i++)
                    {
                        txtOut[i + currentIndex * 32].Enabled = true;
                        txtOut[i + currentIndex * 32].Text = existingData[i + (index - 3) * 32].ToString();
                    }
                    for (int i = 0; i < 28; i++)
                    {
                        txtWOut[i + currentIndex * 28].Enabled = true;
                        txtWOut[i + currentIndex * 28].Text = existingData[i + (index - 3) * 28].ToString();
                    }
                    for (int i = 0; i < 24; i++)
                    {
                        txtROut[i + currentIndex * 24].Enabled = true;
                        txtROut[i + currentIndex * 24].Text = existingData[i + (index - 3) * 24].ToString();
                    }
                }
                else
                {
                    for (int i = 0; i < 32; i++)
                    {
                        txtOut[i + currentIndex * 32].Enabled = true;
                        txtOut[i + currentIndex * 32].Text = simulData[i + (index - 3) * 32].ToString();
                    }
                    for (int i = 0; i < 28; i++)
                    {
                        txtWOut[i + currentIndex * 28].Enabled = true;
                        txtWOut[i + currentIndex * 28].Text = simulWData[i + (index - 3) * 28].ToString();
                    }
                    for (int i = 0; i < 24; i++)
                    {
                        txtROut[i + currentIndex * 24].Enabled = true;
                        txtROut[i + currentIndex * 24].Text = simulRData[i + (index - 3) * 24].ToString();
                    }
                }
            }

            currentIndex++;
        }

        private void setDataToEmpty()
        {
            lblTitle[currentIndex].Text = "";
            lblTitle[currentIndex+3].Text = "";
            lblTitle[currentIndex+6].Text = "";
            for (int i = 0; i < 32; i++)
            {
                txtOut[i + currentIndex * 32].Text = "0";
                txtOut[i + currentIndex * 32].Enabled = false;
            }
            for (int i = 0; i < 28; i++)
            {
                txtWOut[i + currentIndex * 28].Text = "0";
                txtWOut[i + currentIndex * 28].Enabled = false;
            }
            for (int i = 0; i < 24; i++)
            {
                txtROut[i + currentIndex * 24].Text = "0";
                txtROut[i + currentIndex * 24].Enabled = false;
            }
            currentIndex++;
        }

        private void addComma_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || ((sender as TextBox).Text.Length > 0 && (sender as TextBox).Text != "0"))
            {
                (sender as TextBox).Text = String.Format("{0:#,###}", Convert.ToDouble((sender as TextBox).Text.Replace(",", "")));
                (sender as TextBox).SelectionStart = (sender as TextBox).Text.Length;
                if ((sender as TextBox).Text == "")
                    (sender as TextBox).Text = "0";
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            pnlChart.Visible = true;
            for (int i = 0; i < labelWon.Length; i++)
            {
                labelWon[i].Visible = false;
            }
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            pnlChart.Visible = false;
            for(int i =0; i < labelWon.Length; i++){
                labelWon[i].Visible = true;
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            pnlChart2.Visible = true;
            for (int i = 0; i < labelWon.Length; i++)
            {
                labelWon2[i].Visible = false;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            pnlChart2.Visible = false;
            for (int i = 0; i < labelWon.Length; i++)
            {
                labelWon2[i].Visible = true;
            }
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            pnlChart3.Visible = true;
            for (int i = 0; i < labelWon.Length; i++)
            {
                labelWon3[i].Visible = false;
            }
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            pnlChart3.Visible = false;
            for (int i = 0; i < labelWon.Length; i++)
            {
                labelWon3[i].Visible = true;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            applyData((sender as RadioButton).Checked);
        }

        internal bool isFileShowing()
        {
            return radio기존업계평균.Checked;
        }

        internal void saveSimulateFile()
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "LGE File|*.lge|Excel File|*.xlsx";
            saveFileDialog1.Title = "시뮬레이션 파일 저장";
            saveFileDialog1.InitialDirectory = CommonUtil.dataDirectory;
            saveFileDialog1.DefaultExt = "lge";
            saveFileDialog1.AutoUpgradeEnabled = true;
            saveFileDialog1.AddExtension = true;
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = "시뮬레이션_" + CDataControl.g_ReportData.get통신사() + "_" + CDataControl.g_ReportData.get지역() + "_" + CDataControl.g_ReportData.get대리점() + "_" + CDataControl.g_ReportData.get마케터() + "_" + DateTime.Now.ToString("yyyyMMdd");

            // If the directory doesn't exist, create it.
            if (!Directory.Exists(CommonUtil.dataDirectory))
            {
                Directory.CreateDirectory(CommonUtil.dataDirectory);
            }
            if (CommonUtil.saveAsSimulName == null)
            {
                if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (saveFileDialog1.FileName.EndsWith("lge"))
                    {
                        CommonUtil.writeLGEFile(saveFileDialog1.FileName, "|", CommonUtil.파일종류_시뮬레이션);
                    }
                    else if (saveFileDialog1.FileName.EndsWith("xlsx"))
                    {
                        FileInfo fi2 = new FileInfo(CommonUtil.defaultSimulName);
                        fi2.CopyTo(saveFileDialog1.FileName, true);

                        CommonUtil.WriteDataToExcelFile(saveFileDialog1.FileName, CommonUtil.파일종류_시뮬레이션);
                    }
                    else
                    {
                        return;
                    }
                    CommonUtil.saveAsSimulName = saveFileDialog1.FileName;
                }
            }
            else
            {
                if (saveFileDialog1.FileName.EndsWith("lge"))
                {
                    CommonUtil.writeLGEFile(saveFileDialog1.FileName, "|", CommonUtil.파일종류_시뮬레이션);
                }
                else if (saveFileDialog1.FileName.EndsWith("xlsx"))
                {
                    FileInfo fi2 = new FileInfo(CommonUtil.defaultSimulName);
                    fi2.CopyTo(saveFileDialog1.FileName, true);

                    CommonUtil.WriteDataToExcelFile(saveFileDialog1.FileName, CommonUtil.파일종류_시뮬레이션);
                }
            }
        }

        internal void saveAsNewSimulateFile()
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "LGE File|*.lge|Excel File|*.xlsx";
            saveFileDialog1.Title = "시뮬레이션 파일 저장";
            saveFileDialog1.InitialDirectory = CommonUtil.dataDirectory;
            saveFileDialog1.DefaultExt = "lge";
            saveFileDialog1.AutoUpgradeEnabled = true;
            saveFileDialog1.AddExtension = true;
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = "시뮬레이션_" + CDataControl.g_ReportData.get통신사() + "_" + CDataControl.g_ReportData.get지역() + "_" + CDataControl.g_ReportData.get대리점() + "_" + CDataControl.g_ReportData.get마케터() + "_" + DateTime.Now.ToString("yyyyMMdd");

            // If the directory doesn't exist, create it.
            if (!Directory.Exists(CommonUtil.dataDirectory))
            {
                Directory.CreateDirectory(CommonUtil.dataDirectory);
            }
            
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (saveFileDialog1.FileName.EndsWith("lge"))
                {
                    CommonUtil.writeLGEFile(saveFileDialog1.FileName, "|", CommonUtil.파일종류_시뮬레이션);
                }
                else if (saveFileDialog1.FileName.EndsWith("xlsx"))
                {
                    FileInfo fi2 = new FileInfo(CommonUtil.defaultSimulName);
                    fi2.CopyTo(saveFileDialog1.FileName, true);

                    CommonUtil.WriteDataToExcelFile(saveFileDialog1.FileName, CommonUtil.파일종류_시뮬레이션);
                }
                else
                {
                    return;
                }
                CommonUtil.saveAsSimulName = saveFileDialog1.FileName;
            }
        }

        private void set점별손익추정_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text == "")
                return;
            (sender as TextBox).Text = "";
            //if ((sender as TextBox).Text.EndsWith("점 기준"))
            //    return;
            //(sender as TextBox).Text += "점 기준";
        }
    }
}
