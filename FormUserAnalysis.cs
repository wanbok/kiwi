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
    public partial class FormUserAnalysis : Form
    {
        private FormUserOutput mFormUserOutput;

        private TextBox[] txtOut = null;     //전체
        private TextBox[] txtBaseData = null;
        private RichTextBox[] txtComments = null;
        private PictureBox[] picCompare = null;

        public FormUserAnalysis(FormUserOutput FormUserOutput)
        {
            InitializeComponent();

            //더블 버퍼
            this.SetStyle(ControlStyles.DoubleBuffer, true);
            this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            this.SetStyle(ControlStyles.UserPaint, true);

            mFormUserOutput = FormUserOutput;
            
        }

        public FormUserAnalysis()
        {
            InitializeComponent();

            txtOut = new TextBox[64] { txtOut1, txtOut2, txtOut3, txtOut4, txtOut5, txtOut6, txtOut7, txtOut8, txtOut9, txtOut10,
            txtOut11, txtOut12, txtOut13, txtOut14, txtOut15, txtOut16, txtOut17, txtOut17, txtOut19, txtOut20,
            txtOut21, txtOut22, txtOut23, txtOut24, txtOut25, txtOut26, txtOut27, txtOut28, txtOut29, txtOut30,
            txtOut31, txtOut32, txtOut33, txtOut34, txtOut35, txtOut36, txtOut37, txtOut38, txtOut39, txtOut40,
            txtOut41, txtOut42, txtOut43, txtOut44, txtOut45, txtOut46, txtOut47, txtOut48, txtOut49, txtOut50,
            txtOut51, txtOut52, txtOut53, txtOut54, txtOut55, txtOut56, txtOut57, txtOut58, txtOut58, txtOut60,
            txtOut61, txtOut62, txtOut63, txtOut64
            };

            txtBaseData = new TextBox[10] { textBoxBase1, textBoxBase2, textBoxBase3, textBoxBase4, textBoxBase5, 
            textBoxBase6, textBoxBase7, textBoxBase8, textBoxBase9, textBoxBase10};

            txtComments = new RichTextBox[3] { comments1, comments2, comments3 };

            picCompare = new PictureBox[16] { picCompare1, picCompare2, picCompare3, picCompare4, picCompare5, 
                picCompare6, picCompare7, picCompare8, picCompare9, picCompare10,
                picCompare11, picCompare12, picCompare13, picCompare14, picCompare15, picCompare16};


            CResultData BusinessData = CDataControl.g_ResultBusiness as CResultData;
            CResultData StoreData = CDataControl.g_ResultStore as CResultData;
            CResultData FutureData = CDataControl.g_ResultFuture as CResultData;

            setOut();
            setBaseData(CDataControl.g_BasicInput);
            setComments(CDataControl.g_ReportData);
            setCompare();
            setReferrence();


            OpenChart(chart1, BusinessData, StoreData);
            OpenChart(chart2, BusinessData, StoreData);
            OpenChart(chart3, BusinessData, StoreData);
            OpenChart(chart5, BusinessData, StoreData);
            OpenChart(chart6, BusinessData, StoreData);
            OpenChart(chart7, BusinessData, StoreData);
        }

        private void setReferrence()
        {
            if (CDataControl.g_ResultBusinessTotal == null ||
                CDataControl.g_ResultBusiness == null ||
                CDataControl.g_ResultStoreTotal == null ||
                CDataControl.g_ResultStore == null) 
                return;

            long 가입자당ARPU = 0;
            if(CDataControl.g_BasicInput.get누적가입자수_합계()==0)
            {
                가입자당ARPU = 0;
            }
            else
            {
                가입자당ARPU = Convert.ToInt64(CDataControl.g_ResultBusinessTotal.get도매_수익_가입자관리수수료()
                                   / Convert.ToInt64(CDataControl.g_BasicInput.getstr누적가입자수_합계()));
            }

            long 월평균인건비 = 0;
            if (Convert.ToInt64(CDataControl.g_ResultBusinessTotal.getstr전체_비용_인건비_급여_복리후생비())==0)
            {
                월평균인건비 = 0;                                
            }
            else
            {
                월평균인건비 = CDataControl.g_ResultBusiness.get전체_비용_인건비_급여_복리후생비() 
                               /Convert.ToInt64(CDataControl.g_ResultBusinessTotal.getstr전체_비용_인건비_급여_복리후생비());
            }

            long 판촉비비중 = 0;
            if(Convert.ToInt64(CDataControl.g_ResultBusinessTotal.전체_비용_소계)==0)
            {
                판촉비비중 = 0;
            }
            else
            {            
                판촉비비중 = Convert.ToInt64(CDataControl.g_ResultBusinessTotal.get전체_비용_기타판매관리비()) 
                             / Convert.ToInt64(CDataControl.g_ResultBusinessTotal.전체_비용_소계);
            }

            long 인당판매수량 = 0;
            if(Convert.ToInt64(CDataControl.g_BasicInput.getstr도매_직원수_소계())==0)
            {
                인당판매수량 = 0;
            }
            else
            {
                인당판매수량 = (Convert.ToInt64(CDataControl.g_BasicInput.getstr도매_월평균판매대수_신규())
                                    + Convert.ToInt64(CDataControl.g_BasicInput.getstr도매_월평균판매대수_기변()))
                                    / Convert.ToInt64(CDataControl.g_BasicInput.getstr도매_직원수_소계());
            }

            textBox69.Text = 가입자당ARPU.ToString();
            textBox71.Text = 월평균인건비.ToString();
            textBox72.Text = 판촉비비중.ToString();
            textBox74.Text = 인당판매수량.ToString();
        }

        private void setOut()
        {
            if (CDataControl.g_ResultBusinessTotal == null ||
                CDataControl.g_ResultBusiness == null ||
                CDataControl.g_ResultStoreTotal == null ||
                CDataControl.g_ResultStore == null) return;
            for (int i = 0; i < 16; i++)
            {
                txtOut[i].Text = CDataControl.g_ResultBusinessTotal.getArr전체_수익_비용_및_계산포함()[i].ToString();
                txtOut[i + 16].Text = CDataControl.g_ResultBusiness.getArr전체_수익_비용_및_계산포함()[i].ToString();
                txtOut[i + 32].Text = CDataControl.g_ResultStoreTotal.getArr전체_수익_비용_및_계산포함()[i].ToString();
                txtOut[i + 48].Text = CDataControl.g_ResultStore.getArr전체_수익_비용_및_계산포함()[i].ToString();
            }
        }


        private void setBaseData(CBasicInput _basicInput)
        {
            if (_basicInput == null) return;
            int i = 0;

            txtBaseData[i++].Text = _basicInput.getstr도매_누적가입자수();
            txtBaseData[i++].Text = _basicInput.getstr도매_월평균판매대수_신규();
            txtBaseData[i++].Text = _basicInput.getstr소매_월평균판매대수_신규();
            txtBaseData[i++].Text = _basicInput.getstr도매_월평균판매대수_기변();
            txtBaseData[i++].Text = _basicInput.getstr소매_월평균판매대수_기변();
            txtBaseData[i++].Text = _basicInput.getstr도매_월평균유통모델출고대수_소계();
            txtBaseData[i++].Text = _basicInput.getstr도매_거래선수_소계();
            txtBaseData[i++].Text = _basicInput.getstr소매_거래선수_소계();
            txtBaseData[i++].Text = _basicInput.getstr도매_직원수_소계();
            txtBaseData[i++].Text = _basicInput.getstr소매_직원수_소계();
        }

        private void setComments(CReportData _resultData)
        {
            if (_resultData == null) return;
            int i = 0;
            txtComments[i++].Text = _resultData.get분석내용_및_대리점_활동방향().ToString();
            txtComments[i++].Text = _resultData.getLG_지원_활동().ToString();
            txtComments[i++].Text = _resultData.get배경_및_이슈().ToString();
        }

        public void saveComments()
        {
            if (CDataControl.g_ReportData == null) return;
            CDataControl.g_ReportData.set분석내용_및_대리점_활동방향(txtComments[0].Text);
            CDataControl.g_ReportData.setLG_지원_활동(txtComments[1].Text);
            CDataControl.g_ReportData.set배경_및_이슈(txtComments[2].Text);
            //if (CommonUtil.openAsName != null)
            //{
            //    excel.Worksheet _WorkSheet = CommonUtil.GetExcelWorksheet(CommonUtil.openAsName, 3);
            //    _WorkSheet.get_Range("B4").Value2 = txtComments[0].Text;
            //    _WorkSheet.get_Range("C4").Value2 = txtComments[1].Text;
            //    _WorkSheet.SaveAs(CommonUtil.openAsName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //}
            //else {
            //    MessageBox.Show("파일을 열어야 합니다.");
            //}
        }

        private void setCompare() {
            // 17, 49
            Int64 convertedA;
            Int64 convertedB;

            for (int i = 0; i < 16; i++) {
                convertedA = Convert.ToInt64(txtOut[i + 16].Text);
                convertedB = Convert.ToInt64(txtOut[i + 48].Text);
                if (convertedA < convertedB) { picCompare[i].Image = KIWI.Properties.Resources.up5; }
                else if (convertedA > convertedB) { picCompare[i].Image = KIWI.Properties.Resources.down1; }
                else { picCompare[i].Image = KIWI.Properties.Resources.equal; }
            }

        }


        // chart1 - 수익계정 전체    CResultData
        //private void OpenChart(Chart chart, excel.Worksheet sheet)
        private void OpenChart(Chart chart, CResultData _bizResult, CResultData _storeResult)
        {
            if (_bizResult == null || _storeResult == null) return;
            double[] yValues = null;
            double[] yValues2 = null;

            string[] xValues = null;

            if (chart.Name == "chart1")
            {
                xValues = new string[6] { "X1", "X2", "X3", "X4", "X5", "X6" };

                yValues = new double[6]{ 
                                         Convert.ToDouble(_bizResult.getstr도매_수익_가입자관리수수료()), 
                                         Convert.ToDouble(_bizResult.getstr도매_수익_CS관리수수료()), 
                                         Convert.ToDouble(_bizResult.getstr소매_수익_업무취급수수료()), 
                                         Convert.ToDouble(_bizResult.getstr도매_수익_사업자모델매입에따른추가수익()), 
                                         Convert.ToDouble(_bizResult.getstr도매_수익_유통모델매입에따른추가수익_현금_Volume()), 
                                         Convert.ToDouble(_bizResult.getstr소매_수익_직영매장판매수익())
                                        };

                yValues2 = new double[6]{ Convert.ToDouble(txtOut49.Text), Convert.ToDouble(txtOut50.Text), Convert.ToDouble(txtOut51.Text), 
                            Convert.ToDouble(txtOut52.Text), Convert.ToDouble(txtOut53.Text), Convert.ToDouble(txtOut54.Text) };
            }
            else if (chart.Name == "chart2")
            {
                xValues = new string[4] { "X1", "X2", "X3", "X4" };

                yValues = new double[4]{ 
                                         Convert.ToDouble(_bizResult.getstr도매_수익_가입자관리수수료()),
                                         Convert.ToDouble(_bizResult.getstr도매_수익_CS관리수수료()),
                                         Convert.ToDouble(_bizResult.getstr도매_수익_사업자모델매입에따른추가수익()),
                                         Convert.ToDouble(_bizResult.getstr도매_수익_유통모델매입에따른추가수익_현금_Volume())
                                        };

                yValues2 = new double[4]{ 
                                         Convert.ToDouble(_storeResult.getstr도매_수익_가입자관리수수료()),
                                         Convert.ToDouble(_storeResult.getstr도매_수익_CS관리수수료()),
                                         Convert.ToDouble(_storeResult.getstr도매_수익_사업자모델매입에따른추가수익()),
                                         Convert.ToDouble(_storeResult.getstr도매_수익_유통모델매입에따른추가수익_현금_Volume())                                   
                                        };
            }
            else if (chart.Name == "chart3")
            {
                xValues = new string[2] { "X1", "X2" };

                yValues = new double[2]{ 
                                         Convert.ToDouble(_bizResult.getstr소매_수익_업무취급수수료()),
                                         Convert.ToDouble(_bizResult.getstr소매_수익_직영매장판매수익()),                                                   
                                        };

                yValues2 = new double[2]{ 
                                         Convert.ToDouble(_storeResult.getstr소매_수익_업무취급수수료()),
                                         Convert.ToDouble(_storeResult.getstr소매_수익_직영매장판매수익())         
                                        };
            }
            else if (chart.Name == "chart4")
            {
                xValues = new string[6] { "X1", "X2", "X3", "X4", "X5", "X6" };

                yValues = new double[6]{ 
                                         Convert.ToDouble(_bizResult.getstr도매_수익_가입자관리수수료()), 
                                         Convert.ToDouble(_bizResult.getstr도매_수익_CS관리수수료()), 
                                         Convert.ToDouble(_bizResult.getstr소매_수익_업무취급수수료()), 
                                         Convert.ToDouble(_bizResult.getstr도매_수익_사업자모델매입에따른추가수익()), 
                                         Convert.ToDouble(_bizResult.getstr도매_수익_유통모델매입에따른추가수익_현금_Volume()),
                                         Convert.ToDouble(_bizResult.getstr소매_수익_직영매장판매수익()) 
                                        };

                yValues2 = new double[6]{ 
                                         Convert.ToDouble(_storeResult.getstr도매_수익_가입자관리수수료()), 
                                         Convert.ToDouble(_storeResult.getstr도매_수익_CS관리수수료()), 
                                         Convert.ToDouble(_storeResult.getstr소매_수익_업무취급수수료()), 
                                         Convert.ToDouble(_storeResult.getstr도매_수익_사업자모델매입에따른추가수익()), 
                                         Convert.ToDouble(_storeResult.getstr도매_수익_유통모델매입에따른추가수익_현금_Volume()), 
                                         Convert.ToDouble(_storeResult.getstr소매_수익_직영매장판매수익())
                                        };
            }
            else if (chart.Name == "chart5")
            {
                xValues = new string[4] { "X1", "X2", "X3", "X4" };

                yValues = new double[4]{ 
                                         Convert.ToDouble(_bizResult.getstr도매_수익_가입자관리수수료()),
                                         Convert.ToDouble(_bizResult.getstr도매_수익_CS관리수수료()),
                                         Convert.ToDouble(_bizResult.getstr도매_수익_사업자모델매입에따른추가수익()),
                                         Convert.ToDouble(_bizResult.getstr도매_수익_유통모델매입에따른추가수익_현금_Volume())
                                        };

                yValues2 = new double[4]{ 
                                         Convert.ToDouble(_storeResult.getstr도매_수익_가입자관리수수료()),
                                         Convert.ToDouble(_storeResult.getstr도매_수익_CS관리수수료()),
                                         Convert.ToDouble(_storeResult.getstr도매_수익_사업자모델매입에따른추가수익()),
                                         Convert.ToDouble(_storeResult.getstr도매_수익_유통모델매입에따른추가수익_현금_Volume())                                               
                                        };
            }
            else if (chart.Name == "chart6")
            {
                xValues = new string[2] { "X1", "X2" };

                yValues = new double[2]{ 
                                         Convert.ToDouble(_bizResult.getstr소매_수익_업무취급수수료()),
                                         Convert.ToDouble(_bizResult.getstr소매_수익_직영매장판매수익())
                                        };

                yValues2 = new double[2]{ 
                                         Convert.ToDouble(_storeResult.getstr소매_수익_업무취급수수료()),
                                         Convert.ToDouble(_storeResult.getstr소매_수익_직영매장판매수익())
                                        };
            }
            
            chart.Series[0].Name = "업계평균";
            chart.Series[1].Name = "당대리점";

            chart.Series[0].Points.DataBindXY(xValues, yValues);
            chart.Series[1].Points.DataBindXY(xValues, yValues2);
        }

        private void chart1_Click(object sender, EventArgs e)
        {
            FormChartViewer viewer = new FormChartViewer();
            viewer.MakeChart(sender as Chart);
            viewer.ShowDialog();
        }
    }
}
