using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.IO;

namespace KIWI
{
    public partial class FormUserSimulateInput : Form
    {
        private FormUserSimulateOutput mFormUserSimulOutput;

        private TextBox[] txtInput = null;     //기본입력
        private TextBox[] txtDetail = null;    //상세입력

        private string txtMangeInput1 = "";
        private string txtMangeInput2 = "";
        private string txtMangeInput3 = "";
        private string txtMangeInput4 = "";
        private string txtMangeInput5 = "";
        private string txtMangeInput6 = "";
        private string txtMangeInput7 = "";
        private string txtMangeInput8 = "";
        private string txtMangeInput9 = "";
        private string txtMangeInput10 = "";
        private string txtMangeInput11 = "";
        private string txtMangeInput12 = "";
        private string txtMangeInput13 = "";
        private string txtMangeInput14 = "";
        private string txtMangeInput15 = "";
        private string txtMangeInput16 = "";
        private string txtMangeInput17 = "";
        private string txtMangeInput18 = "";
        private string txtMangeInput19 = "";
        private string txtMangeInput20 = "";
        private string txtMangeInput21 = "";
        private string txtMangeInput22 = "";
        private string txtMangeInput23 = "";
        private string txtMangeInput24 = "";
        private string txtMangeInput25 = "";
        private string txtMangeInput26 = "";
        private string txtMangeInput27 = "";
        private string txtMangeInput28 = "";
        private string txtMangeInput29 = "";
        private string txtMangeInput30 = "";
        private string txtMangeInput31 = "";


        private string[] txtMangeInput = null;

        private string area = "";
        private string beanch = "";
        private string name = "";

        //public FormUserSimulateInput(FormUserOutput formUserOutput)
        //{
        //    InitializeComponent();

        //    //더블 버퍼
        //    this.SetStyle(ControlStyles.DoubleBuffer, true);
        //    this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
        //    this.SetStyle(ControlStyles.UserPaint, true);

        //    mFormUserOutput = formUserOutput;
        //}


        public FormUserSimulateInput(FormUserSimulateOutput frmOutput)
        {
            InitializeComponent();
            mFormUserSimulOutput = frmOutput;
            txtMangeInput = new string[31] { txtMangeInput1, txtMangeInput2, txtMangeInput3, txtMangeInput4, txtMangeInput5, 
                txtMangeInput6, txtMangeInput7, txtMangeInput8, txtMangeInput9, txtMangeInput10,
                txtMangeInput11, txtMangeInput12, txtMangeInput13, txtMangeInput14, txtMangeInput15,
                txtMangeInput16, txtMangeInput17, txtMangeInput18, txtMangeInput19, txtMangeInput20,
                txtMangeInput21, txtMangeInput22, txtMangeInput23, txtMangeInput24, txtMangeInput25, 
                txtMangeInput26, txtMangeInput27, txtMangeInput28, txtMangeInput29, txtMangeInput30,
                txtMangeInput31

            };

            txtInput = new TextBox[28] { txtInput1, txtInput2, txtInput3, txtInput4, txtInput5, txtInput6, txtInput7, txtInput8, txtInput9, txtInput10,
            txtInput11, txtInput12, txtInput13, txtInput14, txtInput15, txtInput16, txtInput17, txtInput18, txtInput19, txtInput20,
            txtInput21, txtInput22, txtInput23, txtInput24, txtInput25, txtInput26, txtInput27, txtInput28
            };

            txtDetail = new TextBox[72] { txtDetail1, txtDetail2, txtDetail3, txtDetail4, txtDetail5, txtDetail6, txtDetail7, txtDetail8, txtDetail9, txtDetail10,
            txtDetail11, txtDetail12, txtDetail13, txtDetail14, txtDetail15, txtDetail16, txtDetail17, txtDetail18, txtDetail19, txtDetail20,            
            txtDetail21, txtDetail22, txtDetail23, txtDetail24, txtDetail25, txtDetail26, txtDetail27, txtDetail28, txtDetail29, txtDetail30,            
            txtDetail31, txtDetail32, txtDetail33, txtDetail34, txtDetail35, txtDetail36, txtDetail37, txtDetail38, txtDetail39, txtDetail40,            
            txtDetail41, txtDetail42, txtDetail43, txtDetail44, txtDetail45, txtDetail46, txtDetail47, txtDetail48, txtDetail49, txtDetail50,            
            txtDetail51, txtDetail52, txtDetail53, txtDetail54, txtDetail55, txtDetail56, txtDetail57, txtDetail58, txtDetail59, txtDetail60,            
            txtDetail61, txtDetail62, txtDetail63, txtDetail64, txtDetail65, txtDetail66, txtDetail67, txtDetail68, txtDetail69, txtDetail70,            
            txtDetail71, txtDetail72
            };

            this.txtDetail37.TextChanged += new System.EventHandler(this.txtDetail37_TextChanged);
            this.txtDetail38.TextChanged += new System.EventHandler(this.txtDetail38_TextChanged);
            this.txtDetail39.TextChanged += new System.EventHandler(this.txtDetail39_TextChanged);
            this.txtDetail40.TextChanged += new System.EventHandler(this.txtDetail40_TextChanged);
            this.txtDetail41.TextChanged += new System.EventHandler(this.txtDetail41_TextChanged);

            this.txtDetail43.TextChanged += new System.EventHandler(this.txtDetail43_TextChanged);
            this.txtDetail44.TextChanged += new System.EventHandler(this.txtDetail44_TextChanged);
            this.txtDetail45.TextChanged += new System.EventHandler(this.txtDetail45_TextChanged);
            this.txtDetail46.TextChanged += new System.EventHandler(this.txtDetail46_TextChanged);
            this.txtDetail47.TextChanged += new System.EventHandler(this.txtDetail47_TextChanged);
            this.txtDetail48.TextChanged += new System.EventHandler(this.txtDetail48_TextChanged);
            this.txtDetail49.TextChanged += new System.EventHandler(this.txtDetail49_TextChanged);
            this.txtDetail50.TextChanged += new System.EventHandler(this.txtDetail50_TextChanged);
            this.txtDetail51.TextChanged += new System.EventHandler(this.txtDetail51_TextChanged);
            this.txtDetail52.TextChanged += new System.EventHandler(this.txtDetail52_TextChanged);

            this.txtDetail54.TextChanged += new System.EventHandler(this.txtDetail54_TextChanged);
            this.txtDetail55.TextChanged += new System.EventHandler(this.txtDetail55_TextChanged);

            this.txtDetail57.TextChanged += new System.EventHandler(this.txtDetail57_TextChanged);
            this.txtDetail58.TextChanged += new System.EventHandler(this.txtDetail58_TextChanged);
            this.txtDetail59.TextChanged += new System.EventHandler(this.txtDetail59_TextChanged);
            this.txtDetail60.TextChanged += new System.EventHandler(this.txtDetail60_TextChanged);
            this.txtDetail61.TextChanged += new System.EventHandler(this.txtDetail61_TextChanged);
            this.txtDetail62.TextChanged += new System.EventHandler(this.txtDetail62_TextChanged);

            this.txtDetail64.TextChanged += new System.EventHandler(this.txtDetail64_TextChanged);
            this.txtDetail65.TextChanged += new System.EventHandler(this.txtDetail65_TextChanged);
            this.txtDetail66.TextChanged += new System.EventHandler(this.txtDetail66_TextChanged);
            this.txtDetail67.TextChanged += new System.EventHandler(this.txtDetail67_TextChanged);
            this.txtDetail68.TextChanged += new System.EventHandler(this.txtDetail68_TextChanged);
            this.txtDetail69.TextChanged += new System.EventHandler(this.txtDetail69_TextChanged);
            this.txtDetail70.TextChanged += new System.EventHandler(this.txtDetail70_TextChanged);
            this.txtDetail71.TextChanged += new System.EventHandler(this.txtDetail71_TextChanged);

            CommonUtil.clearTextBox(this.tabPage1);
            CommonUtil.clearTextBox(this.tabPage5);

            getInput();
            getDetail();
            
            
        }
        //상세입력
        private void getDetail()
        {

            // 셀에서 데이터 가져오기

            txtDetail[0].Text = CDataControl.g_DetailInput.getstr도매_수익_월평균관리수수료();
            txtDetail[1].Text = CDataControl.g_DetailInput.getstr도매_수익_CS관리수수료();
            txtDetail[2].Text = CDataControl.g_DetailInput.getstr도매_수익_사업자모델매입관련추가수익();
            txtDetail[3].Text = CDataControl.g_DetailInput.getstr도매_수익_유통모델매입관련추가수익_현금DC();
            txtDetail[4].Text = CDataControl.g_DetailInput.getstr도매_수익_유통모델매입관련추가수익_VolumeDC();
            txtDetail[5].Text = (CommonUtil.StringToDoubleVal(txtDetail[0].Text) + CommonUtil.StringToDoubleVal(txtDetail[1].Text) 
                + CommonUtil.StringToDoubleVal(txtDetail[2].Text) + CommonUtil.StringToDoubleVal(txtDetail[3].Text) 
                + CommonUtil.StringToDoubleVal(txtDetail[4].Text) ).ToString();
            txtDetail[6].Text = CDataControl.g_DetailInput.getstr도매_비용_대리점투자금액_신규();
            txtDetail[7].Text = CDataControl.g_DetailInput.getstr도매_비용_대리점투자금액_기변();
            txtDetail[8].Text = CDataControl.g_DetailInput.getstr도매_비용_직원급여_간부급();
            txtDetail[9].Text = CDataControl.g_DetailInput.getstr도매_비용_직원급여_평사원();
            txtDetail[10].Text = CDataControl.g_DetailInput.getstr도매_비용_지급임차료();
            txtDetail[11].Text = CDataControl.g_DetailInput.getstr도매_비용_운반비();
            txtDetail[12].Text = CDataControl.g_DetailInput.getstr도매_비용_차량유지비();
            txtDetail[13].Text = CDataControl.g_DetailInput.getstr도매_비용_지급수수료();
            txtDetail[14].Text = CDataControl.g_DetailInput.getstr도매_비용_판매촉진비();
            txtDetail[15].Text = CDataControl.g_DetailInput.getstr도매_비용_건물관리비();
            txtDetail[16].Text = (CommonUtil.StringToDoubleVal(txtDetail[6].Text) + CommonUtil.StringToDoubleVal(txtDetail[7].Text)
               + CommonUtil.StringToDoubleVal(txtDetail[8].Text) + CommonUtil.StringToDoubleVal(txtDetail[9].Text)
               + CommonUtil.StringToDoubleVal(txtDetail[10].Text) + CommonUtil.StringToDoubleVal(txtDetail[11].Text)
               + CommonUtil.StringToDoubleVal(txtDetail[12].Text) + CommonUtil.StringToDoubleVal(txtDetail[13].Text)
               + CommonUtil.StringToDoubleVal(txtDetail[14].Text) + CommonUtil.StringToDoubleVal(txtDetail[15].Text)
               ).ToString();
            txtDetail[17].Text = CDataControl.g_DetailInput.getstr소매_수익_월평균업무취급수수료();
            txtDetail[18].Text = CDataControl.g_DetailInput.getstr소매_수익_직영매장판매수익();
            txtDetail[19].Text = (CommonUtil.StringToDoubleVal(txtDetail[17].Text) + CommonUtil.StringToDoubleVal(txtDetail[18].Text)).ToString();

            txtDetail[20].Text = CDataControl.g_DetailInput.getstr소매_비용_직원급여_간부급();
            txtDetail[21].Text = CDataControl.g_DetailInput.getstr소매_비용_직원급여_평사원();
            txtDetail[22].Text = CDataControl.g_DetailInput.getstr소매_비용_지급임차료();
            txtDetail[23].Text = CDataControl.g_DetailInput.getstr소매_비용_지급수수료();
            txtDetail[24].Text = CDataControl.g_DetailInput.getstr소매_비용_판매촉진비();
            txtDetail[25].Text = CDataControl.g_DetailInput.getstr소매_비용_건물관리비();
            txtDetail[26].Text = (CommonUtil.StringToDoubleVal(txtDetail[20].Text) + CommonUtil.StringToDoubleVal(txtDetail[21].Text)
                + CommonUtil.StringToDoubleVal(txtDetail[22].Text) + CommonUtil.StringToDoubleVal(txtDetail[23].Text)
                + CommonUtil.StringToDoubleVal(txtDetail[24].Text) + CommonUtil.StringToDoubleVal(txtDetail[25].Text)).ToString();

            txtDetail[27].Text = CDataControl.g_DetailInput.getstr도소매_비용_복리후생비();
            txtDetail[28].Text = CDataControl.g_DetailInput.getstr도소매_비용_통신비();
            txtDetail[29].Text = CDataControl.g_DetailInput.getstr도소매_비용_공과금();
            txtDetail[30].Text = CDataControl.g_DetailInput.getstr도소매_비용_소모품비();
            txtDetail[31].Text = CDataControl.g_DetailInput.getstr도소매_비용_이자비용();
            txtDetail[32].Text = CDataControl.g_DetailInput.getstr도소매_비용_부가세();
            txtDetail[33].Text = CDataControl.g_DetailInput.getstr도소매_비용_법인세();
            txtDetail[34].Text = CDataControl.g_DetailInput.getstr도소매_비용_기타();
            txtDetail[35].Text = (CommonUtil.StringToDoubleVal(txtDetail[27].Text) + CommonUtil.StringToDoubleVal(txtDetail[28].Text)
                + CommonUtil.StringToDoubleVal(txtDetail[29].Text) + CommonUtil.StringToDoubleVal(txtDetail[30].Text)
                + CommonUtil.StringToDoubleVal(txtDetail[31].Text) + CommonUtil.StringToDoubleVal(txtDetail[32].Text)
                + CommonUtil.StringToDoubleVal(txtDetail[33].Text) + CommonUtil.StringToDoubleVal(txtDetail[34].Text)
                ).ToString();


            txtDetail[36].Text  = CDataControl.g_DetailInput.getstr도매_수익_월평균관리수수료();
            txtDetail[37].Text =  CDataControl.g_DetailInput.getstr도매_수익_CS관리수수료();
            txtDetail[38].Text =  CDataControl.g_DetailInput.getstr도매_수익_사업자모델매입관련추가수익();
            txtDetail[39].Text =  CDataControl.g_DetailInput.getstr도매_수익_유통모델매입관련추가수익_현금DC();
            txtDetail[40].Text =  CDataControl.g_DetailInput.getstr도매_수익_유통모델매입관련추가수익_VolumeDC();
            txtDetail[41].Text =  (CommonUtil.StringToDoubleVal(txtDetail[0].Text) + CommonUtil.StringToDoubleVal(txtDetail[1].Text) 
                + CommonUtil.StringToDoubleVal(txtDetail[2].Text) + CommonUtil.StringToDoubleVal(txtDetail[3].Text)  
                + CommonUtil.StringToDoubleVal(txtDetail[4].Text) ).ToString(); 
            txtDetail[42].Text  = CDataControl.g_DetailInput.getstr도매_비용_대리점투자금액_신규();
            txtDetail[43].Text =  CDataControl.g_DetailInput.getstr도매_비용_대리점투자금액_기변();
            txtDetail[44].Text =  CDataControl.g_DetailInput.getstr도매_비용_직원급여_간부급();
            txtDetail[45].Text =  CDataControl.g_DetailInput.getstr도매_비용_직원급여_평사원();
            txtDetail[46].Text =   CDataControl.g_DetailInput.getstr도매_비용_지급임차료();
            txtDetail[47].Text =   CDataControl.g_DetailInput.getstr도매_비용_운반비();
            txtDetail[48].Text =   CDataControl.g_DetailInput.getstr도매_비용_차량유지비();
            txtDetail[49].Text =   CDataControl.g_DetailInput.getstr도매_비용_지급수수료();
            txtDetail[50].Text =   CDataControl.g_DetailInput.getstr도매_비용_판매촉진비();
            txtDetail[51].Text =   CDataControl.g_DetailInput.getstr도매_비용_건물관리비();
            txtDetail[52].Text =  (CommonUtil.StringToDoubleVal(txtDetail[6].Text) + CommonUtil.StringToDoubleVal(txtDetail[7].Text)
               + CommonUtil.StringToDoubleVal(txtDetail[8].Text) + CommonUtil.StringToDoubleVal(txtDetail[9].Text)  
               + CommonUtil.StringToDoubleVal(txtDetail[10].Text) + CommonUtil.StringToDoubleVal(txtDetail[11].Text)  
               + CommonUtil.StringToDoubleVal(txtDetail[12].Text) + CommonUtil.StringToDoubleVal(txtDetail[13].Text) 
               + CommonUtil.StringToDoubleVal(txtDetail[14].Text) + CommonUtil.StringToDoubleVal(txtDetail[15].Text) 
               ).ToString();
            txtDetail[53].Text  = CDataControl.g_DetailInput.getstr소매_수익_월평균업무취급수수료();
            txtDetail[54].Text =  CDataControl.g_DetailInput.getstr소매_수익_직영매장판매수익();
            txtDetail[55].Text =  (CommonUtil.StringToDoubleVal(txtDetail[17].Text) + CommonUtil.StringToDoubleVal(txtDetail[18].Text)).ToString();
            txtDetail[56].Text =  CDataControl.g_DetailInput.getstr소매_비용_직원급여_간부급();
            txtDetail[57].Text =  CDataControl.g_DetailInput.getstr소매_비용_직원급여_평사원();
            txtDetail[58].Text =  CDataControl.g_DetailInput.getstr소매_비용_지급임차료();
            txtDetail[59].Text =  CDataControl.g_DetailInput.getstr소매_비용_지급수수료();
            txtDetail[60].Text =  CDataControl.g_DetailInput.getstr소매_비용_판매촉진비();
            txtDetail[61].Text =  CDataControl.g_DetailInput.getstr소매_비용_건물관리비();
            txtDetail[62].Text =  (CommonUtil.StringToDoubleVal(txtDetail[20].Text) + CommonUtil.StringToDoubleVal(txtDetail[21].Text)
                + CommonUtil.StringToDoubleVal(txtDetail[22].Text) + CommonUtil.StringToDoubleVal(txtDetail[23].Text) 
                + CommonUtil.StringToDoubleVal(txtDetail[24].Text) + CommonUtil.StringToDoubleVal(txtDetail[25].Text)).ToString();

            txtDetail[63].Text = CDataControl.g_DetailInput.getstr도소매_비용_복리후생비();
            txtDetail[64].Text = CDataControl.g_DetailInput.getstr도소매_비용_통신비();
            txtDetail[65].Text = CDataControl.g_DetailInput.getstr도소매_비용_공과금();
            txtDetail[66].Text = CDataControl.g_DetailInput.getstr도소매_비용_소모품비();
            txtDetail[67].Text = CDataControl.g_DetailInput.getstr도소매_비용_이자비용();
            txtDetail[68].Text = CDataControl.g_DetailInput.getstr도소매_비용_부가세();
            txtDetail[69].Text = CDataControl.g_DetailInput.getstr도소매_비용_법인세();
            txtDetail[70].Text = CDataControl.g_DetailInput.getstr도소매_비용_기타();
            txtDetail[71].Text =  (CommonUtil.StringToDoubleVal(txtDetail[27].Text) + CommonUtil.StringToDoubleVal(txtDetail[28].Text)
                + CommonUtil.StringToDoubleVal(txtDetail[29].Text) + CommonUtil.StringToDoubleVal(txtDetail[30].Text) 
                + CommonUtil.StringToDoubleVal(txtDetail[31].Text) + CommonUtil.StringToDoubleVal(txtDetail[32].Text) 
                + CommonUtil.StringToDoubleVal(txtDetail[33].Text) + CommonUtil.StringToDoubleVal(txtDetail[34].Text) 
                ).ToString(); 

        }

        //기본입력
        private void getInput()
        {
            area =   CDataControl.g_BasicInput.get지역();
            beanch = CDataControl.g_BasicInput.get대리점();
            name = CDataControl.g_BasicInput.get마케터();

            txtInput[0].Text = CDataControl.g_BasicInput.getstr도매_누적가입자수();
            txtInput[1].Text = CDataControl.g_BasicInput.getstr도매_월평균판매대수_신규();
            txtInput[2].Text = CDataControl.g_BasicInput.getstr도매_월평균판매대수_기변();
            txtInput[3].Text = CDataControl.g_BasicInput.getstr도매_월평균유통모델출고대수_LG();
            txtInput[4].Text = CDataControl.g_BasicInput.getstr도매_월평균유통모델출고대수_SS();
            txtInput[5].Text = CDataControl.g_BasicInput.getstr도매_거래선수_개통사무실();
            txtInput[6].Text = CDataControl.g_BasicInput.getstr도매_거래선수_판매점();
            txtInput[7].Text = CDataControl.g_BasicInput.getstr도매_직원수_간부급();
            txtInput[8].Text = CDataControl.g_BasicInput.getstr도매_직원수_평사원();
            txtInput[9].Text = CDataControl.g_BasicInput.getstr도매_누적가입자수();
            txtInput[10].Text = CDataControl.g_BasicInput.getstr도매_월평균판매대수_신규();
            txtInput[11].Text = CDataControl.g_BasicInput.getstr도매_월평균판매대수_기변();
            txtInput[12].Text = CDataControl.g_BasicInput.getstr도매_월평균유통모델출고대수_LG();
            txtInput[13].Text = CDataControl.g_BasicInput.getstr도매_월평균유통모델출고대수_SS();
            txtInput[14].Text = CDataControl.g_BasicInput.getstr도매_거래선수_개통사무실();
            txtInput[15].Text = CDataControl.g_BasicInput.getstr도매_거래선수_판매점();
            txtInput[16].Text = CDataControl.g_BasicInput.getstr도매_직원수_간부급();
            txtInput[17].Text = CDataControl.g_BasicInput.getstr도매_직원수_평사원();


            txtInput[18].Text = CDataControl.g_BasicInput.getstr소매_월평균판매대수_신규();
            txtInput[19].Text = CDataControl.g_BasicInput.getstr소매_월평균판매대수_기변();
            txtInput[20].Text = CDataControl.g_BasicInput.getstr소매_거래선수_직영점();
            txtInput[21].Text = CDataControl.g_BasicInput.getstr소매_직원수_간부급();
            txtInput[22].Text = CDataControl.g_BasicInput.getstr소매_직원수_평사원();

            txtInput[23].Text = CDataControl.g_BasicInput.getstr소매_월평균판매대수_신규();
            txtInput[24].Text = CDataControl.g_BasicInput.getstr소매_월평균판매대수_기변();
            txtInput[25].Text = CDataControl.g_BasicInput.getstr소매_거래선수_직영점();
            txtInput[26].Text = CDataControl.g_BasicInput.getstr소매_직원수_간부급();
            txtInput[27].Text = CDataControl.g_BasicInput.getstr소매_직원수_평사원();
        }

        private void SaveAsInput()
        {
            CBasicInput bi = CDataControl.g_SimBasicInput;
            CBusinessData di = CDataControl.g_SimDetailInput;
            CResultData[] rdts = new CResultData[] { CDataControl.g_SimResultStoreTotal, CDataControl.g_SimResultFutureTotal };
            CResultData[] rds = new CResultData[] { CDataControl.g_SimResultStore, CDataControl.g_SimResultFuture };
            CResultData rdt = null;
            CResultData rd = null;

            bi.set지역(area);
            bi.set대리점(beanch);
            bi.set마케터(name);

            String[] txtWrite = new String[14] { txtInput10.Text, txtInput11.Text, txtInput12.Text, txtInput13.Text, txtInput14.Text,  
                txtInput15.Text, txtInput16.Text, txtInput17.Text, txtInput18.Text, txtInput24.Text, txtInput25.Text, txtInput26.Text, txtInput27.Text, txtInput28.Text};
            
            String[] txtWrite2 = new String[31]  { 
                txtDetail37.Text, txtDetail38.Text, txtDetail39.Text, txtDetail40.Text, txtDetail41.Text,
                txtDetail43.Text, txtDetail44.Text, txtDetail45.Text, txtDetail46.Text, txtDetail47.Text, txtDetail48.Text, txtDetail49.Text, txtDetail50.Text, txtDetail51.Text, txtDetail52.Text,
                txtDetail54.Text, txtDetail55.Text,
                txtDetail57.Text, txtDetail58.Text, txtDetail59.Text, txtDetail60.Text, txtDetail61.Text, txtDetail62.Text,
                txtDetail64.Text, txtDetail65.Text, txtDetail66.Text, txtDetail67.Text, txtDetail68.Text, txtDetail69.Text, txtDetail70.Text, txtDetail71.Text
            };
            CommonUtil.setInputData(txtWrite, txtWrite2, bi, di, rdts, rds, rdt, rd, CDataControl.g_SimResultBusinessTotal, CDataControl.g_SimResultBusiness);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveAsInput();
            mFormUserSimulOutput.applyData(mFormUserSimulOutput.isFileShowing());
            this.Close();
        }

        private void txtDetail37_TextChanged(object sender, EventArgs e) {
            setTxtInput_TextChanged(sender, e);

            txtDetail42.Text = (CommonUtil.StringToDoubleVal(txtDetail37.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail38.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail39.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail40.Text.Replace(",", ""))
                 + CommonUtil.StringToDoubleVal(txtDetail41.Text.Replace(",", ""))).ToString();
        }
        private void txtDetail38_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);
            txtDetail42.Text = (CommonUtil.StringToDoubleVal(txtDetail37.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail38.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail39.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail40.Text.Replace(",", ""))
                 + CommonUtil.StringToDoubleVal(txtDetail41.Text.Replace(",", ""))).ToString();
        }
        private void txtDetail39_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail42.Text = (CommonUtil.StringToDoubleVal(txtDetail37.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail38.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail39.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail40.Text.Replace(",", ""))
                 + CommonUtil.StringToDoubleVal(txtDetail41.Text.Replace(",", ""))).ToString();
        }
        private void txtDetail40_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);
            
            txtDetail42.Text = (CommonUtil.StringToDoubleVal(txtDetail37.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail38.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail39.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail40.Text.Replace(",", ""))
                 + CommonUtil.StringToDoubleVal(txtDetail41.Text.Replace(",", ""))).ToString();
        }
        private void txtDetail41_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);
            
            txtDetail42.Text = (CommonUtil.StringToDoubleVal(txtDetail37.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail38.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail39.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail40.Text.Replace(",", ""))
                 + CommonUtil.StringToDoubleVal(txtDetail41.Text.Replace(",", ""))).ToString();
        }

        private void txtDetail43_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);
            
            txtDetail53.Text = (CommonUtil.StringToDoubleVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail44_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail53.Text = (CommonUtil.StringToDoubleVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail45_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail53.Text = (CommonUtil.StringToDoubleVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail46_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail53.Text = (CommonUtil.StringToDoubleVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail47_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail53.Text = (CommonUtil.StringToDoubleVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail48_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail53.Text = (CommonUtil.StringToDoubleVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail49_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail53.Text = (CommonUtil.StringToDoubleVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail50_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail53.Text = (CommonUtil.StringToDoubleVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail51_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail53.Text = (CommonUtil.StringToDoubleVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail52_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail53.Text = (CommonUtil.StringToDoubleVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail54_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail56.Text = (CommonUtil.StringToDoubleVal(txtDetail54.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail55.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail55_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail56.Text = (CommonUtil.StringToDoubleVal(txtDetail54.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail55.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail57_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail63.Text = (CommonUtil.StringToDoubleVal(txtDetail57.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail58.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail59.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail60.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail61.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail62.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail58_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail63.Text = (CommonUtil.StringToDoubleVal(txtDetail57.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail58.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail59.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail60.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail61.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail62.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail59_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail63.Text = (CommonUtil.StringToDoubleVal(txtDetail57.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail58.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail59.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail60.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail61.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail62.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail60_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail63.Text = (CommonUtil.StringToDoubleVal(txtDetail57.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail58.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail59.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail60.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail61.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail62.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail61_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail63.Text = (CommonUtil.StringToDoubleVal(txtDetail57.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail58.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail59.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail60.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail61.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail62.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail62_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail63.Text = (CommonUtil.StringToDoubleVal(txtDetail57.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail58.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail59.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail60.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail61.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail62.Text.Replace(",", ""))
                 ).ToString();
        }


        private void txtDetail64_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail72.Text = (CommonUtil.StringToDoubleVal(txtDetail64.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail65.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail66.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail67.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail68.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail69.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail70.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail71.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail65_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail72.Text = (CommonUtil.StringToDoubleVal(txtDetail64.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail65.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail66.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail67.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail68.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail69.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail70.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail71.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail66_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail72.Text = (CommonUtil.StringToDoubleVal(txtDetail64.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail65.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail66.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail67.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail68.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail69.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail70.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail71.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail67_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail72.Text = (CommonUtil.StringToDoubleVal(txtDetail64.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail65.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail66.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail67.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail68.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail69.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail70.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail71.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail68_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail72.Text = (CommonUtil.StringToDoubleVal(txtDetail64.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail65.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail66.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail67.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail68.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail69.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail70.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail71.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail69_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail72.Text = (CommonUtil.StringToDoubleVal(txtDetail64.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail65.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail66.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail67.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail68.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail69.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail70.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail71.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail70_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail72.Text = (CommonUtil.StringToDoubleVal(txtDetail64.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail65.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail66.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail67.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail68.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail69.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail70.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail71.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail71_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

            txtDetail72.Text = (CommonUtil.StringToDoubleVal(txtDetail64.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail65.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail66.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail67.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail68.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail69.Text.Replace(",", ""))
                + CommonUtil.StringToDoubleVal(txtDetail70.Text.Replace(",", "")) + CommonUtil.StringToDoubleVal(txtDetail71.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail42_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);


        }

        private void txtDetail53_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);


        }

        private void txtDetail56_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);


        }

        private void txtDetail63_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);


        }

        private void txtDetail72_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);


        }

        //시뮬레이션 기본입력
        private void txtInput10_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);


        }

        private void txtInput11_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

        }

        private void txtInput12_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

        }

        private void txtInput13_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);


        }

        private void txtInput14_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

        }

        private void txtInput15_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

        }

        private void txtInput16_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

        }

        private void txtInput17_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

        }

        private void txtInput18_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

        }

        private void txtInput24_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

        }

        private void txtInput25_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

        }

        private void txtInput26_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

        }

        private void txtInput27_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

        }

        private void txtInput28_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);

        }


        private void setInput_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender, e);
        }


        private string setTxtInput_TextChanged(object sender, EventArgs e)
        {
            TextBox _TextBox = (sender as TextBox);

            try
            {
                long num = Convert.ToInt64(_TextBox.Text.Replace(",", ""));

                if (_TextBox.Text.Length < 24 && _TextBox.Text.Length > 2)
                {
                    int saveCursor = _TextBox.Text.Length - _TextBox.SelectionStart;

                    //if (_TextBox.Text.Length > 3)
                    _TextBox.Text = String.Format("{0:#,###}", num);

                    if (_TextBox.Text.Length < saveCursor)
                        _TextBox.SelectionStart = 0;
                    else
                        _TextBox.SelectionStart = _TextBox.Text.Length - saveCursor;
                }
                else if (_TextBox.Text.Length > 23)
                {
                    int saveCursor = _TextBox.SelectionStart - 1;
                    _TextBox.Text = _TextBox.Text.Remove(saveCursor, 1);
                    _TextBox.SelectionStart = saveCursor;
                }
            }
            catch
            {
                _TextBox.Text = "0";
                _TextBox.SelectionStart = 1;
            }

            return _TextBox.Text;
        }


        private void txtInput1_Click(object sender, EventArgs e)
        {
            TextBox _TextBox = (sender as TextBox);
            if (_TextBox.Text == "0")
            {
                _TextBox.SelectAll();
            }

        }

        private void txtInput_focusOut(object sender, EventArgs e)
        {
            TextBox _TextBox = (sender as TextBox);

            if (_TextBox.Text == "")
            {
                _TextBox.Text = "0";
            }
        }

        private void txtDetail37_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && e.KeyChar != 8)
            {
                if (e.KeyChar == '-')
                {
                    TextBox _TextBox = (sender as TextBox);
                    int saveCursor = _TextBox.Text.Length - _TextBox.SelectionStart;
                    if (_TextBox.Text.IndexOf('-') == -1)
                        _TextBox.Text = "-" + _TextBox.Text;
                    _TextBox.SelectionStart = _TextBox.Text.Length - saveCursor;
                }
                else if (e.KeyChar == '+')
                {

                    TextBox _TextBox = (sender as TextBox);
                    int saveCursor = _TextBox.Text.Length - _TextBox.SelectionStart;
                    if (_TextBox.Text.IndexOf('-') > -1)
                        _TextBox.Text = _TextBox.Text.Replace("-", "");
                    _TextBox.SelectionStart = _TextBox.Text.Length - saveCursor;
                }
                e.Handled = true;
            }
 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

      

    }
}
