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
            txtDetail[5].Text = (CommonUtil.StringToIntVal(txtDetail[0].Text) + CommonUtil.StringToIntVal(txtDetail[1].Text) 
                + CommonUtil.StringToIntVal(txtDetail[2].Text) + CommonUtil.StringToIntVal(txtDetail[3].Text) 
                + CommonUtil.StringToIntVal(txtDetail[4].Text) ).ToString();
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
            txtDetail[16].Text = (CommonUtil.StringToIntVal(txtDetail[6].Text) + CommonUtil.StringToIntVal(txtDetail[7].Text)
               + CommonUtil.StringToIntVal(txtDetail[8].Text) + CommonUtil.StringToIntVal(txtDetail[9].Text)
               + CommonUtil.StringToIntVal(txtDetail[10].Text) + CommonUtil.StringToIntVal(txtDetail[11].Text)
               + CommonUtil.StringToIntVal(txtDetail[12].Text) + CommonUtil.StringToIntVal(txtDetail[13].Text)
               + CommonUtil.StringToIntVal(txtDetail[14].Text) + CommonUtil.StringToIntVal(txtDetail[15].Text)
               ).ToString();
            txtDetail[17].Text = CDataControl.g_DetailInput.getstr소매_수익_월평균업무취급수수료();
            txtDetail[18].Text = CDataControl.g_DetailInput.getstr소매_수익_직영매장판매수익();
            txtDetail[19].Text = (CommonUtil.StringToIntVal(txtDetail[17].Text) + CommonUtil.StringToIntVal(txtDetail[18].Text)).ToString();

            txtDetail[20].Text = CDataControl.g_DetailInput.getstr소매_비용_직원급여_간부급();
            txtDetail[21].Text = CDataControl.g_DetailInput.getstr소매_비용_직원급여_평사원();
            txtDetail[22].Text = CDataControl.g_DetailInput.getstr소매_비용_지급임차료();
            txtDetail[23].Text = CDataControl.g_DetailInput.getstr소매_비용_지급수수료();
            txtDetail[24].Text = CDataControl.g_DetailInput.getstr소매_비용_판매촉진비();
            txtDetail[25].Text = CDataControl.g_DetailInput.getstr소매_비용_건물관리비();
            txtDetail[26].Text = (CommonUtil.StringToIntVal(txtDetail[20].Text) + CommonUtil.StringToIntVal(txtDetail[21].Text)
                + CommonUtil.StringToIntVal(txtDetail[22].Text) + CommonUtil.StringToIntVal(txtDetail[23].Text)
                + CommonUtil.StringToIntVal(txtDetail[24].Text) + CommonUtil.StringToIntVal(txtDetail[25].Text)).ToString();

            txtDetail[27].Text = CDataControl.g_DetailInput.getstr도소매_비용_복리후생비();
            txtDetail[28].Text = CDataControl.g_DetailInput.getstr도소매_비용_통신비();
            txtDetail[29].Text = CDataControl.g_DetailInput.getstr도소매_비용_공과금();
            txtDetail[30].Text = CDataControl.g_DetailInput.getstr도소매_비용_소모품비();
            txtDetail[31].Text = CDataControl.g_DetailInput.getstr도소매_비용_이자비용();
            txtDetail[32].Text = CDataControl.g_DetailInput.getstr도소매_비용_부가세();
            txtDetail[33].Text = CDataControl.g_DetailInput.getstr도소매_비용_법인세();
            txtDetail[34].Text = CDataControl.g_DetailInput.getstr도소매_비용_기타();
            txtDetail[35].Text = (CommonUtil.StringToIntVal(txtDetail[27].Text) + CommonUtil.StringToIntVal(txtDetail[28].Text)
                + CommonUtil.StringToIntVal(txtDetail[29].Text) + CommonUtil.StringToIntVal(txtDetail[30].Text)
                + CommonUtil.StringToIntVal(txtDetail[31].Text) + CommonUtil.StringToIntVal(txtDetail[32].Text)
                + CommonUtil.StringToIntVal(txtDetail[33].Text) + CommonUtil.StringToIntVal(txtDetail[34].Text)
                ).ToString();


            txtDetail[36].Text  = CDataControl.g_DetailInput.getstr도매_수익_월평균관리수수료();
            txtDetail[37].Text =  CDataControl.g_DetailInput.getstr도매_수익_CS관리수수료();
            txtDetail[38].Text =  CDataControl.g_DetailInput.getstr도매_수익_사업자모델매입관련추가수익();
            txtDetail[39].Text =  CDataControl.g_DetailInput.getstr도매_수익_유통모델매입관련추가수익_현금DC();
            txtDetail[40].Text =  CDataControl.g_DetailInput.getstr도매_수익_유통모델매입관련추가수익_VolumeDC();
            txtDetail[41].Text =  (CommonUtil.StringToIntVal(txtDetail[0].Text) + CommonUtil.StringToIntVal(txtDetail[1].Text) 
                + CommonUtil.StringToIntVal(txtDetail[2].Text) + CommonUtil.StringToIntVal(txtDetail[3].Text)  
                + CommonUtil.StringToIntVal(txtDetail[4].Text) ).ToString(); 
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
            txtDetail[52].Text =  (CommonUtil.StringToIntVal(txtDetail[6].Text) + CommonUtil.StringToIntVal(txtDetail[7].Text)
               + CommonUtil.StringToIntVal(txtDetail[8].Text) + CommonUtil.StringToIntVal(txtDetail[9].Text)  
               + CommonUtil.StringToIntVal(txtDetail[10].Text) + CommonUtil.StringToIntVal(txtDetail[11].Text)  
               + CommonUtil.StringToIntVal(txtDetail[12].Text) + CommonUtil.StringToIntVal(txtDetail[13].Text) 
               + CommonUtil.StringToIntVal(txtDetail[14].Text) + CommonUtil.StringToIntVal(txtDetail[15].Text) 
               ).ToString();
            txtDetail[53].Text  = CDataControl.g_DetailInput.getstr소매_수익_월평균업무취급수수료();
            txtDetail[54].Text =  CDataControl.g_DetailInput.getstr소매_수익_직영매장판매수익();
            txtDetail[55].Text =  (CommonUtil.StringToIntVal(txtDetail[17].Text) + CommonUtil.StringToIntVal(txtDetail[18].Text)).ToString();
            txtDetail[56].Text =  CDataControl.g_DetailInput.getstr소매_비용_직원급여_간부급();
            txtDetail[57].Text =  CDataControl.g_DetailInput.getstr소매_비용_직원급여_평사원();
            txtDetail[58].Text =  CDataControl.g_DetailInput.getstr소매_비용_지급임차료();
            txtDetail[59].Text =  CDataControl.g_DetailInput.getstr소매_비용_지급수수료();
            txtDetail[60].Text =  CDataControl.g_DetailInput.getstr소매_비용_판매촉진비();
            txtDetail[61].Text =  CDataControl.g_DetailInput.getstr소매_비용_건물관리비();
            txtDetail[62].Text =  (CommonUtil.StringToIntVal(txtDetail[20].Text) + CommonUtil.StringToIntVal(txtDetail[21].Text)
                + CommonUtil.StringToIntVal(txtDetail[22].Text) + CommonUtil.StringToIntVal(txtDetail[23].Text) 
                + CommonUtil.StringToIntVal(txtDetail[24].Text) + CommonUtil.StringToIntVal(txtDetail[25].Text)).ToString();

            txtDetail[63].Text = CDataControl.g_DetailInput.getstr도소매_비용_복리후생비();
            txtDetail[64].Text = CDataControl.g_DetailInput.getstr도소매_비용_통신비();
            txtDetail[65].Text = CDataControl.g_DetailInput.getstr도소매_비용_공과금();
            txtDetail[66].Text = CDataControl.g_DetailInput.getstr도소매_비용_소모품비();
            txtDetail[67].Text = CDataControl.g_DetailInput.getstr도소매_비용_이자비용();
            txtDetail[68].Text = CDataControl.g_DetailInput.getstr도소매_비용_부가세();
            txtDetail[69].Text = CDataControl.g_DetailInput.getstr도소매_비용_법인세();
            txtDetail[70].Text = CDataControl.g_DetailInput.getstr도소매_비용_기타();
            txtDetail[71].Text =  (CommonUtil.StringToIntVal(txtDetail[27].Text) + CommonUtil.StringToIntVal(txtDetail[28].Text)
                + CommonUtil.StringToIntVal(txtDetail[29].Text) + CommonUtil.StringToIntVal(txtDetail[30].Text) 
                + CommonUtil.StringToIntVal(txtDetail[31].Text) + CommonUtil.StringToIntVal(txtDetail[32].Text) 
                + CommonUtil.StringToIntVal(txtDetail[33].Text) + CommonUtil.StringToIntVal(txtDetail[34].Text) 
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
            CBusinessData di = CDataControl.g_DetailInput;
            CResultData[] rdts = new CResultData[] { CDataControl.g_SimResultStoreTotal, CDataControl.g_SimResultFutureTotal };
            CResultData[] rds = new CResultData[] { CDataControl.g_SimResultStore, CDataControl.g_SimResultFuture };
            CResultData rdt = null;
            CResultData rd = null;

            bi.set지역(area);
            bi.set대리점(beanch);
            bi.set마케터(name);

            String[] txtWrite = new String[14] { txtInput10.Text, txtInput11.Text, txtInput12.Text, txtInput13.Text, txtInput14.Text,  
                txtInput15.Text, txtInput16.Text, txtInput17.Text, txtInput18.Text, txtInput24.Text, txtInput25.Text, txtInput26.Text, txtInput27.Text, txtInput28.Text};
            CDataControl.g_SimBasicInput.setArrData_BasicInput(txtWrite);

            String[] txtWrite2 = new String[31]  { txtDetail37.Text, txtDetail38.Text, txtDetail39.Text, txtDetail40.Text, txtDetail41.Text,
                txtDetail43.Text, txtDetail44.Text, txtDetail45.Text, txtDetail46.Text, txtDetail47.Text,
                txtDetail12.Text, txtDetail15.Text, txtDetail16.Text, txtDetail17.Text, txtDetail18.Text,
                txtDetail19.Text, txtDetail20.Text, txtDetail23.Text, txtDetail24.Text, txtDetail25.Text, 
                txtDetail26.Text, txtDetail27.Text, txtDetail28.Text, txtDetail29.Text, txtDetail30.Text,            
                txtDetail31.Text, txtDetail32.Text, txtDetail33.Text, txtDetail34.Text, txtDetail35.Text, txtDetail36.Text
            };
            di.setArrData_DetailInput(txtWrite2);
            CommonUtil.ReadFileManagerToData();

            for (int i = 0; i < rdts.Length; i++)
            {
                //  당대리점 결과(현재:0, 미래:1)
                rdt = rdts[i];
                rd = rds[i];
                //      도매
                //          총액
                //              수익
                rdt.set도매_수익_가입자관리수수료(i == 0 ? di.get도매_수익_월평균관리수수료() : CommonUtil.Division(di.get도매_수익_월평균관리수수료(), bi.get도매_누적가입자수()) * 18 * bi.get월평균판매대수_소계_합계());
                rdt.set도매_수익_CS관리수수료(di.get도매_수익_CS관리수수료());
                rdt.set도매_수익_사업자모델매입에따른추가수익(di.get도매_수익_사업자모델매입관련추가수익());
                rdt.set도매_수익_유통모델매입에따른추가수익_현금_Volume(di.get도매_수익_유통모델매입관련추가수익_현금DC() + di.get도매_수익_유통모델매입관련추가수익_VolumeDC());
                rdt.도매_수익_소계 = rdt.get도매_수익_가입자관리수수료() + rdt.get도매_수익_CS관리수수료() + rdt.get도매_수익_사업자모델매입에따른추가수익() + rdt.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
                //              비용
                rdt.set도매_비용_대리점투자비용(di.get도매_비용_대리점투자금액_신규() * bi.get도매_월평균판매대수_신규() + di.get도매_비용_대리점투자금액_기변() * bi.get도매_월평균판매대수_기변());
                rdt.set도매_비용_인건비_급여_복리후생비(di.get도매_비용_직원급여_간부급() * bi.get도매_직원수_간부급() + bi.get도매_직원수_평사원() * di.get도매_비용_직원급여_평사원());
                rdt.set도매_비용_임차료(di.get도매_비용_지급임차료());
                rdt.set도매_비용_이자비용(CommonUtil.Division(di.get도소매_비용_이자비용(), bi.get월평균판매대수_소계_합계()) * bi.get도매_월평균판매대수_소계());
                rdt.set도매_비용_부가세(CommonUtil.Division(di.get도소매_비용_부가세(), bi.get월평균판매대수_소계_합계()) * bi.get도매_월평균판매대수_소계());
                rdt.set도매_비용_법인세(CommonUtil.Division(di.get도소매_비용_법인세(), bi.get월평균판매대수_소계_합계()) * bi.get도매_월평균판매대수_소계());
                rdt.set도매_비용_기타판매관리비(di.get도매_비용_운반비() + di.get도매_비용_차량유지비() + di.get도매_비용_지급수수료() + di.get도매_비용_판매촉진비() + di.get도매_비용_건물관리비() + (CommonUtil.Division((di.get도소매_비용_복리후생비() + di.get도소매_비용_통신비() + di.get도소매_비용_공과금() + di.get도소매_비용_소모품비() + di.get도소매_비용_기타()), bi.get월평균판매대수_소계_합계()) * bi.get도매_월평균판매대수_소계()));
                rdt.도매_비용_소계 = rdt.get도매_비용_대리점투자비용() + rdt.get도매_비용_인건비_급여_복리후생비() + rdt.get도매_비용_임차료() + rdt.get도매_비용_이자비용() + rdt.get도매_비용_부가세() + rdt.get도매_비용_법인세() + rdt.get도매_비용_기타판매관리비();
                rdt.도매손익계 = rdt.도매_수익_소계 - rdt.도매_비용_소계;
                //          단위당 금액
                //              수익
                rd.set도매_수익_가입자관리수수료(CommonUtil.Division(rdt.get도매_수익_가입자관리수수료(), bi.get도매_월평균판매대수_소계()));
                rd.set도매_수익_CS관리수수료(CommonUtil.Division(rdt.get도매_수익_CS관리수수료(), bi.get도매_월평균판매대수_소계()));
                rd.set도매_수익_사업자모델매입에따른추가수익(CommonUtil.Division(rdt.get도매_수익_사업자모델매입에따른추가수익(), bi.get도매_월평균판매대수_소계()));
                rd.set도매_수익_유통모델매입에따른추가수익_현금_Volume(CommonUtil.Division(rdt.get도매_수익_유통모델매입에따른추가수익_현금_Volume(), bi.get도매_월평균판매대수_소계()));
                rd.도매_수익_소계 = CommonUtil.Division(rdt.도매_수익_소계, bi.get도매_월평균판매대수_소계());
                //              비용
                rd.set도매_비용_대리점투자비용(CommonUtil.Division(rdt.get도매_비용_대리점투자비용(), bi.get도매_월평균판매대수_소계()));
                rd.set도매_비용_인건비_급여_복리후생비(CommonUtil.Division(rdt.get도매_비용_인건비_급여_복리후생비(), bi.get도매_월평균판매대수_소계()));
                rd.set도매_비용_임차료(CommonUtil.Division(rdt.get도매_비용_임차료(), bi.get도매_월평균판매대수_소계()));
                rd.set도매_비용_이자비용(CommonUtil.Division(rdt.get도매_비용_이자비용(), bi.get도매_월평균판매대수_소계()));
                rd.set도매_비용_부가세(CommonUtil.Division(rdt.get도매_비용_부가세(), bi.get도매_월평균판매대수_소계()));
                rd.set도매_비용_법인세(CommonUtil.Division(rdt.get도매_비용_법인세(), bi.get도매_월평균판매대수_소계()));
                rd.set도매_비용_기타판매관리비(CommonUtil.Division(rdt.get도매_비용_기타판매관리비(), bi.get도매_월평균판매대수_소계()));
                rd.도매_비용_소계 = CommonUtil.Division(rdt.도매_비용_소계, bi.get도매_월평균판매대수_소계());
                rd.도매손익계 = CommonUtil.Division(rdt.도매손익계, bi.get도매_월평균판매대수_소계());
                //      소매
                //          총액
                //              수익
                rdt.set소매_수익_업무취급수수료(di.get소매_수익_월평균업무취급수수료());
                rdt.set소매_수익_직영매장판매수익(di.get소매_수익_직영매장판매수익());
                rdt.소매_수익_소계 = rdt.get소매_수익_업무취급수수료() + rdt.get소매_수익_직영매장판매수익();
                //              비용
                rdt.set소매_비용_인건비_급여_복리후생비(di.get소매_비용_직원급여_간부급() * bi.get소매_직원수_간부급() + bi.get소매_직원수_평사원() * di.get소매_비용_직원급여_평사원());
                rdt.set소매_비용_임차료(di.get소매_비용_지급임차료());
                rdt.set소매_비용_이자비용(CommonUtil.Division(di.get도소매_비용_이자비용(), bi.get월평균판매대수_소계_합계()) * bi.get소매_월평균판매대수_소계());
                rdt.set소매_비용_부가세(CommonUtil.Division(di.get도소매_비용_부가세(), bi.get월평균판매대수_소계_합계()) * bi.get소매_월평균판매대수_소계());
                rdt.set소매_비용_법인세(CommonUtil.Division(di.get도소매_비용_법인세(), bi.get월평균판매대수_소계_합계()) * bi.get소매_월평균판매대수_소계());
                rdt.set소매_비용_기타판매관리비((di.get소매_비용_지급수수료() + di.get소매_비용_판매촉진비() + di.get소매_비용_건물관리비()) + (CommonUtil.Division((di.get도소매_비용_복리후생비() + di.get도소매_비용_통신비() + di.get도소매_비용_공과금() + di.get도소매_비용_소모품비() + di.get도소매_비용_기타()), bi.get월평균판매대수_소계_합계()) * bi.get소매_월평균판매대수_소계()));
                rdt.소매_비용_소계 = rdt.get소매_비용_인건비_급여_복리후생비() + rdt.get소매_비용_임차료() + rdt.get소매_비용_이자비용() + rdt.get소매_비용_부가세() + rdt.get소매_비용_법인세() + rdt.get소매_비용_기타판매관리비();
                rdt.소매손익계 = rdt.소매_수익_소계 - rdt.소매_비용_소계;
                rdt.점별손익추정 = CommonUtil.Division(rdt.소매손익계, bi.get거래선수_직영점_합계());
                //          단위당 금액
                //              수익
                rd.set소매_수익_업무취급수수료(CommonUtil.Division(rdt.get소매_수익_업무취급수수료(), bi.get소매_월평균판매대수_소계()));
                rd.set소매_수익_직영매장판매수익(CommonUtil.Division(rdt.get소매_수익_직영매장판매수익(), bi.get소매_월평균판매대수_소계()));
                rd.소매_수익_소계 = CommonUtil.Division(rdt.소매_수익_소계, bi.get소매_월평균판매대수_소계());
                //              비용
                rd.set소매_비용_인건비_급여_복리후생비(CommonUtil.Division(rdt.get소매_비용_인건비_급여_복리후생비(), bi.get소매_월평균판매대수_소계()));
                rd.set소매_비용_임차료(CommonUtil.Division(rdt.get소매_비용_임차료(), bi.get소매_월평균판매대수_소계()));
                rd.set소매_비용_이자비용(CommonUtil.Division(rdt.get소매_비용_이자비용(), bi.get소매_월평균판매대수_소계()));
                rd.set소매_비용_부가세(CommonUtil.Division(rdt.get소매_비용_부가세(), bi.get소매_월평균판매대수_소계()));
                rd.set소매_비용_법인세(CommonUtil.Division(rdt.get소매_비용_법인세(), bi.get소매_월평균판매대수_소계()));
                rd.set소매_비용_기타판매관리비(CommonUtil.Division(rdt.get소매_비용_기타판매관리비(), bi.get소매_월평균판매대수_소계()));
                rd.소매_비용_소계 = CommonUtil.Division(rdt.소매_비용_소계, bi.get소매_월평균판매대수_소계());
                rd.소매손익계 = CommonUtil.Division(rdt.소매손익계, bi.get소매_월평균판매대수_소계());
                //      전체
                //          총액
                //              수익
                rdt.set전체_수익_가입자관리수수료(rdt.get도매_수익_가입자관리수수료());
                rdt.set전체_수익_CS관리수수료(rdt.get도매_수익_CS관리수수료());
                rdt.set전체_수익_업무취급수수료(rdt.get소매_수익_업무취급수수료());
                rdt.set전체_수익_사업자모델매입에따른추가수익(rdt.get도매_수익_사업자모델매입에따른추가수익());
                rdt.set전체_수익_유통모델매입에따른추가수익_현금_Volume(rdt.get도매_수익_유통모델매입에따른추가수익_현금_Volume());
                rdt.set전체_수익_직영매장판매수익(rdt.get소매_수익_직영매장판매수익());
                rdt.전체_수익_소계 = rdt.get전체_수익_가입자관리수수료() + rdt.get전체_수익_CS관리수수료() + rdt.get전체_수익_업무취급수수료() + rdt.get전체_수익_사업자모델매입에따른추가수익() + rdt.get전체_수익_유통모델매입에따른추가수익_현금_Volume() + rdt.get전체_수익_직영매장판매수익();
                //              비용
                rdt.set전체_비용_대리점투자비용(rdt.get도매_비용_대리점투자비용());
                rdt.set전체_비용_인건비_급여_복리후생비(rdt.get도매_비용_인건비_급여_복리후생비() + rdt.get소매_비용_인건비_급여_복리후생비());
                rdt.set전체_비용_임차료(rdt.get도매_비용_임차료() + rdt.get소매_비용_임차료());
                rdt.set전체_비용_이자비용(di.get도소매_비용_이자비용());
                rdt.set전체_비용_부가세(di.get도소매_비용_부가세());
                rdt.set전체_비용_법인세(di.get도소매_비용_법인세());
                rdt.set전체_비용_기타판매관리비(di.get도매_비용_운반비() + di.get도매_비용_차량유지비() + di.get도매_비용_지급수수료() + di.get도매_비용_판매촉진비() + di.get도매_비용_건물관리비() + di.get소매_비용_지급수수료() + di.get소매_비용_판매촉진비() + di.get소매_비용_건물관리비() + di.get도소매_비용_복리후생비() + di.get도소매_비용_통신비() + di.get도소매_비용_공과금() + di.get도소매_비용_소모품비() + di.get도소매_비용_기타());
                rdt.전체_비용_소계 = rdt.get전체_비용_대리점투자비용() + rdt.get전체_비용_인건비_급여_복리후생비() + rdt.get전체_비용_임차료() + rdt.get전체_비용_이자비용() + rdt.get전체_비용_부가세() + rdt.get전체_비용_법인세() + rdt.get전체_비용_기타판매관리비();
                rdt.전체손익계 = rdt.전체_수익_소계 - rdt.전체_비용_소계;
                //          단위당 금액
                //              수익
                rd.set전체_수익_가입자관리수수료(CommonUtil.Division(rdt.get전체_수익_가입자관리수수료(), bi.get월평균판매대수_소계_합계()));
                rd.set전체_수익_CS관리수수료(CommonUtil.Division(rdt.get전체_수익_CS관리수수료(), bi.get월평균판매대수_소계_합계()));
                rd.set전체_수익_업무취급수수료(CommonUtil.Division(rdt.get전체_수익_업무취급수수료(), bi.get월평균판매대수_소계_합계()));
                rd.set전체_수익_사업자모델매입에따른추가수익(CommonUtil.Division(rdt.get전체_수익_사업자모델매입에따른추가수익(), bi.get월평균판매대수_소계_합계()));
                rd.set전체_수익_유통모델매입에따른추가수익_현금_Volume(CommonUtil.Division(rdt.get전체_수익_유통모델매입에따른추가수익_현금_Volume(), bi.get월평균판매대수_소계_합계()));
                rd.set전체_수익_직영매장판매수익(CommonUtil.Division(rdt.get전체_수익_직영매장판매수익(), bi.get월평균판매대수_소계_합계()));
                rd.전체_수익_소계 = CommonUtil.Division(rdt.전체_수익_소계, bi.get월평균판매대수_소계_합계());
                //              비용
                rd.set전체_비용_대리점투자비용(CommonUtil.Division(rdt.get전체_비용_대리점투자비용(), bi.get월평균판매대수_소계_합계()));
                rd.set전체_비용_인건비_급여_복리후생비(CommonUtil.Division(rdt.get전체_비용_인건비_급여_복리후생비(), bi.get월평균판매대수_소계_합계()));
                rd.set전체_비용_임차료(CommonUtil.Division(rdt.get전체_비용_임차료(), bi.get월평균판매대수_소계_합계()));
                rd.set전체_비용_이자비용(CommonUtil.Division(rdt.get전체_비용_이자비용(), bi.get월평균판매대수_소계_합계()));
                rd.set전체_비용_부가세(CommonUtil.Division(rdt.get전체_비용_부가세(), bi.get월평균판매대수_소계_합계()));
                rd.set전체_비용_법인세(CommonUtil.Division(rdt.get전체_비용_법인세(), bi.get월평균판매대수_소계_합계()));
                rd.set전체_비용_기타판매관리비(CommonUtil.Division(rdt.get전체_비용_기타판매관리비(), bi.get월평균판매대수_소계_합계()));
                rd.전체_비용_소계 = CommonUtil.Division(rdt.전체_비용_소계, bi.get월평균판매대수_소계_합계());
                rd.전체손익계 = CommonUtil.Division(rdt.전체손익계, bi.get월평균판매대수_소계_합계());
            }

            //  업계 평균적용 결과
            rdt = CDataControl.g_SimResultBusinessTotal;
            rd = CDataControl.g_SimResultBusiness;
            di = CDataControl.g_BusinessAvg;     // 관리자가 배포한 업계 단위비용
            //      도매
            //          총액
            //              수익
            rdt.set도매_수익_가입자관리수수료(di.get도매_수익_월평균관리수수료() * bi.get도매_누적가입자수());
            rdt.set도매_수익_CS관리수수료(di.get도매_수익_CS관리수수료() * bi.get도매_누적가입자수());
            rdt.set도매_수익_사업자모델매입에따른추가수익(di.get도매_수익_사업자모델매입관련추가수익() * (bi.get월평균판매대수_소계_합계() - bi.get월평균유통모델출고대수_소계_합계()));
            rdt.set도매_수익_유통모델매입에따른추가수익_현금_Volume((di.get도매_수익_유통모델매입관련추가수익_현금DC() + di.get도매_수익_유통모델매입관련추가수익_VolumeDC()) * bi.get월평균유통모델출고대수_소계_합계());
            rdt.도매_수익_소계 = rdt.get도매_수익_가입자관리수수료() + rdt.get도매_수익_CS관리수수료() + rdt.get도매_수익_사업자모델매입에따른추가수익() + rdt.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
            //              비용
            rdt.set도매_비용_대리점투자비용(di.get도매_비용_대리점투자금액_신규() * bi.get도매_월평균판매대수_신규() + di.get도매_비용_대리점투자금액_기변() * bi.get도매_월평균판매대수_기변());
            rdt.set도매_비용_인건비_급여_복리후생비(di.get도매_비용_직원급여_간부급() * bi.get도매_직원수_간부급() + di.get도매_비용_직원급여_평사원() * bi.get도매_직원수_평사원() + di.get도소매_비용_복리후생비() * bi.get도매_직원수_소계());
            rdt.set도매_비용_임차료(di.get도매_비용_지급임차료() * bi.get도매_거래선수_개통사무실());
            rdt.set도매_비용_이자비용(di.get도소매_비용_이자비용() * bi.get도매_월평균판매대수_소계());
            rdt.set도매_비용_부가세(di.get도소매_비용_부가세() * bi.get도매_월평균판매대수_소계());
            rdt.set도매_비용_법인세(di.get도소매_비용_법인세() * bi.get도매_월평균판매대수_소계());
            // '# Detail3. 업계평균vs.해당대리점'!K10+'# Detail3. 업계평균vs.해당대리점'!K11+'# Detail3. 업계평균vs.해당대리점'!K13+'# Detail3. 업계평균vs.해당대리점'!K14+'# Detail3. 업계평균vs.해당대리점'!K15+'# Detail3. 업계평균vs.해당대리점'!K16+'# Detail3. 업계평균vs.해당대리점'!K17+'# Detail3. 업계평균vs.해당대리점'!K18+'# Detail3. 업계평균vs.해당대리점'!K20
            rdt.set도매_비용_기타판매관리비((di.get도매_비용_운반비() + di.get도매_비용_지급수수료() + di.get도매_비용_판매촉진비() + di.get도소매_비용_소모품비() + di.get도소매_비용_기타()) * bi.get도매_월평균판매대수_소계()
                                        + (di.get도매_비용_건물관리비()) * bi.get도매_거래선수_개통사무실()
                                        + (di.get도매_비용_차량유지비() + di.get도소매_비용_통신비() + di.get도소매_비용_공과금()) * bi.get도매_직원수_소계());
            rdt.도매_비용_소계 = rdt.get도매_비용_대리점투자비용() + rdt.get도매_비용_인건비_급여_복리후생비() + rdt.get도매_비용_임차료() + rdt.get도매_비용_이자비용() + rdt.get도매_비용_부가세() + rdt.get도매_비용_법인세() + rdt.get도매_비용_기타판매관리비();
            rdt.도매손익계 = rdt.도매_수익_소계 - rdt.도매_비용_소계;
            //          단위당 금액
            //              수익
            rd.set도매_수익_가입자관리수수료(CommonUtil.Division(rdt.get도매_수익_가입자관리수수료(), bi.get도매_월평균판매대수_소계()));
            rd.set도매_수익_CS관리수수료(CommonUtil.Division(rdt.get도매_수익_CS관리수수료(), bi.get도매_월평균판매대수_소계()));
            rd.set도매_수익_사업자모델매입에따른추가수익(CommonUtil.Division(rdt.get도매_수익_사업자모델매입에따른추가수익(), bi.get도매_월평균판매대수_소계()));
            rd.set도매_수익_유통모델매입에따른추가수익_현금_Volume(CommonUtil.Division(rdt.get도매_수익_유통모델매입에따른추가수익_현금_Volume(), bi.get도매_월평균판매대수_소계()));
            rd.도매_수익_소계 = CommonUtil.Division(rdt.도매_수익_소계, bi.get도매_월평균판매대수_소계());
            //              비용
            rd.set도매_비용_대리점투자비용(CommonUtil.Division(rdt.get도매_비용_대리점투자비용(), bi.get도매_월평균판매대수_소계()));
            rd.set도매_비용_인건비_급여_복리후생비(CommonUtil.Division(rdt.get도매_비용_인건비_급여_복리후생비(), bi.get도매_월평균판매대수_소계()));
            rd.set도매_비용_임차료(CommonUtil.Division(rdt.get도매_비용_임차료(), bi.get도매_월평균판매대수_소계()));
            rd.set도매_비용_이자비용(CommonUtil.Division(rdt.get도매_비용_이자비용(), bi.get도매_월평균판매대수_소계()));
            rd.set도매_비용_부가세(CommonUtil.Division(rdt.get도매_비용_부가세(), bi.get도매_월평균판매대수_소계()));
            rd.set도매_비용_법인세(CommonUtil.Division(rdt.get도매_비용_법인세(), bi.get도매_월평균판매대수_소계()));
            rd.set도매_비용_기타판매관리비(CommonUtil.Division(rdt.get도매_비용_기타판매관리비(), bi.get도매_월평균판매대수_소계()));
            rd.도매_비용_소계 = CommonUtil.Division(rdt.도매_비용_소계, bi.get도매_월평균판매대수_소계());
            rd.도매손익계 = CommonUtil.Division(rdt.도매손익계, bi.get도매_월평균판매대수_소계());
            //      소매
            //          총액
            //              수익
            rdt.set소매_수익_업무취급수수료(di.get소매_수익_월평균업무취급수수료() * bi.get월평균판매대수_소계_합계());
            rdt.set소매_수익_직영매장판매수익(di.get소매_수익_직영매장판매수익() * bi.get소매_월평균판매대수_소계());
            rdt.소매_수익_소계 = rdt.get소매_수익_업무취급수수료() + rdt.get소매_수익_직영매장판매수익();
            //              비용
            rdt.set소매_비용_인건비_급여_복리후생비(di.get소매_비용_직원급여_간부급() * bi.get소매_직원수_간부급() + di.get소매_비용_직원급여_평사원() * bi.get소매_직원수_평사원() + di.get도소매_비용_복리후생비() * bi.get소매_직원수_소계());
            rdt.set소매_비용_임차료(di.get소매_비용_지급임차료() * bi.get소매_거래선수_소계());
            rdt.set소매_비용_이자비용(di.get도소매_비용_이자비용() * bi.get소매_월평균판매대수_소계());
            rdt.set소매_비용_부가세(di.get도소매_비용_부가세() * bi.get소매_월평균판매대수_소계());
            rdt.set소매_비용_법인세(di.get도소매_비용_법인세() * bi.get소매_월평균판매대수_소계());
            // '# Detail3. 업계평균vs.해당대리점'!L10+'# Detail3. 업계평균vs.해당대리점'!L11+'# Detail3. 업계평균vs.해당대리점'!L13+'# Detail3. 업계평균vs.해당대리점'!L14+'# Detail3. 업계평균vs.해당대리점'!L15+'# Detail3. 업계평균vs.해당대리점'!L16+'# Detail3. 업계평균vs.해당대리점'!L17+'# Detail3. 업계평균vs.해당대리점'!L18+'# Detail3. 업계평균vs.해당대리점'!L20
            rdt.set소매_비용_기타판매관리비((di.get소매_비용_지급수수료() + di.get소매_비용_판매촉진비() + di.get도소매_비용_소모품비() + di.get도소매_비용_기타()) * bi.get소매_월평균판매대수_소계()
                                        + (di.get소매_비용_건물관리비()) * bi.get소매_거래선수_소계()
                                        + (di.get도소매_비용_통신비() + di.get도소매_비용_공과금()) * bi.get소매_직원수_소계());
            rdt.소매_비용_소계 = rdt.get소매_비용_인건비_급여_복리후생비() + rdt.get소매_비용_임차료() + rdt.get소매_비용_이자비용() + rdt.get소매_비용_부가세() + rdt.get소매_비용_법인세() + rdt.get소매_비용_기타판매관리비();
            rdt.소매손익계 = rdt.소매_수익_소계 - rdt.소매_비용_소계;
            rdt.점별손익추정 = CommonUtil.Division(rdt.소매손익계, bi.get거래선수_직영점_합계());
            //          단위당 금액
            //              수익
            rd.set소매_수익_업무취급수수료(CommonUtil.Division(rdt.get소매_수익_업무취급수수료(), bi.get소매_월평균판매대수_소계()));
            rd.set소매_수익_직영매장판매수익(CommonUtil.Division(rdt.get소매_수익_직영매장판매수익(), bi.get소매_월평균판매대수_소계()));
            rd.소매_수익_소계 = CommonUtil.Division(rdt.소매_수익_소계, bi.get소매_월평균판매대수_소계());
            //              비용
            rd.set소매_비용_인건비_급여_복리후생비(CommonUtil.Division(rdt.get소매_비용_인건비_급여_복리후생비(), bi.get소매_월평균판매대수_소계()));
            rd.set소매_비용_임차료(CommonUtil.Division(rdt.get소매_비용_임차료(), bi.get소매_월평균판매대수_소계()));
            rd.set소매_비용_이자비용(CommonUtil.Division(rdt.get소매_비용_이자비용(), bi.get소매_월평균판매대수_소계()));
            rd.set소매_비용_부가세(CommonUtil.Division(rdt.get소매_비용_부가세(), bi.get소매_월평균판매대수_소계()));
            rd.set소매_비용_법인세(CommonUtil.Division(rdt.get소매_비용_법인세(), bi.get소매_월평균판매대수_소계()));
            rd.set소매_비용_기타판매관리비(CommonUtil.Division(rdt.get소매_비용_기타판매관리비(), bi.get소매_월평균판매대수_소계()));
            rd.소매_비용_소계 = CommonUtil.Division(rdt.소매_비용_소계, bi.get소매_월평균판매대수_소계());
            rd.소매손익계 = CommonUtil.Division(rdt.소매손익계, bi.get소매_월평균판매대수_소계());
            //      전체
            //          총액
            //              수익
            rdt.set전체_수익_가입자관리수수료(rdt.get도매_수익_가입자관리수수료());
            rdt.set전체_수익_CS관리수수료(rdt.get도매_수익_CS관리수수료());
            rdt.set전체_수익_업무취급수수료(rdt.get소매_수익_업무취급수수료());
            rdt.set전체_수익_사업자모델매입에따른추가수익(rdt.get도매_수익_사업자모델매입에따른추가수익());
            rdt.set전체_수익_유통모델매입에따른추가수익_현금_Volume(rdt.get도매_수익_유통모델매입에따른추가수익_현금_Volume());
            rdt.set전체_수익_직영매장판매수익(rdt.get소매_수익_직영매장판매수익());
            rdt.전체_수익_소계 = rdt.get전체_수익_가입자관리수수료() + rdt.get전체_수익_CS관리수수료() + rdt.get전체_수익_업무취급수수료() + rdt.get전체_수익_사업자모델매입에따른추가수익() + rdt.get전체_수익_유통모델매입에따른추가수익_현금_Volume() + rdt.get전체_수익_직영매장판매수익();
            //              비용
            rdt.set전체_비용_대리점투자비용(rdt.get도매_비용_대리점투자비용());
            rdt.set전체_비용_인건비_급여_복리후생비(rdt.get도매_비용_인건비_급여_복리후생비() + rdt.get소매_비용_인건비_급여_복리후생비());
            rdt.set전체_비용_임차료(rdt.get도매_비용_임차료() + rdt.get소매_비용_임차료());
            rdt.set전체_비용_이자비용(rdt.get도매_비용_이자비용() + rdt.get소매_비용_이자비용());
            rdt.set전체_비용_부가세(rdt.get도매_비용_부가세() + rdt.get소매_비용_부가세());
            rdt.set전체_비용_법인세(rdt.get도매_비용_법인세() + rdt.get소매_비용_법인세());
            rdt.set전체_비용_기타판매관리비(rdt.get도매_비용_기타판매관리비() + rdt.get소매_비용_기타판매관리비());
            rdt.전체_비용_소계 = rdt.get전체_비용_대리점투자비용() + rdt.get전체_비용_인건비_급여_복리후생비() + rdt.get전체_비용_임차료() + rdt.get전체_비용_이자비용() + rdt.get전체_비용_부가세() + rdt.get전체_비용_법인세() + rdt.get전체_비용_기타판매관리비();
            rdt.전체손익계 = rdt.전체_수익_소계 - rdt.전체_비용_소계;
            //          단위당 금액
            //              수익
            rd.set전체_수익_가입자관리수수료(CommonUtil.Division(rdt.get전체_수익_가입자관리수수료(), bi.get월평균판매대수_소계_합계()));
            rd.set전체_수익_CS관리수수료(CommonUtil.Division(rdt.get전체_수익_CS관리수수료(), bi.get월평균판매대수_소계_합계()));
            rd.set전체_수익_업무취급수수료(CommonUtil.Division(rdt.get전체_수익_업무취급수수료(), bi.get월평균판매대수_소계_합계()));
            rd.set전체_수익_사업자모델매입에따른추가수익(CommonUtil.Division(rdt.get전체_수익_사업자모델매입에따른추가수익(), bi.get월평균판매대수_소계_합계()));
            rd.set전체_수익_유통모델매입에따른추가수익_현금_Volume(CommonUtil.Division(rdt.get전체_수익_유통모델매입에따른추가수익_현금_Volume(), bi.get월평균판매대수_소계_합계()));
            rd.set전체_수익_직영매장판매수익(CommonUtil.Division(rdt.get전체_수익_직영매장판매수익(), bi.get월평균판매대수_소계_합계()));
            rd.전체_수익_소계 = CommonUtil.Division(rdt.전체_수익_소계, bi.get월평균판매대수_소계_합계());
            //              비용
            rd.set전체_비용_대리점투자비용(CommonUtil.Division(rdt.get전체_비용_대리점투자비용(), bi.get월평균판매대수_소계_합계()));
            rd.set전체_비용_인건비_급여_복리후생비(CommonUtil.Division(rdt.get전체_비용_인건비_급여_복리후생비(), bi.get월평균판매대수_소계_합계()));
            rd.set전체_비용_임차료(CommonUtil.Division(rdt.get전체_비용_임차료(), bi.get월평균판매대수_소계_합계()));
            rd.set전체_비용_이자비용(CommonUtil.Division(rdt.get전체_비용_이자비용(), bi.get월평균판매대수_소계_합계()));
            rd.set전체_비용_부가세(CommonUtil.Division(rdt.get전체_비용_부가세(), bi.get월평균판매대수_소계_합계()));
            rd.set전체_비용_법인세(CommonUtil.Division(rdt.get전체_비용_법인세(), bi.get월평균판매대수_소계_합계()));
            rd.set전체_비용_기타판매관리비(CommonUtil.Division(rdt.get전체_비용_기타판매관리비(), bi.get월평균판매대수_소계_합계()));
            rd.전체_비용_소계 = CommonUtil.Division(rdt.전체_비용_소계, bi.get월평균판매대수_소계_합계());
            rd.전체손익계 = CommonUtil.Division(rdt.전체손익계, bi.get월평균판매대수_소계_합계());

        }

































        private void SaveAsInput(excel.Worksheet _WorkSheet1, excel.Worksheet _WorkSheet2)
        {


            _WorkSheet1.get_Range("C63", Type.Missing).Value2 = area;
            _WorkSheet1.get_Range("E63", Type.Missing).Value2 = beanch;
            _WorkSheet1.get_Range("G63", Type.Missing).Value2 = name;



            _WorkSheet1.get_Range("F7", Type.Missing).Value2 = txtInput[9].Text;
            _WorkSheet1.get_Range("H7", Type.Missing).Value2 = txtInput[9].Text;

            _WorkSheet1.get_Range("F8", Type.Missing).Value2 = txtInput[10].Text;
            _WorkSheet1.get_Range("F9", Type.Missing).Value2 = txtInput[11].Text;
            _WorkSheet1.get_Range("F10", Type.Missing).Value2 = CommonUtil.StringToIntVal(txtInput[10].Text) + CommonUtil.StringToIntVal(txtInput[11].Text);

            _WorkSheet1.get_Range("F11", Type.Missing).Value2 = txtInput[12].Text;
            _WorkSheet1.get_Range("F12", Type.Missing).Value2 = txtInput[13].Text;
            _WorkSheet1.get_Range("F13", Type.Missing).Value2 = CommonUtil.StringToIntVal(txtInput[12].Text) + CommonUtil.StringToIntVal(txtInput[13].Text);

            _WorkSheet1.get_Range("F14", Type.Missing).Value2 = txtInput[14].Text;
            _WorkSheet1.get_Range("F16", Type.Missing).Value2 = txtInput[15].Text;
            _WorkSheet1.get_Range("F17", Type.Missing).Value2 = CommonUtil.StringToIntVal(txtInput[14].Text) + CommonUtil.StringToIntVal(txtInput[15].Text);

            _WorkSheet1.get_Range("F18", Type.Missing).Value2 = txtInput[16].Text;
            _WorkSheet1.get_Range("F19", Type.Missing).Value2 = txtInput[17].Text;
            _WorkSheet1.get_Range("F20", Type.Missing).Value2 = CommonUtil.StringToIntVal(txtInput[16].Text) + CommonUtil.StringToIntVal(txtInput[17].Text);



            _WorkSheet1.get_Range("G8", Type.Missing).Value2  = txtInput[23].Text;
            _WorkSheet1.get_Range("G9", Type.Missing).Value2  = txtInput[24].Text;
            _WorkSheet1.get_Range("G10", Type.Missing).Value2 = CommonUtil.StringToIntVal(txtInput[23].Text) + CommonUtil.StringToIntVal(txtInput[24].Text);
            _WorkSheet1.get_Range("G15", Type.Missing).Value2 = txtInput[25].Text;
            _WorkSheet1.get_Range("G17", Type.Missing).Value2 = txtInput[25].Text;
            _WorkSheet1.get_Range("G18", Type.Missing).Value2 = txtInput[26].Text;
            _WorkSheet1.get_Range("G19", Type.Missing).Value2 = txtInput[27].Text;
            _WorkSheet1.get_Range("G20", Type.Missing).Value2 = CommonUtil.StringToIntVal(txtInput[26].Text) + CommonUtil.StringToIntVal(txtInput[27].Text);

            _WorkSheet1.get_Range("H8", Type.Missing).Value2 = CommonUtil.StringToIntVal(txtInput[10].Text) + CommonUtil.StringToIntVal(txtInput[23].Text);
            _WorkSheet1.get_Range("H9", Type.Missing).Value2 = CommonUtil.StringToIntVal(txtInput[11].Text) + CommonUtil.StringToIntVal(txtInput[24].Text);
            _WorkSheet1.get_Range("H10", Type.Missing).Value2 = CommonUtil.StringToIntVal(txtInput[10].Text) + CommonUtil.StringToIntVal(txtInput[11].Text)
                + CommonUtil.StringToIntVal(txtInput[23].Text) + CommonUtil.StringToIntVal(txtInput[24].Text);
            _WorkSheet1.get_Range("H11", Type.Missing).Value2 = txtInput[12].Text;
            _WorkSheet1.get_Range("H12", Type.Missing).Value2 = txtInput[13].Text;
            _WorkSheet1.get_Range("H13", Type.Missing).Value2 = CommonUtil.StringToIntVal(txtInput[12].Text) + CommonUtil.StringToIntVal(txtInput[13].Text);
            _WorkSheet1.get_Range("H14", Type.Missing).Value2 = txtInput[14].Text;
            _WorkSheet1.get_Range("H15", Type.Missing).Value2 = txtInput[25].Text;
            _WorkSheet1.get_Range("H16", Type.Missing).Value2 = txtInput[15].Text;
            _WorkSheet1.get_Range("H17", Type.Missing).Value2 = CommonUtil.StringToIntVal(txtInput[14].Text) + CommonUtil.StringToIntVal(txtInput[15].Text) + CommonUtil.StringToIntVal(txtInput[25].Text);
            _WorkSheet1.get_Range("H18", Type.Missing).Value2 = CommonUtil.StringToIntVal(txtInput[16].Text) + CommonUtil.StringToIntVal(txtInput[26].Text);
            _WorkSheet1.get_Range("H19", Type.Missing).Value2 = CommonUtil.StringToIntVal(txtInput[17].Text) + CommonUtil.StringToIntVal(txtInput[27].Text);
            _WorkSheet1.get_Range("H20", Type.Missing).Value2 = CommonUtil.StringToIntVal(txtInput[16].Text) + CommonUtil.StringToIntVal(txtInput[17].Text)
                + CommonUtil.StringToIntVal(txtInput[26].Text) + CommonUtil.StringToIntVal(txtInput[27].Text);


            //상세입력
            _WorkSheet1.get_Range("G26", Type.Missing).Value2 = txtDetail[36].Text;
            _WorkSheet1.get_Range("G27", Type.Missing).Value2 = txtDetail[37].Text;
            _WorkSheet1.get_Range("G28", Type.Missing).Value2 = CommonUtil.StringToIntVal(txtDetail[37].Text) * CommonUtil.QUARTER;
            _WorkSheet1.get_Range("G29", Type.Missing).Value2 = txtDetail[38].Text;
            _WorkSheet1.get_Range("G30", Type.Missing).Value2 = txtDetail[39].Text;
            _WorkSheet1.get_Range("G31", Type.Missing).Value2 = txtDetail[40].Text;

            _WorkSheet1.get_Range("G32", Type.Missing).Value2 = txtDetail[42].Text;
            _WorkSheet1.get_Range("G33", Type.Missing).Value2 = txtDetail[43].Text;
            _WorkSheet1.get_Range("G34", Type.Missing).Value2 = (CommonUtil.StringToIntVal(txtDetail[44].Text) * CommonUtil.StringToIntVal(txtInput[16].Text)).ToString();
            _WorkSheet1.get_Range("G35", Type.Missing).Value2 = (CommonUtil.StringToIntVal(txtDetail[45].Text) * CommonUtil.StringToIntVal(txtInput[17].Text)).ToString();
            _WorkSheet1.get_Range("G36", Type.Missing).Value2 = txtDetail[44].Text;
            _WorkSheet1.get_Range("G37", Type.Missing).Value2 = txtDetail[45].Text;
            _WorkSheet1.get_Range("G38", Type.Missing).Value2 = txtDetail[46].Text;
            _WorkSheet1.get_Range("G39", Type.Missing).Value2 = txtDetail[47].Text;
            _WorkSheet1.get_Range("G40", Type.Missing).Value2 = txtDetail[48].Text;
            _WorkSheet1.get_Range("G41", Type.Missing).Value2 = txtDetail[49].Text;
            _WorkSheet1.get_Range("G42", Type.Missing).Value2 = txtDetail[50].Text;
            _WorkSheet1.get_Range("G43", Type.Missing).Value2 = txtDetail[51].Text;

            _WorkSheet1.get_Range("G44", Type.Missing).Value2 = txtDetail[53].Text;
            _WorkSheet1.get_Range("G45", Type.Missing).Value2 = txtDetail[54].Text;
            _WorkSheet1.get_Range("G46", Type.Missing).Value2 = (CommonUtil.StringToIntVal(txtDetail[56].Text) * CommonUtil.StringToIntVal(txtInput[26].Text)).ToString();
            _WorkSheet1.get_Range("G47", Type.Missing).Value2 = (CommonUtil.StringToIntVal(txtDetail[57].Text) * CommonUtil.StringToIntVal(txtInput[27].Text)).ToString();

            _WorkSheet1.get_Range("G48", Type.Missing).Value2 = txtDetail[56].Text;
            _WorkSheet1.get_Range("G49", Type.Missing).Value2 = txtDetail[57].Text;

            _WorkSheet1.get_Range("G50", Type.Missing).Value2 = txtDetail[58].Text;
            _WorkSheet1.get_Range("G51", Type.Missing).Value2 = txtDetail[59].Text;
            _WorkSheet1.get_Range("G52", Type.Missing).Value2 = txtDetail[60].Text;
            _WorkSheet1.get_Range("G53", Type.Missing).Value2 = txtDetail[61].Text;


            _WorkSheet1.get_Range("G54", Type.Missing).Value2 = txtDetail[63].Text;
            _WorkSheet1.get_Range("G55", Type.Missing).Value2 = txtDetail[64].Text;
            _WorkSheet1.get_Range("G56", Type.Missing).Value2 = txtDetail[65].Text;
            _WorkSheet1.get_Range("G57", Type.Missing).Value2 = txtDetail[66].Text;
            _WorkSheet1.get_Range("G58", Type.Missing).Value2 = txtDetail[67].Text;
            _WorkSheet1.get_Range("G59", Type.Missing).Value2 = txtDetail[68].Text;
            _WorkSheet1.get_Range("G60", Type.Missing).Value2 = txtDetail[69].Text;
            _WorkSheet1.get_Range("G61", Type.Missing).Value2 = txtDetail[70].Text;


            //관리자 파일을 읽어 넣는다
            try
            {
                string csv = System.IO.File.ReadAllText(CommonUtil.defaultManagerName);
                string[] splitedCsv = csv.Split(',');
                for (int i = 0; i < txtMangeInput.Length; i++)
                {
                    txtMangeInput[i] = splitedCsv[i];
                }
            }
            catch (Exception ex)
            {
                // 파일이 없음
                for (int i = 0; i < txtMangeInput.Length; i++)
                {
                    txtMangeInput[i] = 0.ToString();
                }
            }

            Int64 sumSubDE = 0;
            _WorkSheet2.get_Range("E7", Type.Missing).Value2 = txtMangeInput[0];
            sumSubDE += Convert.ToInt64(txtMangeInput[0]);
            _WorkSheet2.get_Range("E8", Type.Missing).Value2 = txtMangeInput[1];
            sumSubDE += Convert.ToInt64(txtMangeInput[1]);
            _WorkSheet2.get_Range("E9", Type.Missing).Value2 = txtMangeInput[15];
            sumSubDE += Convert.ToInt64(txtMangeInput[16]);
            _WorkSheet2.get_Range("E10", Type.Missing).Value2 = txtMangeInput[2];
            sumSubDE += Convert.ToInt64(txtMangeInput[2]);
            _WorkSheet2.get_Range("E11", Type.Missing).Value2 = Convert.ToInt64(txtMangeInput[3]) + Convert.ToInt64(txtMangeInput[4]);
            sumSubDE += Convert.ToInt64(Convert.ToInt64(txtMangeInput[3]) + Convert.ToInt64(txtMangeInput[4]));
            _WorkSheet2.get_Range("E12", Type.Missing).Value2 = txtMangeInput[16];
            sumSubDE += Convert.ToInt64(txtMangeInput[17]);
            _WorkSheet2.get_Range("E13", Type.Missing).Value2 = sumSubDE;

            Int64 sumSubCo = 0;
            _WorkSheet2.get_Range("E14", Type.Missing).Value2 = CommonUtil.Division((Convert.ToInt64(txtMangeInput[5]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F8", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[6]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F9", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division((Convert.ToInt64(txtMangeInput[5]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F8", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[6]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F9", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2)));

            _WorkSheet2.get_Range("E15", Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((Convert.ToInt64(txtMangeInput[7]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F18", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[8]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F19", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[17]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G18", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[18]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G19", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[23]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2)));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division((Convert.ToInt64(txtMangeInput[7]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F18", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[8]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F19", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[17]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G18", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[18]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G19", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[23]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2)));


            _WorkSheet2.get_Range("E16", Type.Missing).Value2 = Convert.ToInt64(txtMangeInput[9]) + Convert.ToInt64(txtMangeInput[19]);
            sumSubCo += Convert.ToInt64(txtMangeInput[9]) + Convert.ToInt64(txtMangeInput[19]);
            _WorkSheet2.get_Range("E17", Type.Missing).Value2 = Convert.ToInt64(txtMangeInput[27]);
            sumSubCo += Convert.ToInt64(txtMangeInput[27]);
            _WorkSheet2.get_Range("E18", Type.Missing).Value2 = Convert.ToInt64(txtMangeInput[28]);
            sumSubCo += Convert.ToInt64(txtMangeInput[28]);
            _WorkSheet2.get_Range("E19", Type.Missing).Value2 = Convert.ToInt64(txtMangeInput[29]);
            sumSubCo += Convert.ToInt64(txtMangeInput[29]);

            _WorkSheet2.get_Range("E20", Type.Missing).Value2 = CommonUtil.StringToIntVal(txtMangeInput[10]) + CommonUtil.StringToIntVal(txtMangeInput[11])
                + CommonUtil.StringToIntVal(txtMangeInput[12]) + CommonUtil.StringToIntVal(txtMangeInput[13])
                + CommonUtil.StringToIntVal(txtMangeInput[14]) + CommonUtil.StringToIntVal(txtMangeInput[20])
                + CommonUtil.StringToIntVal(txtMangeInput[21]) + CommonUtil.StringToIntVal(txtMangeInput[22])
                + CommonUtil.StringToIntVal(txtMangeInput[24]) + CommonUtil.StringToIntVal(txtMangeInput[25])
                + CommonUtil.StringToIntVal(txtMangeInput[26]) + CommonUtil.StringToIntVal(txtMangeInput[30]);
            sumSubCo += CommonUtil.StringToIntVal(txtMangeInput[10]) + CommonUtil.StringToIntVal(txtMangeInput[11])
                + CommonUtil.StringToIntVal(txtMangeInput[12]) + CommonUtil.StringToIntVal(txtMangeInput[13])
                + CommonUtil.StringToIntVal(txtMangeInput[14]) + CommonUtil.StringToIntVal(txtMangeInput[20])
                + CommonUtil.StringToIntVal(txtMangeInput[21]) + CommonUtil.StringToIntVal(txtMangeInput[22])
                + CommonUtil.StringToIntVal(txtMangeInput[24]) + CommonUtil.StringToIntVal(txtMangeInput[25])
                + CommonUtil.StringToIntVal(txtMangeInput[26]) + CommonUtil.StringToIntVal(txtMangeInput[30]);
            _WorkSheet2.get_Range("E21", Type.Missing).Value2 = sumSubCo;
            _WorkSheet2.get_Range("E22", Type.Missing).Value2 = sumSubDE - sumSubCo;

            sumSubDE = 0;
            //도매 수익
            _WorkSheet2.get_Range("E28", Type.Missing).Value2 = txtMangeInput[0];
            sumSubDE += Convert.ToInt64(txtMangeInput[0]);
            _WorkSheet2.get_Range("E29", Type.Missing).Value2 = txtMangeInput[1];
            sumSubDE += Convert.ToInt64(txtMangeInput[1]);
            _WorkSheet2.get_Range("E30", Type.Missing).Value2 = txtMangeInput[2];
            sumSubDE += Convert.ToInt64(txtMangeInput[2]);
            _WorkSheet2.get_Range("E31", Type.Missing).Value2 = Convert.ToInt64(txtMangeInput[3]) + Convert.ToInt64(txtMangeInput[4]);
            sumSubDE += Convert.ToInt64(Convert.ToInt64(txtMangeInput[3]) + Convert.ToInt64(txtMangeInput[4]));
            _WorkSheet2.get_Range("E32", Type.Missing).Value2 = sumSubDE;
            //도매비용
            sumSubCo = 0;
            _WorkSheet2.get_Range("E33", Type.Missing).Value2 = CommonUtil.Division((Convert.ToInt64(txtMangeInput[5]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F8", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[6]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F9", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division((Convert.ToInt64(txtMangeInput[5]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F8", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[6]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F9", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2)));
            _WorkSheet2.get_Range("E34", Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((Convert.ToInt64(txtMangeInput[7]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F18", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[8]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F19", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[23]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F20", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F20", Type.Missing).Value2)));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division((Convert.ToInt64(txtMangeInput[7]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F18", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[8]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F19", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[23]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F20", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F20", Type.Missing).Value2)));
            _WorkSheet2.get_Range("E35", Type.Missing).Value2 = txtMangeInput[9];
            sumSubCo += Convert.ToInt64(txtMangeInput[9]);
            _WorkSheet2.get_Range("E36", Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[27], CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[27], CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            _WorkSheet2.get_Range("E37", Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[28], CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[28], CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            _WorkSheet2.get_Range("E38", Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[29], CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[29], CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));

            _WorkSheet2.get_Range("E36", Type.Missing).Value2 = txtMangeInput[27];
            sumSubCo += Convert.ToInt64(txtMangeInput[27]);
            _WorkSheet2.get_Range("E37", Type.Missing).Value2 = txtMangeInput[28];
            sumSubCo += Convert.ToInt64(txtMangeInput[28]);
            _WorkSheet2.get_Range("E38", Type.Missing).Value2 = txtMangeInput[29];
            sumSubCo += Convert.ToInt64(txtMangeInput[29]);
            _WorkSheet2.get_Range("E39", Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtMangeInput[10]) + CommonUtil.StringToIntVal(txtMangeInput[11])
                + CommonUtil.StringToIntVal(txtMangeInput[12]) + CommonUtil.StringToIntVal(txtMangeInput[13])
                + CommonUtil.StringToIntVal(txtMangeInput[14]) + CommonUtil.StringToIntVal(txtMangeInput[24])
                + CommonUtil.StringToIntVal(txtMangeInput[25]) + CommonUtil.StringToIntVal(txtMangeInput[26])
                + CommonUtil.StringToIntVal(txtMangeInput[30])).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2)); ;
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtMangeInput[10]) + CommonUtil.StringToIntVal(txtMangeInput[11])
                + CommonUtil.StringToIntVal(txtMangeInput[12]) + CommonUtil.StringToIntVal(txtMangeInput[13])
                + CommonUtil.StringToIntVal(txtMangeInput[14]) + CommonUtil.StringToIntVal(txtMangeInput[24])
                + CommonUtil.StringToIntVal(txtMangeInput[25]) + CommonUtil.StringToIntVal(txtMangeInput[26])
                + CommonUtil.StringToIntVal(txtMangeInput[30])).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2)); ;

            _WorkSheet2.get_Range("E40", Type.Missing).Value2 = sumSubCo;
            _WorkSheet2.get_Range("E41", Type.Missing).Value2 = sumSubDE - sumSubCo;

            //소매
            sumSubDE = 0;
            _WorkSheet2.get_Range("E46", Type.Missing).Value2 = txtMangeInput[15];
            sumSubDE += Convert.ToInt64(txtMangeInput[16]);
            _WorkSheet2.get_Range("E47", Type.Missing).Value2 = txtMangeInput[16];
            sumSubDE += Convert.ToInt64(txtMangeInput[17]);

            _WorkSheet2.get_Range("E48", Type.Missing).Value2 = sumSubDE;

            sumSubCo = 0;
            _WorkSheet2.get_Range("E49", Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((Convert.ToInt64(txtMangeInput[17]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G18", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[18]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G19", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[23]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G20", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("G20", Type.Missing).Value2)));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division((Convert.ToInt64(txtMangeInput[17]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G18", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[18]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G19", Type.Missing).Value2))
                + Convert.ToInt64(txtMangeInput[23]) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G20", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("G20", Type.Missing).Value2)));
            _WorkSheet2.get_Range("E50", Type.Missing).Value2 = txtMangeInput[19];
            sumSubCo += Convert.ToInt64(txtMangeInput[19]);
            _WorkSheet2.get_Range("E51", Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[27], CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[27], CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            _WorkSheet2.get_Range("E52", Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[28], CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[28], CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            _WorkSheet2.get_Range("E53", Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[29], CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[29], CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));



            _WorkSheet2.get_Range("E54", Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtMangeInput[20]) + CommonUtil.StringToIntVal(txtMangeInput[21])
                + CommonUtil.StringToIntVal(txtMangeInput[22]) + CommonUtil.StringToIntVal(txtMangeInput[24])
                + CommonUtil.StringToIntVal(txtMangeInput[25]) + CommonUtil.StringToIntVal(txtMangeInput[26])
                + CommonUtil.StringToIntVal(txtMangeInput[30])).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtMangeInput[20]) + CommonUtil.StringToIntVal(txtMangeInput[21])
                + CommonUtil.StringToIntVal(txtMangeInput[22]) + CommonUtil.StringToIntVal(txtMangeInput[24])
                + CommonUtil.StringToIntVal(txtMangeInput[25]) + CommonUtil.StringToIntVal(txtMangeInput[26])
                + CommonUtil.StringToIntVal(txtMangeInput[30])).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));

            _WorkSheet2.get_Range("E55", Type.Missing).Value2 = sumSubCo;
            _WorkSheet2.get_Range("E56", Type.Missing).Value2 = sumSubDE - sumSubCo;



            //업계 평균적용 결과(기본입력값으로 자동산출)
            for (int i = 7; i < 58; i++)
            {
                if ((i >= 7 && i < 23) || (i >= 28 && i < 42) || (i >= 46 && i < 57))
                {

                    string ColumnName = "E" + i.ToString();

                    string temp = CommonUtil.NullToString0(_WorkSheet2.get_Range(ColumnName, Type.Missing).Value2);
                    ColumnName = "D" + i.ToString();
                    string txtInput1 = "0";
                    if (i >= 7 && i < 23)
                        txtInput1 = CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2);
                    else if (i >= 28 && i < 42)
                        txtInput1 = CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2);
                    else if (i >= 46 && i < 57)
                        txtInput1 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2);
                    _WorkSheet2.get_Range(ColumnName, Type.Missing).Value2 = CommonUtil.StringToIntVal(temp) * CommonUtil.StringToIntVal(txtInput1);
                }
                if (i == 57)
                {
                    string ColumnName = "E" + i.ToString();
                    _WorkSheet2.get_Range(ColumnName, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("H15", Type.Missing).Value2);
                    ColumnName = "D" + i.ToString();
                    string temp = CommonUtil.NullToString0(_WorkSheet2.get_Range("D" + (i - 1).ToString(), Type.Missing).Value2);
                    _WorkSheet2.get_Range(ColumnName, Type.Missing).Value2 = CommonUtil.Division(temp, CommonUtil.NullToString0(_WorkSheet1.get_Range("H15", Type.Missing).Value2));
                }

            }

            //당대리점 결과(세부항목별 값 입력 결과) 수익계정
            Int64 SumSubBenefitTotal = 0;
            string Column = "I7";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G26", Type.Missing).Value2);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G26", Type.Missing).Value2));
            Column = "J7";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G26", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "I8";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G27", Type.Missing).Value2);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G27", Type.Missing).Value2));
            Column = "J8";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G27", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "I9";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G44", Type.Missing).Value2);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G44", Type.Missing).Value2));
            Column = "J9";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G44", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "I10";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G29", Type.Missing).Value2);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G29", Type.Missing).Value2));
            Column = "J10";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G29", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "I11";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G30", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G31", Type.Missing).Value2))).ToString();
            SumSubBenefitTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G30", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G31", Type.Missing).Value2)));
            Column = "J11";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G30", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G31", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "I12";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G45", Type.Missing).Value2);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G45", Type.Missing).Value2));
            Column = "J12";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G45", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "I13";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = SumSubBenefitTotal.ToString();
            Column = "J13";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(SumSubBenefitTotal.ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));

            //당대리점 결과(세부항목별 값 입력 결과) 비용계정
            Int64 SumSubCostTotal = 0;
            Column = "I14";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G32", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F8", Type.Missing).Value2)) + (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G33", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F9", Type.Missing).Value2))));
            SumSubCostTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G32", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F8", Type.Missing).Value2)) + (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G33", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F9", Type.Missing).Value2))));
            Column = "J14";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G26", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));

            Column = "I15";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G36", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G37", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G48", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G49", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G36", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G37", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G48", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G49", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2)));
            Column = "J15";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G36", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G37", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G48", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G49", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "I16";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G38", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G50", Type.Missing).Value2)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G38", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G50", Type.Missing).Value2)));
            Column = "J16";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G38", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G50", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "I17";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2);
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2));
            Column = "J17";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "I18";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2)));
            Column = "J18";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "I19";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2);
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2));
            Column = "J19";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "I20";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2)));
            Column = "J20";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "I21";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = SumSubCostTotal.ToString();
            Column = "J21";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(SumSubCostTotal.ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "I22";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (SumSubBenefitTotal - SumSubCostTotal).ToString();
            Column = "J22";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((SumSubBenefitTotal - SumSubCostTotal).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));


            Int64 SumSubBenefitWillTotal = 0;

            Column = "O7";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G26", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F7", Type.Missing).Value2))) * 18;
            Column = "N7";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G26", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F7", Type.Missing).Value2))) * 18 * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G26", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F7", Type.Missing).Value2))) * 18 * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "O8";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G27", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F7", Type.Missing).Value2))) * 18;
            Column = "N8";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G27", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F7", Type.Missing).Value2))) * 18 * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G27", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F7", Type.Missing).Value2))) * 18 * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "N9";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G44", Type.Missing).Value2));
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G44", Type.Missing).Value2));
            Column = "O9";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G44", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "N10";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G29", Type.Missing).Value2);
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G29", Type.Missing).Value2));
            Column = "O10";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G29", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "N11";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G30", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G31", Type.Missing).Value2))).ToString();
            SumSubBenefitWillTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G30", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G31", Type.Missing).Value2)));
            Column = "O11";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G30", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G31", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "N12";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G45", Type.Missing).Value2);
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G45", Type.Missing).Value2));
            Column = "O12";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G45", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "N13";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = SumSubBenefitWillTotal.ToString();
            Column = "O13";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(SumSubBenefitWillTotal.ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));


            Int64 SumSubCostWillTotal = 0;
            Column = "N14";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G32", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F8", Type.Missing).Value2)) + (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G33", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F9", Type.Missing).Value2))));
            SumSubCostWillTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G32", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F8", Type.Missing).Value2)) + (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G33", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F9", Type.Missing).Value2))));
            Column = "O14";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G26", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));

            Column = "N15";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G36", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G37", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G48", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G49", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2)));
            SumSubCostWillTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G36", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G37", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G48", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G49", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2)));
            Column = "O15";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G36", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G37", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G48", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G49", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "N16";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G38", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G50", Type.Missing).Value2)));
            SumSubCostWillTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G38", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G50", Type.Missing).Value2)));
            Column = "O16";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G38", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G50", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "N17";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2);
            SumSubCostWillTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2));
            Column = "O17";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "N18";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2)));
            SumSubCostWillTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2)));
            Column = "O18";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "N19";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2);
            SumSubCostWillTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2));
            Column = "O19";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "N20";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2)));
            SumSubCostWillTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2)));
            Column = "O20";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "N21";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = SumSubCostWillTotal.ToString();
            Column = "O21";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(SumSubCostWillTotal.ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "N22";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (SumSubBenefitWillTotal - SumSubCostWillTotal).ToString();
            Column = "O22";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((SumSubBenefitWillTotal - SumSubCostWillTotal).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));






            //도매
            //당대리점 결과(세부항목별 값 입력 결과) 수익계정
            SumSubBenefitTotal = 0;
            Column = "I28";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G26", Type.Missing).Value2);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G26", Type.Missing).Value2));
            Column = "J28";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G26", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "I29";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G27", Type.Missing).Value2);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G27", Type.Missing).Value2));
            Column = "J29";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G27", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "I30";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G29", Type.Missing).Value2);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G29", Type.Missing).Value2));
            Column = "J30";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G29", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "I31";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G30", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G31", Type.Missing).Value2))).ToString();
            SumSubBenefitTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G30", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G31", Type.Missing).Value2)));
            Column = "J31";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G30", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G31", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "I32";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = SumSubBenefitTotal.ToString();
            Column = "J32";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(SumSubBenefitTotal.ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));

            //당대리점 결과(세부항목별 값 입력 결과) 비용계정
            SumSubCostTotal = 0;
            Column = "I33";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G32", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F8", Type.Missing).Value2)) + (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G33", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F9", Type.Missing).Value2))));
            SumSubCostTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G32", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F8", Type.Missing).Value2)) + (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G33", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F9", Type.Missing).Value2))));
            Column = "J33";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G26", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "I34";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G36", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G37", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F20", Type.Missing).Value2)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G36", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G37", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F20", Type.Missing).Value2)));
            Column = "J34";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G36", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G37", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F20", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "I35";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G38", Type.Missing).Value2)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G38", Type.Missing).Value2)));
            Column = "J35";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G38", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "I36";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "J36";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));

            Column = "I37";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "J37";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "I38";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "J38";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "I39";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "J39";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "I40";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = SumSubCostTotal.ToString();
            Column = "J40";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(SumSubCostTotal.ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "I41";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (SumSubBenefitTotal - SumSubCostTotal).ToString();
            Column = "J41";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((SumSubBenefitTotal - SumSubCostTotal).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));


            SumSubBenefitWillTotal = 0;
            Column = "N28";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G26", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F7", Type.Missing).Value2))) * 18 * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G26", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F7", Type.Missing).Value2))) * 18 * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "O28";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G26", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F7", Type.Missing).Value2))) * 18 * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "N29";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G27", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F7", Type.Missing).Value2))) * 18 * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G27", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F7", Type.Missing).Value2))) * 18 * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2));
            Column = "O29";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G27", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F7", Type.Missing).Value2))) * 18 * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "N30";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G29", Type.Missing).Value2);
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G29", Type.Missing).Value2));
            Column = "O30";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G29", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "N31";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G30", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G31", Type.Missing).Value2))).ToString();
            SumSubBenefitWillTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G30", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G31", Type.Missing).Value2)));
            Column = "O31";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G30", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G31", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "N32";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = SumSubBenefitWillTotal.ToString();
            Column = "O32";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(SumSubBenefitWillTotal.ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));

            SumSubCostTotal = 0;
            Column = "N33";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G32", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F8", Type.Missing).Value2)) + (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G33", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F9", Type.Missing).Value2))));
            SumSubCostTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G32", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F8", Type.Missing).Value2)) + (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G33", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F9", Type.Missing).Value2))));
            Column = "O33";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G26", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "N34";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G36", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G37", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F20", Type.Missing).Value2)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G36", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G37", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F20", Type.Missing).Value2)));
            Column = "O34";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G36", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G37", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F20", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "N35";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G38", Type.Missing).Value2)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G38", Type.Missing).Value2)));
            Column = "O35";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G38", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "N36";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "O36";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));

            Column = "N37";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "O37";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "N38";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "O38";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "N39";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "O39";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "N40";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = SumSubCostTotal.ToString();
            Column = "O40";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(SumSubCostTotal.ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "N41";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (SumSubBenefitTotal - SumSubCostTotal).ToString();
            Column = "O41";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((SumSubBenefitTotal - SumSubCostTotal).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));


            //소매 당대리점
            SumSubBenefitTotal = 0;
            Column = "I46";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G44", Type.Missing).Value2);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G44", Type.Missing).Value2));
            Column = "J46";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G44", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "I47";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G45", Type.Missing).Value2);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G45", Type.Missing).Value2));
            Column = "J47";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G45", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "I48";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = SumSubBenefitTotal.ToString();
            Column = "J48";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(SumSubBenefitTotal.ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));

            //당대리점 결과(세부항목별 값 입력 결과) 비용계정
            SumSubCostTotal = 0;
            Column = "I49";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G48", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G49", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G20", Type.Missing).Value2)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G48", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G49", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G20", Type.Missing).Value2)));
            Column = "J49";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G48", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G49", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G20", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "I50";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G50", Type.Missing).Value2)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G50", Type.Missing).Value2)));
            Column = "J50";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G50", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "I51";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "J51";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));

            Column = "I52";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "J52";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "I53";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "J53";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "I54";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "J54";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "I55";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = SumSubCostTotal.ToString();
            Column = "J55";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(SumSubCostTotal.ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "I56";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (SumSubBenefitTotal - SumSubCostTotal).ToString();
            Column = "J56";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((SumSubBenefitTotal - SumSubCostTotal).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));

            Column = "I57";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((SumSubBenefitTotal - SumSubCostTotal).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H15", Type.Missing).Value2));
            Column = "J57";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("H15", Type.Missing).Value2);


            //소매 당대리점미래
            SumSubBenefitWillTotal = 0;
            Column = "N46";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G44", Type.Missing).Value2);
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G44", Type.Missing).Value2));
            Column = "O46";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G44", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "N47";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("G45", Type.Missing).Value2);
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G45", Type.Missing).Value2));
            Column = "O47";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(CommonUtil.NullToString0(_WorkSheet1.get_Range("G45", Type.Missing).Value2), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "N48";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = SumSubBenefitWillTotal.ToString();
            Column = "O48";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(SumSubBenefitWillTotal.ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));

            //당대리점 결과(세부항목별 값 입력 결과) 비용계정
            SumSubCostWillTotal = 0;
            Column = "N49";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G48", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G49", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G20", Type.Missing).Value2)));
            SumSubCostWillTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G48", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G49", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G20", Type.Missing).Value2)));
            Column = "O49";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G48", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G18", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G49", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G19", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G20", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "N50";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G50", Type.Missing).Value2)));
            SumSubCostWillTotal += (CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G50", Type.Missing).Value2)));
            Column = "O50";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G50", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "N51";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            SumSubCostWillTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "O51";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G58", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));

            Column = "N52";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            SumSubCostWillTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "O52";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G59", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "N53";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            SumSubCostWillTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "O53";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G60", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "N54";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            SumSubCostWillTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "O54";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G39", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G40", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G41", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G42", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G43", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G51", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G52", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G53", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G54", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G55", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G56", Type.Missing).Value2)) + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G57", Type.Missing).Value2))
                + CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G61", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H10", Type.Missing).Value2))) * CommonUtil.StringToIntVal(CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2))).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("G10", Type.Missing).Value2));
            Column = "N55";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = SumSubCostWillTotal.ToString();
            Column = "O55";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division(SumSubCostWillTotal.ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));
            Column = "N56";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = (SumSubBenefitWillTotal - SumSubCostWillTotal).ToString();
            Column = "O56";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((SumSubBenefitWillTotal - SumSubCostWillTotal).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("F10", Type.Missing).Value2));

            Column = "N57";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.Division((SumSubBenefitWillTotal - SumSubCostWillTotal).ToString(), CommonUtil.NullToString0(_WorkSheet1.get_Range("H15", Type.Missing).Value2));
            Column = "O57";
            _WorkSheet2.get_Range(Column, Type.Missing).Value2 = CommonUtil.NullToString0(_WorkSheet1.get_Range("H15", Type.Missing).Value2);


        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel File|*.xlsx";
            saveFileDialog1.Title = "Select a Excel File";
            saveFileDialog1.ShowDialog();

            // If the file name is not an empty string open it for saving.
            if (saveFileDialog1.FileName != "")
            {
                FileInfo fi2 = new FileInfo(CommonUtil.defaultName);
                fi2.CopyTo(saveFileDialog1.FileName, true);
                CommonUtil.saveAsSimulName = saveFileDialog1.FileName;
                excel.Workbook _Workbook = CommonUtil.GetExcel_WorkBookForSimul(saveFileDialog1.FileName);
                excel.Worksheet _WorkSheet1 = _Workbook.Sheets[1] as excel.Worksheet;
                excel.Worksheet _WorkSheet2 = _Workbook.Sheets[2] as excel.Worksheet;
                SaveAsInput(_WorkSheet1, _WorkSheet2);
                CommonUtil.GetExcel_WorkBook_CLOSE();
            }
            SaveAsInput();
            mFormUserSimulOutput.applyData();
            this.Close();
        }

        private void txtDetail37_TextChanged(object sender, EventArgs e) {
            setTxtInput_TextChanged(sender);

            txtDetail42.Text = (CommonUtil.StringToIntVal(txtDetail37.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail38.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail39.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail40.Text.Replace(",", ""))
                 + CommonUtil.StringToIntVal(txtDetail41.Text.Replace(",", ""))).ToString();
        }
        private void txtDetail38_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
            txtDetail42.Text = (CommonUtil.StringToIntVal(txtDetail37.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail38.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail39.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail40.Text.Replace(",", ""))
                 + CommonUtil.StringToIntVal(txtDetail41.Text.Replace(",", ""))).ToString();
        }
        private void txtDetail39_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail42.Text = (CommonUtil.StringToIntVal(txtDetail37.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail38.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail39.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail40.Text.Replace(",", ""))
                 + CommonUtil.StringToIntVal(txtDetail41.Text.Replace(",", ""))).ToString();
        }
        private void txtDetail40_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
            
            txtDetail42.Text = (CommonUtil.StringToIntVal(txtDetail37.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail38.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail39.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail40.Text.Replace(",", ""))
                 + CommonUtil.StringToIntVal(txtDetail41.Text.Replace(",", ""))).ToString();
        }
        private void txtDetail41_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
            
            txtDetail42.Text = (CommonUtil.StringToIntVal(txtDetail37.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail38.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail39.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail40.Text.Replace(",", ""))
                 + CommonUtil.StringToIntVal(txtDetail41.Text.Replace(",", ""))).ToString();
        }

        private void txtDetail43_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
            
            txtDetail53.Text = (CommonUtil.StringToIntVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail44_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail53.Text = (CommonUtil.StringToIntVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail45_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail53.Text = (CommonUtil.StringToIntVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail46_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail53.Text = (CommonUtil.StringToIntVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail47_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail53.Text = (CommonUtil.StringToIntVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail48_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail53.Text = (CommonUtil.StringToIntVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail49_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail53.Text = (CommonUtil.StringToIntVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail50_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail53.Text = (CommonUtil.StringToIntVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail51_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail53.Text = (CommonUtil.StringToIntVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail52_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail53.Text = (CommonUtil.StringToIntVal(txtDetail43.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail44.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail45.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail46.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail47.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail48.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail49.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail50.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail51.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail52.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail54_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail56.Text = (CommonUtil.StringToIntVal(txtDetail54.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail55.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail55_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail56.Text = (CommonUtil.StringToIntVal(txtDetail54.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail55.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail57_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail63.Text = (CommonUtil.StringToIntVal(txtDetail57.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail58.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail59.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail60.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail61.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail62.Text.Replace(",", ""))
                 ).ToString();
        }
        private void txtDetail58_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail63.Text = (CommonUtil.StringToIntVal(txtDetail57.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail58.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail59.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail60.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail61.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail62.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail59_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail63.Text = (CommonUtil.StringToIntVal(txtDetail57.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail58.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail59.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail60.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail61.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail62.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail60_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail63.Text = (CommonUtil.StringToIntVal(txtDetail57.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail58.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail59.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail60.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail61.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail62.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail61_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail63.Text = (CommonUtil.StringToIntVal(txtDetail57.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail58.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail59.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail60.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail61.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail62.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail62_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail63.Text = (CommonUtil.StringToIntVal(txtDetail57.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail58.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail59.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail60.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail61.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail62.Text.Replace(",", ""))
                 ).ToString();
        }


        private void txtDetail64_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail72.Text = (CommonUtil.StringToIntVal(txtDetail64.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail65.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail66.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail67.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail68.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail69.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail70.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail71.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail65_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail72.Text = (CommonUtil.StringToIntVal(txtDetail64.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail65.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail66.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail67.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail68.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail69.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail70.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail71.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail66_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail72.Text = (CommonUtil.StringToIntVal(txtDetail64.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail65.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail66.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail67.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail68.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail69.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail70.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail71.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail67_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail72.Text = (CommonUtil.StringToIntVal(txtDetail64.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail65.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail66.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail67.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail68.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail69.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail70.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail71.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail68_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail72.Text = (CommonUtil.StringToIntVal(txtDetail64.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail65.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail66.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail67.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail68.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail69.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail70.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail71.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail69_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail72.Text = (CommonUtil.StringToIntVal(txtDetail64.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail65.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail66.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail67.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail68.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail69.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail70.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail71.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail70_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail72.Text = (CommonUtil.StringToIntVal(txtDetail64.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail65.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail66.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail67.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail68.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail69.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail70.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail71.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail71_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtDetail72.Text = (CommonUtil.StringToIntVal(txtDetail64.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail65.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail66.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail67.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail68.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail69.Text.Replace(",", ""))
                + CommonUtil.StringToIntVal(txtDetail70.Text.Replace(",", "")) + CommonUtil.StringToIntVal(txtDetail71.Text.Replace(",", ""))
                 ).ToString();
        }

        private void txtDetail42_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);


        }

        private void txtDetail53_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);


        }

        private void txtDetail56_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);


        }

        private void txtDetail63_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);


        }

        private void txtDetail72_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);


        }

        //시뮬레이션 기본입력
        private void txtInput10_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);


        }

        private void txtInput11_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

        }

        private void txtInput12_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

        }

        private void txtInput13_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);


        }

        private void txtInput14_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

        }

        private void txtInput15_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

        }

        private void txtInput16_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

        }

        private void txtInput17_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

        }

        private void txtInput18_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

        }

        private void txtInput24_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

        }

        private void txtInput25_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

        }

        private void txtInput26_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

        }

        private void txtInput27_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

        }

        private void txtInput28_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

        }





        private string setTxtInput_TextChanged(object sender)
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
    }
}
