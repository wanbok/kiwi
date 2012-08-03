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
    public partial class FormUserInput : Form
    {


        private TextBox[] txtBasicInput = null;     //기본입력
        private TextBox[] txtDetailInput = null;    //상세입력

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

        private string[] txtMangerInput = null;

        public FormUserInput()
        {
            InitializeComponent();
            txtMangerInput = new string[31] { txtMangeInput1, txtMangeInput2, txtMangeInput3, txtMangeInput4, txtMangeInput5, 
                txtMangeInput6, txtMangeInput7, txtMangeInput8, txtMangeInput9, txtMangeInput10,
                txtMangeInput11, txtMangeInput12, txtMangeInput13, txtMangeInput14, txtMangeInput15,
                txtMangeInput16, txtMangeInput17, txtMangeInput18, txtMangeInput19, txtMangeInput20,
                txtMangeInput21, txtMangeInput22, txtMangeInput23, txtMangeInput24, txtMangeInput25, 
                txtMangeInput26, txtMangeInput27, txtMangeInput28, txtMangeInput29, txtMangeInput30,
                txtMangeInput31

            };
            //더블 버퍼
            this.SetStyle(ControlStyles.DoubleBuffer, true);
            this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            this.SetStyle(ControlStyles.UserPaint, true);
            this.UpdateStyles();

            txtBasicInput = new TextBox[35] { txtInput1, txtInput2, txtInput3, txtInput4, txtInput5, txtInput6, txtInput7, txtInput8, txtInput9, txtInput10,
            txtInput11, txtInput12, txtInput13, txtInput14, txtInput15, txtInput16, txtInput17, txtInput18, txtInput19, txtInput20,
            txtInput21, txtInput22, txtInput23, txtInput24, txtInput25, txtInput26, txtInput27, txtInput28, txtInput29, txtInput30,
            txtInput31, txtInput32, txtInput33, txtInput34, txtInput35
            };

            txtDetailInput = new TextBox[36] { txtDetail1, txtDetail2, txtDetail3, txtDetail4, txtDetail5, txtDetail6, txtDetail7, txtDetail8, txtDetail9, txtDetail10,
            txtDetail11, txtDetail12, txtDetail13, txtDetail14, txtDetail15, txtDetail16, txtDetail17, txtDetail18, txtDetail19, txtDetail20,            
            txtDetail21, txtDetail22, txtDetail23, txtDetail24, txtDetail25, txtDetail26, txtDetail27, txtDetail28, txtDetail29, txtDetail30,            
            txtDetail31, txtDetail32, txtDetail33, txtDetail34, txtDetail35, txtDetail36
            };

            this.txtInput1.TextChanged += new System.EventHandler(this.txtInput1_TextChanged);
            this.txtInput2.TextChanged += new System.EventHandler(this.txtInput2_TextChanged);
            this.txtInput3.TextChanged += new System.EventHandler(this.txtInput3_TextChanged);
            this.txtInput4.TextChanged += new System.EventHandler(this.txtInput4_TextChanged);
            this.txtInput5.TextChanged += new System.EventHandler(this.txtInput5_TextChanged);
            this.txtInput6.TextChanged += new System.EventHandler(this.txtInput6_TextChanged);
            this.txtInput7.TextChanged += new System.EventHandler(this.txtInput7_TextChanged);
            this.txtInput8.TextChanged += new System.EventHandler(this.txtInput8_TextChanged);
            this.txtInput9.TextChanged += new System.EventHandler(this.txtInput9_TextChanged);
            this.txtInput10.TextChanged += new System.EventHandler(this.txtInput10_TextChanged);
            this.txtInput11.TextChanged += new System.EventHandler(this.txtInput11_TextChanged);
            this.txtInput12.TextChanged += new System.EventHandler(this.txtInput12_TextChanged);
            this.txtInput13.TextChanged += new System.EventHandler(this.txtInput13_TextChanged);
            this.txtInput14.TextChanged += new System.EventHandler(this.txtInput14_TextChanged);
            this.txtInput15.TextChanged += new System.EventHandler(this.txtInput15_TextChanged);
            this.txtInput16.TextChanged += new System.EventHandler(this.txtInput16_TextChanged);
            this.txtInput17.TextChanged += new System.EventHandler(this.txtInput17_TextChanged);
            this.txtInput18.TextChanged += new System.EventHandler(this.txtInput18_TextChanged);
            this.txtInput19.TextChanged += new System.EventHandler(this.txtInput19_TextChanged);
            this.txtInput20.TextChanged += new System.EventHandler(this.txtInput20_TextChanged);
            this.txtInput21.TextChanged += new System.EventHandler(this.txtInput21_TextChanged);
            this.txtInput4.ReadOnly = true;
            this.txtInput7.ReadOnly = true;
            this.txtInput10.ReadOnly = true;
            this.txtInput13.ReadOnly = true;
            this.txtInput16.ReadOnly = true;
            this.txtInput18.ReadOnly = true;
            this.txtInput21.ReadOnly = true;
            this.txtInput22.ReadOnly = true;
            this.txtInput23.ReadOnly = true;
            this.txtInput24.ReadOnly = true;
            this.txtInput25.ReadOnly = true;
            this.txtInput26.ReadOnly = true;
            this.txtInput27.ReadOnly = true;
            this.txtInput28.ReadOnly = true;
            this.txtInput29.ReadOnly = true;
            this.txtInput30.ReadOnly = true;
            this.txtInput31.ReadOnly = true;
            this.txtInput32.ReadOnly = true;
            this.txtInput33.ReadOnly = true;
            this.txtInput34.ReadOnly = true;
            this.txtInput35.ReadOnly = true;


            this.radioButton5.CheckedChanged += new System.EventHandler(this.radioButton5_CheckedChanged);
            this.radioButton6.CheckedChanged += new System.EventHandler(this.radioButton6_CheckedChanged);
            this.radioButton1.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            this.radioButton2.CheckedChanged += new System.EventHandler(this.radioButton2_CheckedChanged);
            this.radioButton3.CheckedChanged += new System.EventHandler(this.radioButton3_CheckedChanged);
            this.radioButton4.CheckedChanged += new System.EventHandler(this.radioButton4_CheckedChanged);

            CommonUtil.clearTextBox(this.tabPage1);
            CommonUtil.clearTextBox(this.tabPage5);
            
            getInput();
            getDetail(CDataControl.g_BasicInput);

            radioButton5.Checked = true;
            radioButton2.Checked = true;
            radioButton3.Checked = true;
        }

        //기본입력
        private void getInput()
        {
            comboBox1.SelectedItem = CDataControl.g_BasicInput.get지역();
            textBox6.Text = CDataControl.g_BasicInput.get대리점();
            textBox9.Text = CDataControl.g_BasicInput.get마케터();
            Int64[] arrvalue = CDataControl.g_BasicInput.getArrData_BasicInput();
            // 셀에서 데이터 가져오기
            for (int i = 0; i < txtBasicInput.Length; i++)
            {
                txtBasicInput[i].Text = arrvalue[i].ToString();
            }
        }

        //상세입력
        private void getDetail(CBasicInput g_BasicInput)
        {
            //
            Int64[] arrvalue = CDataControl.g_DetailInput.getArrData_DetailInput(g_BasicInput.get도매_직원수_간부급(), g_BasicInput.get도매_직원수_평사원()
                , g_BasicInput.get소매_직원수_간부급(), g_BasicInput.get소매_직원수_평사원());

            // 셀에서 데이터 가져오기
            for (int i = 0; i < txtDetailInput.Length; i++)
            {
                txtDetailInput[i].Text = arrvalue[i].ToString();
            }
        }

        public void saveAsInput()
        {
            CBasicInput bi = CDataControl.g_BasicInput;
            CBusinessData di = CDataControl.g_DetailInput;
            CResultData[] rdts = new CResultData[]{CDataControl.g_ResultStoreTotal, CDataControl.g_ResultFutureTotal};
            CResultData[] rds = new CResultData[]{CDataControl.g_ResultStore, CDataControl.g_ResultFuture};
            CResultData rdt = null;
            CResultData rd = null;

            bi.set지역(comboBox1.SelectedIndex == -1 ? "" : comboBox1.Items[comboBox1.SelectedIndex].ToString());
            bi.set대리점(textBox6.Text);
            bi.set마케터(textBox9.Text);

            String[] txtWrite = new String[14] { txtInput1.Text, txtInput2.Text, txtInput3.Text, txtInput5.Text, txtInput6.Text,  
                txtInput8.Text, txtInput9.Text, txtInput11.Text, txtInput12.Text, txtInput14.Text, txtInput15.Text, txtInput17.Text, txtInput19.Text, txtInput20.Text};
            bi.setArrData_BasicInput(txtWrite);

            String[] txtWrite2 = new String[31]  {
                txtDetail1.Text, txtDetail2.Text, txtDetail4.Text, txtDetail5.Text, txtDetail6.Text, // 도매 수익
                txtDetail7.Text, txtDetail8.Text, txtDetail11.Text, txtDetail12.Text, txtDetail13.Text, txtDetail14.Text, txtDetail15.Text, txtDetail16.Text, txtDetail17.Text, txtDetail18.Text, // 도매 비용
                txtDetail19.Text, txtDetail20.Text, txtDetail23.Text, txtDetail24.Text, txtDetail25.Text, txtDetail26.Text, txtDetail27.Text, txtDetail28.Text, // 소매
                txtDetail29.Text, txtDetail30.Text, txtDetail31.Text, txtDetail32.Text, txtDetail33.Text, txtDetail34.Text, txtDetail35.Text, txtDetail36.Text  // 도소매합산
            };
            di.setArrData_DetailInput(txtWrite2);
            CommonUtil.ReadFileManagerToData();

            
            for (int i = 0; i < rdts.Length; i++) {
                //  당대리점 결과(현재:0, 미래:1)
                rdt = rdts[i];
                rd = rds[i];
                //      도매
                //          총액
                //              수익
                rdt.set도매_수익_가입자관리수수료(i == 0 ? di.get도매_수익_월평균관리수수료() : CommonUtil.Division(di.get도매_수익_월평균관리수수료(),bi.get도매_누적가입자수())*18*bi.get월평균판매대수_소계_합계());
                rdt.set도매_수익_CS관리수수료(di.get도매_수익_CS관리수수료());
                rdt.set도매_수익_사업자모델매입에따른추가수익(di.get도매_수익_사업자모델매입관련추가수익());
                rdt.set도매_수익_유통모델매입에따른추가수익_현금_Volume(di.get도매_수익_유통모델매입관련추가수익_현금DC() + di.get도매_수익_유통모델매입관련추가수익_VolumeDC());
                rdt.도매_수익_소계 = rdt.get도매_수익_가입자관리수수료() + rdt.get도매_수익_CS관리수수료() + rdt.get도매_수익_사업자모델매입에따른추가수익() + rdt.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
                //              비용
                rdt.set도매_비용_대리점투자비용(di.get도매_비용_대리점투자금액_신규() * bi.get도매_월평균판매대수_신규() + di.get도매_비용_대리점투자금액_기변() * bi.get도매_월평균판매대수_기변());
                rdt.set도매_비용_인건비_급여_복리후생비(di.get도매_비용_직원급여_간부급() * bi.get도매_직원수_간부급() + di.get도매_비용_직원급여_평사원() * bi.get도매_직원수_평사원() + di.get도소매_비용_복리후생비());
                rdt.set도매_비용_임차료(di.get도매_비용_지급임차료());
                rdt.set도매_비용_이자비용(CommonUtil.Division(di.get도소매_비용_이자비용(),bi.get월평균판매대수_소계_합계())*bi.get도매_월평균판매대수_소계());
                rdt.set도매_비용_부가세(CommonUtil.Division(di.get도소매_비용_부가세(),bi.get월평균판매대수_소계_합계())*bi.get도매_월평균판매대수_소계());
                rdt.set도매_비용_법인세(CommonUtil.Division(di.get도소매_비용_법인세(),bi.get월평균판매대수_소계_합계())*bi.get도매_월평균판매대수_소계());
                rdt.set도매_비용_기타판매관리비(di.get도매_비용_운반비()+di.get도매_비용_차량유지비()+di.get도매_비용_지급수수료()+di.get도매_비용_판매촉진비()+di.get도매_비용_건물관리비()+(CommonUtil.Division((di.get도소매_비용_복리후생비()+di.get도소매_비용_통신비()+di.get도소매_비용_공과금()+di.get도소매_비용_소모품비()+di.get도소매_비용_기타()),bi.get월평균판매대수_소계_합계())*bi.get도매_월평균판매대수_소계()));
                rdt.도매_비용_소계 = rdt.get도매_비용_대리점투자비용() + rdt.get도매_비용_인건비_급여_복리후생비() + rdt.get도매_비용_임차료() + rdt.get도매_비용_이자비용() + rdt.get도매_비용_부가세() + rdt.get도매_비용_법인세() + rdt.get도매_비용_기타판매관리비();
                rdt.도매손익계 = rdt.도매_수익_소계 - rdt.도매_비용_소계;
                //          단위당 금액
                //              수익
                rd.set도매_수익_가입자관리수수료(CommonUtil.Division(rdt.get도매_수익_가입자관리수수료(),bi.get도매_월평균판매대수_소계()));
                rd.set도매_수익_CS관리수수료(CommonUtil.Division(rdt.get도매_수익_CS관리수수료(),bi.get도매_월평균판매대수_소계()));
                rd.set도매_수익_사업자모델매입에따른추가수익(CommonUtil.Division(rdt.get도매_수익_사업자모델매입에따른추가수익(),bi.get도매_월평균판매대수_소계()));
                rd.set도매_수익_유통모델매입에따른추가수익_현금_Volume(CommonUtil.Division(rdt.get도매_수익_유통모델매입에따른추가수익_현금_Volume(),bi.get도매_월평균판매대수_소계()));
                rd.도매_수익_소계 = CommonUtil.Division(rdt.도매_수익_소계,bi.get도매_월평균판매대수_소계());
                //              비용
                rd.set도매_비용_대리점투자비용(CommonUtil.Division(rdt.get도매_비용_대리점투자비용(),bi.get도매_월평균판매대수_소계()));
                rd.set도매_비용_인건비_급여_복리후생비(CommonUtil.Division(rdt.get도매_비용_인건비_급여_복리후생비(),bi.get도매_월평균판매대수_소계()));
                rd.set도매_비용_임차료(CommonUtil.Division(rdt.get도매_비용_임차료(),bi.get도매_월평균판매대수_소계()));
                rd.set도매_비용_이자비용(CommonUtil.Division(rdt.get도매_비용_이자비용(),bi.get도매_월평균판매대수_소계()));
                rd.set도매_비용_부가세(CommonUtil.Division(rdt.get도매_비용_부가세(),bi.get도매_월평균판매대수_소계()));
                rd.set도매_비용_법인세(CommonUtil.Division(rdt.get도매_비용_법인세(),bi.get도매_월평균판매대수_소계()));
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
                rdt.set소매_비용_인건비_급여_복리후생비(di.get소매_비용_직원급여_간부급() * bi.get소매_직원수_간부급() + di.get소매_비용_직원급여_평사원() * bi.get소매_직원수_평사원() + di.get도소매_비용_복리후생비());
                rdt.set소매_비용_임차료(di.get소매_비용_지급임차료());
                rdt.set소매_비용_이자비용(CommonUtil.Division(di.get도소매_비용_이자비용(),bi.get월평균판매대수_소계_합계())*bi.get소매_월평균판매대수_소계());
                rdt.set소매_비용_부가세(CommonUtil.Division(di.get도소매_비용_부가세(),bi.get월평균판매대수_소계_합계())*bi.get소매_월평균판매대수_소계());
                rdt.set소매_비용_법인세(CommonUtil.Division(di.get도소매_비용_법인세(),bi.get월평균판매대수_소계_합계())*bi.get소매_월평균판매대수_소계());
                rdt.set소매_비용_기타판매관리비((di.get소매_비용_지급수수료()+di.get소매_비용_판매촉진비()+di.get소매_비용_건물관리비())+(CommonUtil.Division((di.get도소매_비용_복리후생비()+di.get도소매_비용_통신비()+di.get도소매_비용_공과금()+di.get도소매_비용_소모품비()+di.get도소매_비용_기타()),bi.get월평균판매대수_소계_합계())*bi.get소매_월평균판매대수_소계()));
                rdt.소매_비용_소계 = rdt.get소매_비용_인건비_급여_복리후생비() + rdt.get소매_비용_임차료() + rdt.get소매_비용_이자비용() + rdt.get소매_비용_부가세() + rdt.get소매_비용_법인세() + rdt.get소매_비용_기타판매관리비();
                rdt.소매손익계 = rdt.소매_수익_소계 - rdt.소매_비용_소계;
                rdt.점별손익추정 = CommonUtil.Division(rdt.소매손익계, bi.get거래선수_직영점_합계());
                //          단위당 금액
                //              수익
                rd.set소매_수익_업무취급수수료(CommonUtil.Division(rdt.get소매_수익_업무취급수수료(), bi.get소매_월평균판매대수_소계()));
                rd.set소매_수익_직영매장판매수익(CommonUtil.Division(rdt.get소매_수익_직영매장판매수익(), bi.get소매_월평균판매대수_소계()));
                rd.소매_수익_소계 = CommonUtil.Division(rdt.소매_수익_소계, bi.get소매_월평균판매대수_소계());
                //              비용
                rd.set소매_비용_인건비_급여_복리후생비(CommonUtil.Division(rdt.get소매_비용_인건비_급여_복리후생비() , bi.get소매_월평균판매대수_소계()));
                rd.set소매_비용_임차료(CommonUtil.Division(rdt.get소매_비용_임차료() , bi.get소매_월평균판매대수_소계()));
                rd.set소매_비용_이자비용(CommonUtil.Division(rdt.get소매_비용_이자비용() , bi.get소매_월평균판매대수_소계()));
                rd.set소매_비용_부가세(CommonUtil.Division(rdt.get소매_비용_부가세() , bi.get소매_월평균판매대수_소계()));
                rd.set소매_비용_법인세(CommonUtil.Division(rdt.get소매_비용_법인세() , bi.get소매_월평균판매대수_소계()));
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
                rdt.set전체_비용_기타판매관리비(di.get도매_비용_운반비()+di.get도매_비용_차량유지비()+di.get도매_비용_지급수수료()+di.get도매_비용_판매촉진비()+di.get도매_비용_건물관리비()+di.get소매_비용_지급수수료()+di.get소매_비용_판매촉진비()+di.get소매_비용_건물관리비()+di.get도소매_비용_복리후생비()+di.get도소매_비용_통신비()+di.get도소매_비용_공과금()+di.get도소매_비용_소모품비()+di.get도소매_비용_기타());
                rdt.전체_비용_소계 = rdt.get전체_비용_대리점투자비용() + rdt.get전체_비용_인건비_급여_복리후생비() + rdt.get전체_비용_임차료() + rdt.get전체_비용_이자비용() + rdt.get전체_비용_부가세() + rdt.get전체_비용_법인세() + rdt.get전체_비용_기타판매관리비();
                rdt.전체손익계 = rdt.전체_수익_소계 - rdt.전체_비용_소계;
                //          단위당 금액
                //              수익
                rd.set전체_수익_가입자관리수수료(CommonUtil.Division(rdt.get전체_수익_가입자관리수수료(),bi.get월평균판매대수_소계_합계()));
                rd.set전체_수익_CS관리수수료(CommonUtil.Division(rdt.get전체_수익_CS관리수수료(),bi.get월평균판매대수_소계_합계()));
                rd.set전체_수익_업무취급수수료(CommonUtil.Division(rdt.get전체_수익_업무취급수수료(),bi.get월평균판매대수_소계_합계()));
                rd.set전체_수익_사업자모델매입에따른추가수익(CommonUtil.Division(rdt.get전체_수익_사업자모델매입에따른추가수익(),bi.get월평균판매대수_소계_합계()));
                rd.set전체_수익_유통모델매입에따른추가수익_현금_Volume(CommonUtil.Division(rdt.get전체_수익_유통모델매입에따른추가수익_현금_Volume(),bi.get월평균판매대수_소계_합계()));
                rd.set전체_수익_직영매장판매수익(CommonUtil.Division(rdt.get전체_수익_직영매장판매수익(),bi.get월평균판매대수_소계_합계()));
                rd.전체_수익_소계 = CommonUtil.Division(rdt.전체_수익_소계, bi.get월평균판매대수_소계_합계());
                //              비용
                rd.set전체_비용_대리점투자비용(CommonUtil.Division(rdt.get전체_비용_대리점투자비용(),bi.get월평균판매대수_소계_합계()));
                rd.set전체_비용_인건비_급여_복리후생비(CommonUtil.Division(rdt.get전체_비용_인건비_급여_복리후생비() , bi.get월평균판매대수_소계_합계()));
                rd.set전체_비용_임차료(CommonUtil.Division(rdt.get전체_비용_임차료() , bi.get월평균판매대수_소계_합계()));
                rd.set전체_비용_이자비용(CommonUtil.Division(rdt.get전체_비용_이자비용() , bi.get월평균판매대수_소계_합계()));
                rd.set전체_비용_부가세(CommonUtil.Division(rdt.get전체_비용_부가세() , bi.get월평균판매대수_소계_합계()));
                rd.set전체_비용_법인세(CommonUtil.Division(rdt.get전체_비용_법인세() , bi.get월평균판매대수_소계_합계()));
                rd.set전체_비용_기타판매관리비(CommonUtil.Division(rdt.get전체_비용_기타판매관리비(), bi.get월평균판매대수_소계_합계()));
                rd.전체_비용_소계 = CommonUtil.Division(rdt.전체_비용_소계, bi.get월평균판매대수_소계_합계());
                rd.전체손익계 = CommonUtil.Division(rdt.전체손익계, bi.get월평균판매대수_소계_합계());
            }
            
            //  업계 평균적용 결과
            rdt = CDataControl.g_ResultBusinessTotal;
            rd = CDataControl.g_ResultBusiness;
            di = CDataControl.g_BusinessAvg;     // 관리자가 배포한 업계 단위비용
            //      도매
            //          총액
            //              수익
            rdt.set도매_수익_가입자관리수수료(di.get도매_수익_월평균관리수수료() * bi.get도매_누적가입자수());
            rdt.set도매_수익_CS관리수수료(di.get도매_수익_CS관리수수료() * bi.get도매_누적가입자수());
            rdt.set도매_수익_사업자모델매입에따른추가수익(di.get도매_수익_사업자모델매입관련추가수익() * (bi.get월평균판매대수_소계_합계()-bi.get월평균유통모델출고대수_소계_합계()));
            rdt.set도매_수익_유통모델매입에따른추가수익_현금_Volume((di.get도매_수익_유통모델매입관련추가수익_현금DC() + di.get도매_수익_유통모델매입관련추가수익_VolumeDC()) * bi.get월평균유통모델출고대수_소계_합계());
            rdt.도매_수익_소계 = rdt.get도매_수익_가입자관리수수료() + rdt.get도매_수익_CS관리수수료() + rdt.get도매_수익_사업자모델매입에따른추가수익() + rdt.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
            //              비용
            rdt.set도매_비용_대리점투자비용(di.get도매_비용_대리점투자금액_신규() * bi.get도매_월평균판매대수_신규() + di.get도매_비용_대리점투자금액_기변() * bi.get도매_월평균판매대수_기변());
            rdt.set도매_비용_인건비_급여_복리후생비(di.get도매_비용_직원급여_간부급() * bi.get도매_직원수_간부급() + di.get도매_비용_직원급여_평사원() * bi.get도매_직원수_평사원() + di.get도소매_비용_복리후생비() * bi.get도매_직원수_소계());
            rdt.set도매_비용_임차료(di.get도매_비용_지급임차료() * bi.get도매_거래선수_개통사무실());
            rdt.set도매_비용_이자비용(di.get도소매_비용_이자비용() *bi.get도매_월평균판매대수_소계());
            rdt.set도매_비용_부가세(di.get도소매_비용_부가세() *bi.get도매_월평균판매대수_소계());
            rdt.set도매_비용_법인세(di.get도소매_비용_법인세() *bi.get도매_월평균판매대수_소계());
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
            rdt.set소매_수익_업무취급수수료(di.get소매_수익_월평균업무취급수수료()*bi.get월평균판매대수_소계_합계());
            rdt.set소매_수익_직영매장판매수익(di.get소매_수익_직영매장판매수익()*bi.get소매_월평균판매대수_소계());
            rdt.소매_수익_소계 = rdt.get소매_수익_업무취급수수료() + rdt.get소매_수익_직영매장판매수익();
            //              비용
            rdt.set소매_비용_인건비_급여_복리후생비(di.get소매_비용_직원급여_간부급() * bi.get소매_직원수_간부급() + di.get소매_비용_직원급여_평사원() * bi.get소매_직원수_평사원() + di.get도소매_비용_복리후생비() * bi.get소매_직원수_소계());
            rdt.set소매_비용_임차료(di.get소매_비용_지급임차료()*bi.get소매_거래선수_소계());
            rdt.set소매_비용_이자비용(di.get도소매_비용_이자비용()*bi.get소매_월평균판매대수_소계());
            rdt.set소매_비용_부가세(di.get도소매_비용_부가세()*bi.get소매_월평균판매대수_소계());
            rdt.set소매_비용_법인세(di.get도소매_비용_법인세()*bi.get소매_월평균판매대수_소계());
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
        


        //도매입력시 합계치 변경
        //누적가입자수
        private void txtInput1_TextChanged(object sender, EventArgs e)
        {
            txtInput22.Text = setTxtInput_TextChanged(sender);
        }
        //월평균 판매대수 신규
        private void txtInput2_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput4.Text = CommonUtil.Sum_Values(txtInput2.Text, txtInput3.Text);
            txtInput23.Text = CommonUtil.Sum_Values(txtInput2.Text, txtInput14.Text);
        }
        //월평균 판매대수 기변
        private void txtInput3_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput4.Text = CommonUtil.Sum_Values(txtInput2.Text, txtInput3.Text);
            txtInput24.Text = CommonUtil.Sum_Values(txtInput3.Text, txtInput15.Text);
        }
        //월평균 판매대수 계
        private void txtInput4_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput25.Text = CommonUtil.Sum_Values(txtInput4.Text, txtInput16.Text);
        }
        //월평균 유통모델 출고대수 LG
        private void txtInput5_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput7.Text = CommonUtil.Sum_Values(txtInput5.Text, txtInput6.Text);
            txtInput26.Text = txtInput5.Text;
        }
        //월평균 유통모델 출고대수 SS
        private void txtInput6_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput7.Text = CommonUtil.Sum_Values(txtInput5.Text, txtInput6.Text);
            txtInput27.Text = txtInput6.Text;
        }
        //월평균 유통모델 출고대수 계
        private void txtInput7_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput28.Text = txtInput7.Text;
        }

        //거래선 수 개통사무실
        private void txtInput8_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput10.Text = CommonUtil.Sum_Values(txtInput8.Text, txtInput9.Text);
            txtInput29.Text = txtInput8.Text;
        }
        //거래선 수 판매점
        private void txtInput9_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput10.Text = CommonUtil.Sum_Values(txtInput8.Text, txtInput9.Text);
            txtInput31.Text = txtInput9.Text;
        }
        //거래선 수 계
        private void txtInput10_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput32.Text = CommonUtil.Sum_Values(txtInput10.Text, txtInput18.Text);
        }

        //직원수 간부급
        private void txtInput11_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);


            txtInput13.Text = CommonUtil.Sum_Values(txtInput11.Text, txtInput12.Text);
            txtInput33.Text = CommonUtil.Sum_Values(txtInput11.Text, txtInput19.Text);
        }
        //직원수 평사원
        private void txtInput12_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput13.Text = CommonUtil.Sum_Values(txtInput11.Text, txtInput12.Text);
            txtInput34.Text = CommonUtil.Sum_Values(txtInput12.Text, txtInput20.Text);
        }
        //직원수 계
        private void txtInput13_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);


            txtInput35.Text = CommonUtil.Sum_Values(txtInput13.Text, txtInput21.Text);
        }

        //소매입력시 합계치 변경
        //월판매대수 신규
        private void txtInput14_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput16.Text = CommonUtil.Sum_Values(txtInput14.Text, txtInput15.Text);
            txtInput23.Text = CommonUtil.Sum_Values(txtInput2.Text, txtInput14.Text);
        }
        //월판매대수 신규
        private void txtInput15_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput16.Text = CommonUtil.Sum_Values(txtInput14.Text, txtInput15.Text);
            txtInput24.Text = CommonUtil.Sum_Values(txtInput3.Text, txtInput15.Text);
        }
        //월판매대수 신규
        private void txtInput16_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput25.Text = CommonUtil.Sum_Values(txtInput4.Text, txtInput16.Text);
        }

        //거래선 수 직영점
        private void txtInput17_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput18.Text = txtInput17.Text;
            txtInput30.Text = txtInput17.Text;
        }
        //거래선 수 계
        private void txtInput18_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput32.Text = CommonUtil.Sum_Values(txtInput10.Text, txtInput18.Text);
        }
        //직원수 간부급
        private void txtInput19_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput21.Text = CommonUtil.Sum_Values(txtInput19.Text, txtInput20.Text);
            txtInput33.Text = CommonUtil.Sum_Values(txtInput11.Text, txtInput19.Text);
        }
        //직원수 평사원
        private void txtInput20_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput21.Text = CommonUtil.Sum_Values(txtInput19.Text, txtInput20.Text);
            txtInput34.Text = CommonUtil.Sum_Values(txtInput12.Text, txtInput20.Text);
        }
        //직원수 계
        private void txtInput21_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            txtInput35.Text = CommonUtil.Sum_Values(txtInput13.Text, txtInput21.Text);
        }



        private void txtInput23_TextChanged(object sender, EventArgs e)
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

        private void txtInput32_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtInput33_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtInput34_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtInput35_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }




        //도매 수익 CS관리수수료 월총액
        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            foreach (Control SelectedButton in panel41.Controls)
            {
                if (SelectedButton is RadioButton)
                {
                    if (((RadioButton)SelectedButton).Checked && ((RadioButton)SelectedButton).Name == "radioButton5")
                    {
                        txtDetail2.ReadOnly = false;
                        txtDetail3.ReadOnly = true;

                        txtDetail2.BackColor = Color.White;
                        txtDetail3.BackColor = Color.Wheat;

                        txtDetail2.BorderStyle = BorderStyle.FixedSingle;
                        txtDetail3.BorderStyle = BorderStyle.None;

                        //txtDetail3.Text = (CommonUtil.StringToIntVal(txtDetail2.Text) * CommonUtil.QUARTER).ToString();
                    }
                }
            }  
        }
        //도매 수익 CS관리수수료 분기총액
        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            foreach (Control selectedbutton in panel41.Controls)
            {
                if (selectedbutton is RadioButton)
                {
                    if (((RadioButton)selectedbutton).Checked && ((RadioButton)selectedbutton).Name == "radioButton6")
                    {
                        txtDetail2.ReadOnly = true;
                        txtDetail3.ReadOnly = false;

                        txtDetail2.BackColor = Color.Wheat;
                        txtDetail3.BackColor = Color.White;

                        txtDetail2.BorderStyle = BorderStyle.None;
                        txtDetail3.BorderStyle = BorderStyle.FixedSingle;

                        //txtDetail2.Text = CommonUtil.Division(txtDetail3.Text, CommonUtil.QUARTER.ToString());
                    }
                }
            }
        }
        //도매 비용 직원급여 총액
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            foreach (Control SelectedButton in panel40.Controls)
            {
                if (SelectedButton is RadioButton)
                {
                    if (((RadioButton)SelectedButton).Checked && ((RadioButton)SelectedButton).Name == "radioButton1")
                    {
                        txtDetail9.ReadOnly = false;
                        txtDetail10.ReadOnly = false;
                        txtDetail11.ReadOnly = true;
                        txtDetail12.ReadOnly = true;

                        txtDetail9.BackColor = Color.White;
                        txtDetail10.BackColor = Color.White;
                        txtDetail11.BackColor = Color.Wheat;
                        txtDetail12.BackColor = Color.Wheat;

                        txtDetail9.BorderStyle = BorderStyle.FixedSingle;
                        txtDetail10.BorderStyle = BorderStyle.FixedSingle;
                        txtDetail11.BorderStyle = BorderStyle.None;
                        txtDetail12.BorderStyle = BorderStyle.None;

                        //txtDetail11.Text = CommonUtil.Division(txtDetail9.Text, txtInput11.Text);
                        //txtDetail12.Text = CommonUtil.Division(txtDetail10.Text, txtInput12.Text);
                    }
                }
            }
        }

        //도매 비용 직원급여 월평균
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            foreach (Control SelectedButton in panel40.Controls)
            {
                if (SelectedButton is RadioButton)
                {
                    if (((RadioButton)SelectedButton).Checked && ((RadioButton)SelectedButton).Name == "radioButton2")
                    {
                        txtDetail9.ReadOnly = true;
                        txtDetail10.ReadOnly = true;
                        txtDetail11.ReadOnly = false;
                        txtDetail12.ReadOnly = false;

                        txtDetail9.BackColor = Color.Wheat;
                        txtDetail10.BackColor = Color.Wheat;
                        txtDetail11.BackColor = Color.White;
                        txtDetail12.BackColor = Color.White;

                        txtDetail9.BorderStyle = BorderStyle.None;
                        txtDetail10.BorderStyle = BorderStyle.None;
                        txtDetail11.BorderStyle = BorderStyle.FixedSingle;
                        txtDetail12.BorderStyle = BorderStyle.FixedSingle;

                        //txtDetail9.Text = (CommonUtil.StringToIntVal(txtDetail9.Text) * CommonUtil.StringToIntVal(txtInput11.Text)).ToString();
                        //txtDetail10.Text = (CommonUtil.StringToIntVal(txtDetail10.Text) * CommonUtil.StringToIntVal(txtInput12.Text)).ToString();
                    }
                }
            }
        }

        //소매 비용 직원급여 총액
        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            foreach (Control SelectedButton in panel42.Controls)
            {
                if (SelectedButton is RadioButton)
                {
                    if (((RadioButton)SelectedButton).Checked && ((RadioButton)SelectedButton).Name == "radioButton4")
                    {
                        txtDetail21.ReadOnly = false;
                        txtDetail22.ReadOnly = false;
                        txtDetail23.ReadOnly = true;
                        txtDetail24.ReadOnly = true;

                        txtDetail21.BackColor = Color.White;
                        txtDetail22.BackColor = Color.White;
                        txtDetail23.BackColor = Color.Wheat;
                        txtDetail24.BackColor = Color.Wheat;

                        txtDetail21.BorderStyle = BorderStyle.FixedSingle;
                        txtDetail22.BorderStyle = BorderStyle.FixedSingle;
                        txtDetail23.BorderStyle = BorderStyle.None;
                        txtDetail24.BorderStyle = BorderStyle.None;

                        //txtDetail23.Text = CommonUtil.Division(txtDetail21.Text, txtInput19.Text);
                        //txtDetail24.Text = CommonUtil.Division(txtDetail22.Text, txtInput20.Text);

                    }
                }
            }
        }

        //소매 비용 직원급여 월평균
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            foreach (Control SelectedButton in panel42.Controls)
            {
                if (SelectedButton is RadioButton)
                {
                    if (((RadioButton)SelectedButton).Checked && ((RadioButton)SelectedButton).Name == "radioButton3")
                    {
                        txtDetail21.ReadOnly = true;
                        txtDetail22.ReadOnly = true;
                        txtDetail23.ReadOnly = false;
                        txtDetail24.ReadOnly = false;

                        txtDetail21.BackColor = Color.Wheat;
                        txtDetail22.BackColor = Color.Wheat;
                        txtDetail23.BackColor = Color.White;
                        txtDetail24.BackColor = Color.White;

                        txtDetail21.BorderStyle = BorderStyle.None;
                        txtDetail22.BorderStyle = BorderStyle.None;
                        txtDetail23.BorderStyle = BorderStyle.FixedSingle;
                        txtDetail24.BorderStyle = BorderStyle.FixedSingle;

                        //txtDetail21.Text = (CommonUtil.StringToIntVal(txtDetail23.Text) * CommonUtil.StringToIntVal(txtInput19.Text)).ToString();
                        //txtDetail22.Text = (CommonUtil.StringToIntVal(txtDetail24.Text) * CommonUtil.StringToIntVal(txtInput20.Text)).ToString();
                    }
                }
            }
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
                string filename = CommonUtil.saveAsName;
                if (filename == null)
                    filename = CommonUtil.defaultName;

                FileInfo fi2 = new FileInfo(filename);
                    fi2.CopyTo(saveFileDialog1.FileName, true);

                CommonUtil.saveAsName = saveFileDialog1.FileName;

                //excel.Workbook _Workbook = CommonUtil.GetExcel_WorkBook(saveFileDialog1.FileName);
                //excel.Worksheet _WorkSheet1 = _Workbook.Sheets[1] as excel.Worksheet;
                //excel.Worksheet _WorkSheet2 = _Workbook.Sheets[2] as excel.Worksheet;
                saveAsInput();
                CommonUtil.WriteDataToExcelFile(CommonUtil.saveAsName, CDataControl.g_BasicInput, CDataControl.g_DetailInput);
            }
        }



        //CS 관리 수수료 처리용 시작
        private void txtDetail2_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            if (radioButton5.Checked)
            {
                txtDetail3.Text = (CommonUtil.StringToIntVal(txtDetail2.Text.Replace(",","")) * CommonUtil.QUARTER).ToString();
            }
        }

        private void txtDetail3_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            if (radioButton6.Checked)
            {
                txtDetail2.Text = CommonUtil.Division(txtDetail3.Text.Replace(",", ""), CommonUtil.QUARTER.ToString());
            }
        }

        private void txtInput1_KeyPress(object sender, KeyPressEventArgs e)
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
                }else if (e.KeyChar == '+'){

                    TextBox _TextBox = (sender as TextBox);
                    int saveCursor = _TextBox.Text.Length - _TextBox.SelectionStart;
                    if (_TextBox.Text.IndexOf('-') > -1)
                        _TextBox.Text = _TextBox.Text.Replace("-", "");
                    _TextBox.SelectionStart = _TextBox.Text.Length - saveCursor;
                }
                e.Handled = true;
            }
 
        }

        
        //CS 관리 수수료 처리용 끝
        //도매 직원급여 총액 처리용 시작
        private void txtDetail9_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            if (radioButton1.Checked)
            {
                txtDetail11.Text = CommonUtil.Division(txtDetail9.Text.Replace(",", ""), txtInput11.Text.Replace(",", ""));
            }
        }

        private void txtDetail10_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            if (radioButton1.Checked)
            {
                txtDetail12.Text = CommonUtil.Division(txtDetail10.Text.Replace(",", ""), txtInput12.Text.Replace(",", ""));
            }
        }

        private void txtDetail11_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            if (radioButton2.Checked)
            {
                txtDetail9.Text = (CommonUtil.StringToIntVal(txtDetail11.Text.Replace(",", "")) * CommonUtil.StringToIntVal(txtInput11.Text.Replace(",", ""))).ToString();
            }

        }

        private void txtDetail12_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            if (radioButton2.Checked)
            {
                txtDetail10.Text = (CommonUtil.StringToIntVal(txtDetail12.Text.Replace(",", "")) * CommonUtil.StringToIntVal(txtInput12.Text.Replace(",", ""))).ToString();
            }

        }
        //도매 직원급여 총액 처리용 끝

        //소매 직원급여 총액 처리용 시작
        private void txtDetail21_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            if (radioButton4.Checked)
            {
                txtDetail23.Text = CommonUtil.Division(txtDetail21.Text.Replace(",", ""), txtInput19.Text.Replace(",", ""));
            }

        }

        private void txtDetail22_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
            
            if (radioButton4.Checked)
            {
                txtDetail24.Text = CommonUtil.Division(txtDetail22.Text.Replace(",", ""), txtInput20.Text.Replace(",", ""));
            }

        }

        private void txtDetail23_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
            
            if (radioButton3.Checked)
            {
                txtDetail21.Text = (CommonUtil.StringToIntVal(txtDetail23.Text.Replace(",", "")) * CommonUtil.StringToIntVal(txtInput19.Text.Replace(",", ""))).ToString();
            }

        }

        private void txtDetail24_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
            
            if (radioButton3.Checked)
            {
                txtDetail22.Text = (CommonUtil.StringToIntVal(txtDetail24.Text.Replace(",", "")) * CommonUtil.StringToIntVal(txtInput20.Text.Replace(",", ""))).ToString();
            }

        }




        //소매 직원급여 총액 처리용 끝


        private void txtDetail1_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail4_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail5_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

        }

        private void txtDetail6_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

        }

        private void txtDetail7_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

        }

        private void txtDetail8_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail13_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail14_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail15_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail16_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail17_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail18_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail19_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail20_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail25_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail26_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail27_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail28_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail29_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail30_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail31_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail32_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail33_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail34_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail35_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
        }

        private void txtDetail36_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);
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
    }
}
