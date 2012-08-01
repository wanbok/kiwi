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

        private string[] txtMangeInput = null;

        public FormUserInput()
        {
            InitializeComponent();
            txtMangeInput = new string[31] { txtMangeInput1, txtMangeInput2, txtMangeInput3, txtMangeInput4, txtMangeInput5, 
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

            if (!string.IsNullOrEmpty(CommonUtil.openAsName))
            {
                CommonUtil.ReadExcelFileToData();
                getInput();
                getDetail(CDataControl.g_FileBasicInput);
            }
            else
            {
                CommonUtil.clearTextBox(this.tabPage1);
                CommonUtil.clearTextBox(this.tabPage5);
            }



        }

        //상세입력
        private void getDetail(CBasicInput g_BasicInput)
        {
            //
            Int64[] arrvalue = CDataControl.g_FileDetailInput.getArrData_DetailInput(g_BasicInput.get도매_직원수_간부급(), g_BasicInput.get도매_직원수_평사원()
                ,g_BasicInput.get소매_직원수_간부급(), g_BasicInput.get소매_직원수_평사원());

            // 셀에서 데이터 가져오기
            for (int i = 0; i < txtDetailInput.Length; i++)
            {
                txtDetailInput[i].Text = arrvalue[i].ToString();
            }

        }

        //기본입력
        private void getInput()
        {
            comboBox1.SelectedItem = CDataControl.g_FileBasicInput.get지역();
            textBox6.Text = CDataControl.g_FileBasicInput.get대리점();
            textBox9.Text = CDataControl.g_FileBasicInput.get마케터();
            Int64[] arrvalue = CDataControl.g_FileBasicInput.getArrData_BasicInput();
            // 셀에서 데이터 가져오기
            for (int i = 0; i < txtBasicInput.Length; i++)
            {
                txtBasicInput[i].Text = arrvalue[i].ToString();
            }
        }

        public void SaveAsInput()
        {
            CDataControl.g_BasicInput.set지역(comboBox1.SelectedIndex == -1 ? "" : comboBox1.Items[comboBox1.SelectedIndex].ToString());
            CDataControl.g_BasicInput.set대리점(textBox6.Text);
            CDataControl.g_BasicInput.set마케터(textBox9.Text);

            String[] txtWrite = new String[14] { txtInput1.Text, txtInput2.Text, txtInput3.Text, txtInput5.Text, txtInput6.Text,  
                txtInput8.Text, txtInput9.Text, txtInput11.Text, txtInput12.Text, txtInput14.Text, txtInput15.Text, txtInput17.Text, txtInput19.Text, txtInput20.Text};
            CDataControl.g_BasicInput.setArrData_BasicInput(txtWrite);

            String[] txtWrite2 = new String[31]  { txtDetail1.Text, txtDetail2.Text, txtDetail4.Text, txtDetail5.Text, txtDetail6.Text,
                txtDetail7.Text, txtDetail8.Text, txtDetail9.Text, txtDetail10.Text, txtDetail11.Text,
                txtDetail12.Text, txtDetail15.Text, txtDetail16.Text, txtDetail17.Text, txtDetail18.Text,
                txtDetail19.Text, txtDetail20.Text, txtDetail23.Text, txtDetail24.Text, txtDetail25.Text, 
                txtDetail26.Text, txtDetail27.Text, txtDetail28.Text, txtDetail29.Text, txtDetail30.Text,            
                txtDetail31.Text, txtDetail32.Text, txtDetail33.Text, txtDetail34.Text, txtDetail35.Text, txtDetail36.Text
            };
            CDataControl.g_DetailInput.setArrData_DetailInput(txtWrite2);
            CommonUtil.ReadFileManagerToData();
            txtMangeInput = CDataControl.g_BusinessAvg.getArrData_BusinessAvg();


            ////업계 평균적용 결과 단위당 금액
            Int64 sumSubDE = 0;
            CDataControl.g_ResultBusiness.전체_수익_가입자수수료 = CommonUtil.StringToIntVal(txtMangeInput[0]);
            sumSubDE += Convert.ToInt64(txtMangeInput[0]);
            CDataControl.g_ResultBusiness.전체_수익_CS관리수수료 = CommonUtil.StringToIntVal(txtMangeInput[1]);
            sumSubDE += Convert.ToInt64(txtMangeInput[1]);
            CDataControl.g_ResultBusiness.전체_수익_업무취급수수료 = CommonUtil.StringToIntVal(txtMangeInput[15]);
            sumSubDE += Convert.ToInt64(txtMangeInput[16]);
            CDataControl.g_ResultBusiness.전체_수익_사업자모델매입에따른추가수익 = CommonUtil.StringToIntVal(txtMangeInput[2]);
            sumSubDE += Convert.ToInt64(txtMangeInput[2]);
            CDataControl.g_ResultBusiness.전체_수익_유통모델매입에따른추가수익_현금_Volume = (Convert.ToInt64(txtMangeInput[3]) + Convert.ToInt64(txtMangeInput[4]));
            sumSubDE += Convert.ToInt64(Convert.ToInt64(txtMangeInput[3]) + Convert.ToInt64(txtMangeInput[4]));
            CDataControl.g_ResultBusiness.전체_수익_직영매장판매수익 = CommonUtil.StringToIntVal(txtMangeInput[16]);
            sumSubDE += Convert.ToInt64(txtMangeInput[17]);
            CDataControl.g_ResultBusiness.전체_수익_소계 = sumSubDE;

            Int64 sumSubCo = 0;
            CDataControl.g_ResultBusiness.set전체_비용_대리점투자비용(CommonUtil.Division((Convert.ToInt64(txtMangeInput[5]) * CommonUtil.StringToIntVal(txtInput2.Text)
                + Convert.ToInt64(txtMangeInput[6]) * CommonUtil.StringToIntVal(txtInput3.Text)).ToString(), txtInput4.Text));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division((Convert.ToInt64(txtMangeInput[5]) * CommonUtil.StringToIntVal(txtInput2.Text)
                + Convert.ToInt64(txtMangeInput[6]) * CommonUtil.StringToIntVal(txtInput3.Text)).ToString(), txtInput4.Text));

            CDataControl.g_ResultBusiness.set전체_비용_인건비_급여_복리후생비(CommonUtil.StringToIntVal(CommonUtil.Division((Convert.ToInt64(txtMangeInput[7]) * CommonUtil.StringToIntVal(txtInput11.Text)
                + Convert.ToInt64(txtMangeInput[8]) * CommonUtil.StringToIntVal(txtInput12.Text)
                + Convert.ToInt64(txtMangeInput[17]) * CommonUtil.StringToIntVal(txtInput19.Text)
                + Convert.ToInt64(txtMangeInput[18]) * CommonUtil.StringToIntVal(txtInput20.Text)
                + Convert.ToInt64(txtMangeInput[23]) * CommonUtil.StringToIntVal(txtInput25.Text)).ToString(), txtInput25.Text)));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division((Convert.ToInt64(txtMangeInput[7]) * CommonUtil.StringToIntVal(txtInput11.Text)
                + Convert.ToInt64(txtMangeInput[8]) * CommonUtil.StringToIntVal(txtInput12.Text)
                + Convert.ToInt64(txtMangeInput[17]) * CommonUtil.StringToIntVal(txtInput19.Text)
                + Convert.ToInt64(txtMangeInput[18]) * CommonUtil.StringToIntVal(txtInput20.Text)
                + Convert.ToInt64(txtMangeInput[23]) * CommonUtil.StringToIntVal(txtInput25.Text)).ToString(), txtInput25.Text));


            CDataControl.g_ResultBusiness.set전체_비용_임차료(Convert.ToInt64(txtMangeInput[9]) + Convert.ToInt64(txtMangeInput[19]));
            sumSubCo += Convert.ToInt64(txtMangeInput[9]) + Convert.ToInt64(txtMangeInput[19]);
            CDataControl.g_ResultBusiness.set전체_비용_이자비용(Convert.ToInt64(txtMangeInput[27]));
            sumSubCo += Convert.ToInt64(txtMangeInput[27]);
            CDataControl.g_ResultBusiness.set전체_비용_부가세(Convert.ToInt64(txtMangeInput[28]));
            sumSubCo += Convert.ToInt64(txtMangeInput[28]);
            CDataControl.g_ResultBusiness.set전체_비용_법인세(Convert.ToInt64(txtMangeInput[29]));
            sumSubCo += Convert.ToInt64(txtMangeInput[29]);

            CDataControl.g_ResultBusiness.set전체_비용_기타판매관리비(CommonUtil.StringToIntVal(txtMangeInput[10]) + CommonUtil.StringToIntVal(txtMangeInput[11])
                + CommonUtil.StringToIntVal(txtMangeInput[12]) + CommonUtil.StringToIntVal(txtMangeInput[13])
                + CommonUtil.StringToIntVal(txtMangeInput[14]) + CommonUtil.StringToIntVal(txtMangeInput[20]) 
                + CommonUtil.StringToIntVal(txtMangeInput[21]) + CommonUtil.StringToIntVal(txtMangeInput[22])
                + CommonUtil.StringToIntVal(txtMangeInput[24]) + CommonUtil.StringToIntVal(txtMangeInput[25])
                + CommonUtil.StringToIntVal(txtMangeInput[26]) + CommonUtil.StringToIntVal(txtMangeInput[30]));
            sumSubCo += CommonUtil.StringToIntVal(txtMangeInput[10]) + CommonUtil.StringToIntVal(txtMangeInput[11])
                + CommonUtil.StringToIntVal(txtMangeInput[12]) + CommonUtil.StringToIntVal(txtMangeInput[13])
                + CommonUtil.StringToIntVal(txtMangeInput[14]) + CommonUtil.StringToIntVal(txtMangeInput[20])
                + CommonUtil.StringToIntVal(txtMangeInput[21]) + CommonUtil.StringToIntVal(txtMangeInput[22])
                + CommonUtil.StringToIntVal(txtMangeInput[24]) + CommonUtil.StringToIntVal(txtMangeInput[25])
                + CommonUtil.StringToIntVal(txtMangeInput[26]) + CommonUtil.StringToIntVal(txtMangeInput[30]);
            CDataControl.g_ResultBusiness.전체_비용_소계 = sumSubCo;
            CDataControl.g_ResultBusiness.전체손익계 = sumSubDE - sumSubCo;

            sumSubDE = 0;
            //도매 수익
            CDataControl.g_ResultBusiness.set도매_수익_가입자관리수수료( txtMangeInput[0]);
            sumSubDE += Convert.ToInt64(txtMangeInput[0]);
            CDataControl.g_ResultBusiness.set도매_수익_CS관리수수료( txtMangeInput[1]);
            sumSubDE += Convert.ToInt64(txtMangeInput[1]);
           CDataControl.g_ResultBusiness.set도매_수익_사업자모델매입에따른추가수익( txtMangeInput[2]);
            sumSubDE += Convert.ToInt64(txtMangeInput[2]);
            CDataControl.g_ResultBusiness.set도매_수익_유통모델매입에따른추가수익_현금_Volume(Convert.ToInt64(txtMangeInput[3]) + Convert.ToInt64(txtMangeInput[4]));
            sumSubDE += Convert.ToInt64(Convert.ToInt64(txtMangeInput[3]) + Convert.ToInt64(txtMangeInput[4]));
            CDataControl.g_ResultBusiness.도매_수익_소계 = sumSubDE;
            //도매비용
            sumSubCo = 0;
            CDataControl.g_ResultBusiness.set도매_비용_대리점투자비용( CommonUtil.Division((Convert.ToInt64(txtMangeInput[5]) * CommonUtil.StringToIntVal(txtInput2.Text)
                + Convert.ToInt64(txtMangeInput[6]) * CommonUtil.StringToIntVal(txtInput3.Text)).ToString(), txtInput4.Text));
            sumSubCo += CommonUtil.StringToIntVal( CommonUtil.Division((Convert.ToInt64(txtMangeInput[5]) * CommonUtil.StringToIntVal(txtInput2.Text)
                + Convert.ToInt64(txtMangeInput[6]) * CommonUtil.StringToIntVal(txtInput3.Text)).ToString(), txtInput4.Text));
            CDataControl.g_ResultBusiness.set도매_비용_인건비_급여_복리후생비( CommonUtil.StringToIntVal(CommonUtil.Division((Convert.ToInt64(txtMangeInput[7]) * CommonUtil.StringToIntVal(txtInput11.Text)
                + Convert.ToInt64(txtMangeInput[8]) * CommonUtil.StringToIntVal(txtInput12.Text)
                + Convert.ToInt64(txtMangeInput[23]) * CommonUtil.StringToIntVal(txtInput13.Text)).ToString(), txtInput13.Text)));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division((Convert.ToInt64(txtMangeInput[7]) * CommonUtil.StringToIntVal(txtInput11.Text)
                + Convert.ToInt64(txtMangeInput[8]) * CommonUtil.StringToIntVal(txtInput12.Text)
                + Convert.ToInt64(txtMangeInput[23]) * CommonUtil.StringToIntVal(txtInput13.Text)).ToString(), txtInput13.Text));
            CDataControl.g_ResultBusiness.set도매_비용_임차료( txtMangeInput[9]);
            sumSubCo += Convert.ToInt64(txtMangeInput[9]);
           CDataControl.g_ResultBusiness.set도매_비용_이자비용( CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[27], txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[27], txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text);
            CDataControl.g_ResultBusiness.set도매_비용_부가세( CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[28], txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[28], txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text);
            CDataControl.g_ResultBusiness.set도매_비용_법인세( CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[29], txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[29], txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text);
            CDataControl.g_ResultBusiness.set도매_비용_기타판매관리비( CommonUtil.StringToIntVal(CommonUtil.Division( (CommonUtil.StringToIntVal(txtMangeInput[10]) + CommonUtil.StringToIntVal(txtMangeInput[11])
                + CommonUtil.StringToIntVal(txtMangeInput[12]) + CommonUtil.StringToIntVal(txtMangeInput[13])
                + CommonUtil.StringToIntVal(txtMangeInput[14]) + CommonUtil.StringToIntVal(txtMangeInput[24])
                + CommonUtil.StringToIntVal(txtMangeInput[25]) + CommonUtil.StringToIntVal(txtMangeInput[26])
                + CommonUtil.StringToIntVal(txtMangeInput[30])).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtMangeInput[10]) + CommonUtil.StringToIntVal(txtMangeInput[11])
                + CommonUtil.StringToIntVal(txtMangeInput[12]) + CommonUtil.StringToIntVal(txtMangeInput[13])
                + CommonUtil.StringToIntVal(txtMangeInput[14]) + CommonUtil.StringToIntVal(txtMangeInput[24])
                + CommonUtil.StringToIntVal(txtMangeInput[25]) + CommonUtil.StringToIntVal(txtMangeInput[26])
                + CommonUtil.StringToIntVal(txtMangeInput[30])).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text); 

            CDataControl.g_ResultBusiness.도매_비용_소계 = sumSubCo;
            CDataControl.g_ResultBusiness.도매손익계 =  sumSubDE - sumSubCo;

            //소매
            sumSubDE = 0;
            CDataControl.g_ResultBusiness.set소매_수익_업무취급수수료( txtMangeInput[15]);
            sumSubDE += Convert.ToInt64(txtMangeInput[16]);
            CDataControl.g_ResultBusiness.set소매_수익_직영매장판매수익(txtMangeInput[16]);
            sumSubDE += Convert.ToInt64(txtMangeInput[17]);

            CDataControl.g_ResultBusiness.소매_수익_소계 = sumSubDE;

            sumSubCo = 0;
            CDataControl.g_ResultBusiness.set소매_비용_인건비_급여_복리후생비( CommonUtil.StringToIntVal(CommonUtil.Division((Convert.ToInt64(txtMangeInput[17]) * CommonUtil.StringToIntVal(txtInput19.Text)
                + Convert.ToInt64(txtMangeInput[18]) * CommonUtil.StringToIntVal(txtInput20.Text)
                + Convert.ToInt64(txtMangeInput[23]) * CommonUtil.StringToIntVal(txtInput21.Text)).ToString(), txtInput21.Text)));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division((Convert.ToInt64(txtMangeInput[17]) * CommonUtil.StringToIntVal(txtInput19.Text)
                + Convert.ToInt64(txtMangeInput[18]) * CommonUtil.StringToIntVal(txtInput20.Text)
                + Convert.ToInt64(txtMangeInput[23]) * CommonUtil.StringToIntVal(txtInput21.Text)).ToString(), txtInput21.Text));
            CDataControl.g_ResultBusiness.set소매_비용_임차료(txtMangeInput[19]);
            sumSubCo += Convert.ToInt64(txtMangeInput[19]);
            CDataControl.g_ResultBusiness.set소매_비용_이자비용( CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[27], txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[27], txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text);
            CDataControl.g_ResultBusiness.set소매_비용_부가세( CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[28], txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[28], txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text);
            CDataControl.g_ResultBusiness.set소매_비용_법인세(CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[29], txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division(txtMangeInput[29], txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text);



            CDataControl.g_ResultBusiness.set소매_비용_기타판매관리비( CommonUtil.StringToIntVal(CommonUtil.Division( (CommonUtil.StringToIntVal(txtMangeInput[20]) + CommonUtil.StringToIntVal(txtMangeInput[21])
                + CommonUtil.StringToIntVal(txtMangeInput[22]) + CommonUtil.StringToIntVal(txtMangeInput[24])
                + CommonUtil.StringToIntVal(txtMangeInput[25]) + CommonUtil.StringToIntVal(txtMangeInput[26])
                + CommonUtil.StringToIntVal(txtMangeInput[30])).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text));
            sumSubCo += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtMangeInput[20]) + CommonUtil.StringToIntVal(txtMangeInput[21])
                + CommonUtil.StringToIntVal(txtMangeInput[22]) + CommonUtil.StringToIntVal(txtMangeInput[24])
                + CommonUtil.StringToIntVal(txtMangeInput[25]) + CommonUtil.StringToIntVal(txtMangeInput[26])
                + CommonUtil.StringToIntVal(txtMangeInput[30])).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text);

            CDataControl.g_ResultBusiness.소매_비용_소계 = sumSubCo;
            CDataControl.g_ResultBusiness.소매손익계 = sumSubDE - sumSubCo;
            CDataControl.g_ResultBusiness.점별손익추정 = CDataControl.g_BasicInput.get소매_거래선수_직영점();


            //업계 평균적용 결과 총액
            Int64[] tempInt = new Int64[42];
            for (int i = 0; i < CDataControl.g_ResultBusiness.getArrayOutput전체().Length; i++)
            {
                if (i >= 0 && i < 41)
                {
                    string temp = "0";
                    string txtInput = "0";
                    if (i >= 0 && i < 16)
                    {
                        txtInput = txtInput25.Text;
                        temp = CDataControl.g_ResultBusiness.getArrayOutput전체()[i].ToString();
                    }
                    else if (i >= 16 && i < 30)
                    {
                        txtInput = txtInput4.Text;
                        temp = CDataControl.g_ResultBusiness.getArrayOutput전체()[i].ToString();
                    }
                    else if (i >= 30 && i < 41)
                    {
                        txtInput = txtInput16.Text;
                        temp = CDataControl.g_ResultBusiness.getArrayOutput전체()[i].ToString();
                    }

                    tempInt[i] = CommonUtil.StringToIntVal(temp) * CommonUtil.StringToIntVal(txtInput);
                }
                else if (i == 41)
                {
                    Int64 tempStore = CommonUtil.StringToIntVal(CDataControl.g_ResultBusiness.getArrayOutput전체()[i - 1].ToString()) * CommonUtil.StringToIntVal(txtInput16.Text);
                    tempInt[i] = CommonUtil.StringToIntVal(CommonUtil.Division(tempStore.ToString(), txtInput30.Text));
                }

            }

            //당대리점 결과(세부항목별 값 입력 결과) 수익계정
            Int64 SumSubBenefitTotal = 0;
            CDataControl.g_ResultStoreTotal.전체_수익_가입자수수료 = CommonUtil.StringToIntVal( txtDetail1.Text);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(txtDetail1.Text);
            CDataControl.g_ResultStore.전체_수익_가입자수수료 = CommonUtil.StringToIntVal(  CommonUtil.Division(txtDetail1.Text, txtInput25.Text));
            CDataControl.g_ResultStoreTotal.전체_수익_CS관리수수료 = CommonUtil.StringToIntVal(  txtDetail2.Text);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(txtDetail2.Text);
            CDataControl.g_ResultStore.전체_수익_CS관리수수료 = CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail2.Text, txtInput25.Text));
            CDataControl.g_ResultStoreTotal.전체_수익_업무취급수수료 = CommonUtil.StringToIntVal(  txtDetail19.Text);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(txtDetail19.Text);
            CDataControl.g_ResultStore.전체_수익_업무취급수수료 = CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail19.Text, txtInput25.Text));
            CDataControl.g_ResultStoreTotal.전체_수익_사업자모델매입에따른추가수익 = CommonUtil.StringToIntVal(  txtDetail4.Text);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(txtDetail4.Text);
            CDataControl.g_ResultStore.전체_수익_사업자모델매입에따른추가수익 = CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail4.Text, txtInput25.Text));
            CDataControl.g_ResultStoreTotal.전체_수익_유통모델매입에따른추가수익_현금_Volume = CommonUtil.StringToIntVal(  (CommonUtil.StringToIntVal(txtDetail5.Text) + CommonUtil.StringToIntVal(txtDetail6.Text)).ToString());
            SumSubBenefitTotal += (CommonUtil.StringToIntVal(txtDetail5.Text) + CommonUtil.StringToIntVal(txtDetail6.Text));
            CDataControl.g_ResultStore.전체_수익_유통모델매입에따른추가수익_현금_Volume = CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail5.Text) + CommonUtil.StringToIntVal(txtDetail6.Text)).ToString(), txtInput25.Text));
            CDataControl.g_ResultStoreTotal.전체_수익_직영매장판매수익 = CommonUtil.StringToIntVal(  txtDetail20.Text);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(txtDetail20.Text);
            CDataControl.g_ResultStore.전체_수익_직영매장판매수익 = CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail20.Text, txtInput25.Text));
            CDataControl.g_ResultStoreTotal.전체_수익_소계 = CommonUtil.StringToIntVal(  SumSubBenefitTotal.ToString());
            CDataControl.g_ResultStore.전체_수익_소계 = CommonUtil.StringToIntVal(CommonUtil.Division(SumSubBenefitTotal.ToString(), txtInput25.Text));

            //당대리점 결과(세부항목별 값 입력 결과) 비용계정
            Int64 SumSubCostTotal = 0;
            CDataControl.g_ResultStoreTotal.set전체_비용_대리점투자비용( (CommonUtil.StringToIntVal(txtDetail7.Text) * CommonUtil.StringToIntVal(txtInput2.Text) + (CommonUtil.StringToIntVal(txtDetail8.Text) * CommonUtil.StringToIntVal(txtInput3.Text))));
            SumSubCostTotal += (CommonUtil.StringToIntVal(txtDetail7.Text) * CommonUtil.StringToIntVal(txtInput2.Text) + (CommonUtil.StringToIntVal(txtDetail8.Text) * CommonUtil.StringToIntVal(txtInput3.Text)));
            CDataControl.g_ResultStore.set전체_비용_대리점투자비용( CommonUtil.Division(txtDetail1.Text, txtInput25.Text));

            CDataControl.g_ResultStoreTotal.set전체_비용_인건비_급여_복리후생비(  (CommonUtil.StringToIntVal(txtDetail11.Text) * CommonUtil.StringToIntVal(txtInput11.Text)
                + CommonUtil.StringToIntVal(txtDetail12.Text) * CommonUtil.StringToIntVal(txtInput12.Text)
                + CommonUtil.StringToIntVal(txtDetail23.Text) * CommonUtil.StringToIntVal(txtInput19.Text)
                + CommonUtil.StringToIntVal(txtDetail24.Text) * CommonUtil.StringToIntVal(txtInput20.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput25.Text)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(txtDetail11.Text) * CommonUtil.StringToIntVal(txtInput11.Text) 
                + CommonUtil.StringToIntVal(txtDetail12.Text) * CommonUtil.StringToIntVal(txtInput12.Text) 
                + CommonUtil.StringToIntVal(txtDetail23.Text) * CommonUtil.StringToIntVal(txtInput19.Text)
                + CommonUtil.StringToIntVal(txtDetail24.Text) * CommonUtil.StringToIntVal(txtInput20.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput25.Text));
            CDataControl.g_ResultStore.set전체_비용_인건비_급여_복리후생비( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail11.Text) * CommonUtil.StringToIntVal(txtInput11.Text)
                + CommonUtil.StringToIntVal(txtDetail12.Text) * CommonUtil.StringToIntVal(txtInput12.Text)
                + CommonUtil.StringToIntVal(txtDetail23.Text) * CommonUtil.StringToIntVal(txtInput19.Text)
                + CommonUtil.StringToIntVal(txtDetail24.Text) * CommonUtil.StringToIntVal(txtInput20.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput25.Text)).ToString(), txtInput25.Text));
            CDataControl.g_ResultStoreTotal.set전체_비용_임차료(  (CommonUtil.StringToIntVal(txtDetail13.Text) + CommonUtil.StringToIntVal(txtDetail25.Text)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(txtDetail13.Text) + CommonUtil.StringToIntVal(txtDetail25.Text));
            CDataControl.g_ResultStore.set전체_비용_임차료( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail13.Text) + CommonUtil.StringToIntVal(txtDetail25.Text)).ToString(), txtInput25.Text));
            CDataControl.g_ResultStoreTotal.set전체_비용_이자비용(  txtDetail33.Text);
            SumSubCostTotal += CommonUtil.StringToIntVal(txtDetail33.Text);
            CDataControl.g_ResultStore.set전체_비용_이자비용( CommonUtil.Division(txtDetail33.Text, txtInput25.Text));
            CDataControl.g_ResultStoreTotal.set전체_비용_부가세(  (CommonUtil.StringToIntVal(txtDetail34.Text)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(txtDetail34.Text));
            CDataControl.g_ResultStore.set전체_비용_부가세( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail34.Text)).ToString(), txtInput25.Text));
            CDataControl.g_ResultStoreTotal.set전체_비용_법인세(  txtDetail35.Text);
            SumSubCostTotal += CommonUtil.StringToIntVal(txtDetail35.Text);
            CDataControl.g_ResultStore.set전체_비용_법인세( CommonUtil.Division(txtDetail35.Text, txtInput25.Text));

            CDataControl.g_ResultStoreTotal.set전체_비용_기타판매관리비(  (CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text));
           CDataControl.g_ResultStore.set전체_비용_기타판매관리비( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text)).ToString(), txtInput25.Text));
            CDataControl.g_ResultStoreTotal.전체_비용_소계 =  SumSubCostTotal;
            CDataControl.g_ResultStore.전체_비용_소계 =  CommonUtil.StringToIntVal(CommonUtil.Division(SumSubCostTotal.ToString(), txtInput25.Text));
            CDataControl.g_ResultStoreTotal.전체손익계 = SumSubBenefitTotal - SumSubCostTotal;
            CDataControl.g_ResultStore.전체손익계 =  CommonUtil.StringToIntVal(CommonUtil.Division((SumSubBenefitTotal - SumSubCostTotal).ToString(), txtInput25.Text));


            Int64 SumSubBenefitWillTotal = 0;

            CDataControl.g_ResultFuture.전체_수익_가입자수수료 = CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail1.Text, txtInput1.Text)) * 18;
            CDataControl.g_ResultFutureTotal.전체_수익_가입자수수료 = CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail1.Text, txtInput1.Text)) * 18 * CommonUtil.StringToIntVal(txtInput25.Text);
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail1.Text, txtInput1.Text)) * 18 * CommonUtil.StringToIntVal(txtInput25.Text);
            CDataControl.g_ResultFuture.전체_수익_CS관리수수료 = CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail2.Text, txtInput1.Text)) * 18;
            CDataControl.g_ResultFutureTotal.전체_수익_CS관리수수료 = CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail2.Text, txtInput1.Text)) * 18 * CommonUtil.StringToIntVal(txtInput25.Text);
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail2.Text, txtInput1.Text)) * 18 * CommonUtil.StringToIntVal(txtInput25.Text);
            CDataControl.g_ResultFutureTotal.전체_수익_업무취급수수료 = CommonUtil.StringToIntVal(txtDetail19.Text);
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(txtDetail19.Text);
            CDataControl.g_ResultFuture.전체_수익_업무취급수수료 =  CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail19.Text, txtInput25.Text));
            CDataControl.g_ResultFutureTotal.전체_수익_사업자모델매입에따른추가수익 =  CommonUtil.StringToIntVal(txtDetail4.Text);
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(txtDetail4.Text);
            CDataControl.g_ResultFuture.전체_수익_사업자모델매입에따른추가수익 =  CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail4.Text, txtInput25.Text));
            CDataControl.g_ResultFutureTotal.전체_수익_유통모델매입에따른추가수익_현금_Volume = (CommonUtil.StringToIntVal(txtDetail5.Text) + CommonUtil.StringToIntVal(txtDetail6.Text));
            SumSubBenefitWillTotal += (CommonUtil.StringToIntVal(txtDetail5.Text) + CommonUtil.StringToIntVal(txtDetail6.Text));
            CDataControl.g_ResultFuture.전체_수익_유통모델매입에따른추가수익_현금_Volume =  CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail5.Text) + CommonUtil.StringToIntVal(txtDetail6.Text)).ToString(), txtInput25.Text));
            CDataControl.g_ResultFutureTotal.전체_수익_직영매장판매수익 =  CommonUtil.StringToIntVal(txtDetail20.Text);
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(txtDetail20.Text);
            CDataControl.g_ResultFuture.전체_수익_직영매장판매수익 =  CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail20.Text, txtInput25.Text));
            CDataControl.g_ResultFutureTotal.전체_수익_소계 = CommonUtil.StringToIntVal(SumSubBenefitWillTotal.ToString());
            CDataControl.g_ResultFuture.전체_수익_소계 =  CommonUtil.StringToIntVal(CommonUtil.Division(SumSubBenefitWillTotal.ToString(), txtInput25.Text));


            Int64 SumSubCostWillTotal = 0;
            CDataControl.g_ResultFutureTotal.set전체_비용_대리점투자비용( (CommonUtil.StringToIntVal(txtDetail7.Text) * CommonUtil.StringToIntVal(txtInput2.Text) + (CommonUtil.StringToIntVal(txtDetail8.Text) * CommonUtil.StringToIntVal(txtInput3.Text))));
            SumSubCostWillTotal += (CommonUtil.StringToIntVal(txtDetail7.Text) * CommonUtil.StringToIntVal(txtInput2.Text) + (CommonUtil.StringToIntVal(txtDetail8.Text) * CommonUtil.StringToIntVal(txtInput3.Text)));
            CDataControl.g_ResultFuture.set전체_비용_대리점투자비용( CommonUtil.Division(txtDetail1.Text, txtInput25.Text));
            CDataControl.g_ResultFutureTotal.set전체_비용_인건비_급여_복리후생비( (CommonUtil.StringToIntVal(txtDetail11.Text) * CommonUtil.StringToIntVal(txtInput11.Text)
                + CommonUtil.StringToIntVal(txtDetail12.Text) * CommonUtil.StringToIntVal(txtInput12.Text)
                + CommonUtil.StringToIntVal(txtDetail23.Text) * CommonUtil.StringToIntVal(txtInput19.Text)
                + CommonUtil.StringToIntVal(txtDetail24.Text) * CommonUtil.StringToIntVal(txtInput20.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput25.Text)));
            SumSubCostWillTotal += (CommonUtil.StringToIntVal(txtDetail11.Text) * CommonUtil.StringToIntVal(txtInput11.Text)
                + CommonUtil.StringToIntVal(txtDetail12.Text) * CommonUtil.StringToIntVal(txtInput12.Text)
                + CommonUtil.StringToIntVal(txtDetail23.Text) * CommonUtil.StringToIntVal(txtInput19.Text)
                + CommonUtil.StringToIntVal(txtDetail24.Text) * CommonUtil.StringToIntVal(txtInput20.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput25.Text));
            CDataControl.g_ResultFuture.set전체_비용_인건비_급여_복리후생비( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail11.Text) * CommonUtil.StringToIntVal(txtInput11.Text)
                + CommonUtil.StringToIntVal(txtDetail12.Text) * CommonUtil.StringToIntVal(txtInput12.Text)
                + CommonUtil.StringToIntVal(txtDetail23.Text) * CommonUtil.StringToIntVal(txtInput19.Text)
                + CommonUtil.StringToIntVal(txtDetail24.Text) * CommonUtil.StringToIntVal(txtInput20.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput25.Text)).ToString(), txtInput25.Text));
            CDataControl.g_ResultFutureTotal.set전체_비용_임차료( (CommonUtil.StringToIntVal(txtDetail13.Text) + CommonUtil.StringToIntVal(txtDetail25.Text)));
            SumSubCostWillTotal += (CommonUtil.StringToIntVal(txtDetail13.Text) + CommonUtil.StringToIntVal(txtDetail25.Text));
            CDataControl.g_ResultFuture.set전체_비용_임차료( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail13.Text) + CommonUtil.StringToIntVal(txtDetail25.Text)).ToString(), txtInput25.Text));
            CDataControl.g_ResultFutureTotal.set전체_비용_이자비용( txtDetail33.Text);
            SumSubCostWillTotal += CommonUtil.StringToIntVal(txtDetail33.Text);
            CDataControl.g_ResultFuture.set전체_비용_이자비용( CommonUtil.Division(txtDetail33.Text, txtInput25.Text));
            CDataControl.g_ResultFutureTotal.set전체_비용_부가세( (CommonUtil.StringToIntVal(txtDetail34.Text)));
            SumSubCostWillTotal += (CommonUtil.StringToIntVal(txtDetail34.Text));
            CDataControl.g_ResultFuture.set전체_비용_부가세( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail34.Text)).ToString(), txtInput25.Text));
            CDataControl.g_ResultFutureTotal.set전체_비용_법인세( txtDetail35.Text);
            SumSubCostWillTotal += CommonUtil.StringToIntVal(txtDetail35.Text);
            CDataControl.g_ResultFuture.set전체_비용_법인세( CommonUtil.Division(txtDetail35.Text, txtInput25.Text));
            CDataControl.g_ResultFutureTotal.set전체_비용_기타판매관리비( (CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text)));
            SumSubCostWillTotal += (CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text));
            CDataControl.g_ResultFuture.set전체_비용_기타판매관리비( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text)).ToString(), txtInput25.Text));
            CDataControl.g_ResultFutureTotal.전체_수익_소계 =  SumSubCostWillTotal;
            CDataControl.g_ResultFuture.전체_수익_소계 = CommonUtil.StringToIntVal(CommonUtil.Division(SumSubCostWillTotal.ToString(), txtInput25.Text));
            CDataControl.g_ResultFutureTotal.전체손익계 =  (SumSubBenefitWillTotal - SumSubCostWillTotal);
            CDataControl.g_ResultFuture.전체손익계 = CommonUtil.StringToIntVal(CommonUtil.Division((SumSubBenefitWillTotal - SumSubCostWillTotal).ToString(), txtInput25.Text));






            //도매
            //당대리점 결과(세부항목별 값 입력 결과) 수익계정
            SumSubBenefitTotal = 0;
            CDataControl.g_ResultStoreTotal.set도매_수익_가입자관리수수료( txtDetail1.Text);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(txtDetail1.Text);
            CDataControl.g_ResultStore.set도매_수익_가입자관리수수료( CommonUtil.Division(txtDetail1.Text, txtInput4.Text));
            CDataControl.g_ResultStoreTotal.set도매_수익_CS관리수수료( txtDetail2.Text);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(txtDetail2.Text);
            CDataControl.g_ResultStore.set도매_수익_CS관리수수료( CommonUtil.Division(txtDetail2.Text, txtInput4.Text));
            CDataControl.g_ResultStoreTotal.set도매_수익_사업자모델매입에따른추가수익( txtDetail4.Text);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(txtDetail4.Text);
            CDataControl.g_ResultStore.set도매_수익_사업자모델매입에따른추가수익( CommonUtil.Division(txtDetail4.Text, txtInput4.Text));
            CDataControl.g_ResultStoreTotal.set도매_수익_유통모델매입에따른추가수익_현금_Volume( (CommonUtil.StringToIntVal(txtDetail5.Text) + CommonUtil.StringToIntVal(txtDetail6.Text)).ToString());
            SumSubBenefitTotal += (CommonUtil.StringToIntVal(txtDetail5.Text) + CommonUtil.StringToIntVal(txtDetail6.Text));
            CDataControl.g_ResultStore.set도매_수익_유통모델매입에따른추가수익_현금_Volume( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail5.Text) + CommonUtil.StringToIntVal(txtDetail6.Text)).ToString(), txtInput4.Text));
            CDataControl.g_ResultStoreTotal.도매_수익_소계 = SumSubBenefitTotal;
            CDataControl.g_ResultStore.도매_수익_소계 = CommonUtil.StringToIntVal( CommonUtil.Division(SumSubBenefitTotal.ToString(), txtInput4.Text));

            //당대리점 결과(세부항목별 값 입력 결과) 비용계정
            SumSubCostTotal = 0;
            CDataControl.g_ResultStoreTotal.set도매_비용_대리점투자비용( (CommonUtil.StringToIntVal(txtDetail7.Text) * CommonUtil.StringToIntVal(txtInput2.Text) + (CommonUtil.StringToIntVal(txtDetail8.Text) * CommonUtil.StringToIntVal(txtInput3.Text))));
            SumSubCostTotal += (CommonUtil.StringToIntVal(txtDetail7.Text) * CommonUtil.StringToIntVal(txtInput2.Text) + (CommonUtil.StringToIntVal(txtDetail8.Text) * CommonUtil.StringToIntVal(txtInput3.Text)));
            CDataControl.g_ResultStore.set도매_비용_대리점투자비용( CommonUtil.Division(txtDetail1.Text, txtInput4.Text));
            CDataControl.g_ResultStoreTotal.set도매_비용_인건비_급여_복리후생비( (CommonUtil.StringToIntVal(txtDetail11.Text) * CommonUtil.StringToIntVal(txtInput11.Text)
                + CommonUtil.StringToIntVal(txtDetail12.Text) * CommonUtil.StringToIntVal(txtInput12.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput13.Text)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(txtDetail11.Text) * CommonUtil.StringToIntVal(txtInput11.Text)
                + CommonUtil.StringToIntVal(txtDetail12.Text) * CommonUtil.StringToIntVal(txtInput12.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput13.Text));
            CDataControl.g_ResultStore.set도매_비용_인건비_급여_복리후생비( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail11.Text) * CommonUtil.StringToIntVal(txtInput11.Text)
                + CommonUtil.StringToIntVal(txtDetail12.Text) * CommonUtil.StringToIntVal(txtInput12.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput13.Text)).ToString(), txtInput4.Text));
            CDataControl.g_ResultStoreTotal.set도매_비용_임차료( (CommonUtil.StringToIntVal(txtDetail13.Text)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(txtDetail13.Text));
            CDataControl.g_ResultStore.set도매_비용_임차료( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail13.Text)).ToString(), txtInput4.Text));
            CDataControl.g_ResultStoreTotal.set도매_비용_이자비용( CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail33.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail33.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text);
            CDataControl.g_ResultStore.set도매_비용_이자비용( CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail33.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text)).ToString(), txtInput4.Text));
            
            CDataControl.g_ResultStoreTotal.set도매_비용_부가세( CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail34.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail34.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text); 
            CDataControl.g_ResultStore.set도매_비용_부가세( CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail34.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text)).ToString(), txtInput4.Text));
            CDataControl.g_ResultStoreTotal.set도매_비용_법인세( CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail35.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail35.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text);
            CDataControl.g_ResultStore.set도매_비용_법인세( CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail35.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text)).ToString(), txtInput4.Text));
            CDataControl.g_ResultStoreTotal.set도매_비용_기타판매관리비( CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text) );
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text);
            CDataControl.g_ResultStore.set도매_비용_기타판매관리비( CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text)).ToString(), txtInput4.Text));
            CDataControl.g_ResultStoreTotal.도매_비용_소계 = SumSubCostTotal;
            CDataControl.g_ResultStore.도매_비용_소계 = CommonUtil.StringToIntVal(CommonUtil.Division(SumSubCostTotal.ToString(), txtInput4.Text));
            CDataControl.g_ResultStoreTotal.도매손익계 =  (SumSubBenefitTotal - SumSubCostTotal);
            CDataControl.g_ResultStore.도매손익계 =  CommonUtil.StringToIntVal(CommonUtil.Division((SumSubBenefitTotal - SumSubCostTotal).ToString(), txtInput4.Text));


            SumSubBenefitWillTotal = 0;
            CDataControl.g_ResultFutureTotal.set도매_수익_가입자관리수수료( CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail1.Text, txtInput1.Text)) * 18 * CommonUtil.StringToIntVal(txtInput25.Text));
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail1.Text, txtInput1.Text)) * 18 * CommonUtil.StringToIntVal(txtInput25.Text);
            CDataControl.g_ResultFuture.set도매_수익_가입자관리수수료(  CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail1.Text, txtInput1.Text)) * 18 * CommonUtil.StringToIntVal(txtInput25.Text)).ToString(),  txtInput4.Text));
            CDataControl.g_ResultFutureTotal.set도매_수익_CS관리수수료( CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail2.Text, txtInput1.Text)) * 18 * CommonUtil.StringToIntVal(txtInput25.Text));
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail2.Text, txtInput1.Text)) * 18 * CommonUtil.StringToIntVal(txtInput25.Text);
            CDataControl.g_ResultFuture.set도매_수익_CS관리수수료( CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division(txtDetail2.Text, txtInput1.Text)) * 18 * CommonUtil.StringToIntVal(txtInput25.Text)).ToString(), txtInput4.Text));
            CDataControl.g_ResultFutureTotal.set도매_수익_사업자모델매입에따른추가수익( txtDetail4.Text);
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(txtDetail4.Text);
            CDataControl.g_ResultFuture.set도매_수익_사업자모델매입에따른추가수익( CommonUtil.Division(txtDetail4.Text, txtInput4.Text));
            CDataControl.g_ResultFutureTotal.set도매_수익_유통모델매입에따른추가수익_현금_Volume( (CommonUtil.StringToIntVal(txtDetail5.Text) + CommonUtil.StringToIntVal(txtDetail6.Text)).ToString());
            SumSubBenefitWillTotal += (CommonUtil.StringToIntVal(txtDetail5.Text) + CommonUtil.StringToIntVal(txtDetail6.Text));
            CDataControl.g_ResultFuture.set도매_수익_유통모델매입에따른추가수익_현금_Volume( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail5.Text) + CommonUtil.StringToIntVal(txtDetail6.Text)).ToString(), txtInput4.Text));
            CDataControl.g_ResultFutureTotal.도매_수익_소계 = SumSubBenefitWillTotal;
            CDataControl.g_ResultFuture.도매_수익_소계 = CommonUtil.StringToIntVal(CommonUtil.Division(SumSubBenefitWillTotal.ToString(), txtInput4.Text));

            SumSubCostTotal = 0;
            CDataControl.g_ResultFutureTotal.set도매_비용_대리점투자비용( (CommonUtil.StringToIntVal(txtDetail7.Text) * CommonUtil.StringToIntVal(txtInput2.Text) + (CommonUtil.StringToIntVal(txtDetail8.Text) * CommonUtil.StringToIntVal(txtInput3.Text))));
            SumSubCostTotal += (CommonUtil.StringToIntVal(txtDetail7.Text) * CommonUtil.StringToIntVal(txtInput2.Text) + (CommonUtil.StringToIntVal(txtDetail8.Text) * CommonUtil.StringToIntVal(txtInput3.Text)));
            CDataControl.g_ResultFuture.set도매_비용_대리점투자비용( CommonUtil.Division(txtDetail1.Text, txtInput4.Text));
            CDataControl.g_ResultFutureTotal.set도매_비용_인건비_급여_복리후생비( (CommonUtil.StringToIntVal(txtDetail11.Text) * CommonUtil.StringToIntVal(txtInput11.Text)
                + CommonUtil.StringToIntVal(txtDetail12.Text) * CommonUtil.StringToIntVal(txtInput12.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput13.Text)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(txtDetail11.Text) * CommonUtil.StringToIntVal(txtInput11.Text)
                + CommonUtil.StringToIntVal(txtDetail12.Text) * CommonUtil.StringToIntVal(txtInput12.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput13.Text));
            CDataControl.g_ResultFuture.set도매_비용_인건비_급여_복리후생비( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail11.Text) * CommonUtil.StringToIntVal(txtInput11.Text)
                + CommonUtil.StringToIntVal(txtDetail12.Text) * CommonUtil.StringToIntVal(txtInput12.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput13.Text)).ToString(), txtInput4.Text));
            CDataControl.g_ResultFutureTotal.set도매_비용_임차료( (CommonUtil.StringToIntVal(txtDetail13.Text)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(txtDetail13.Text));
            CDataControl.g_ResultFuture.set도매_비용_임차료( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail13.Text)).ToString(), txtInput4.Text));
            CDataControl.g_ResultFutureTotal.set도매_비용_이자비용( CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail33.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail33.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text);
            CDataControl.g_ResultFuture.set도매_비용_이자비용( CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail33.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text)).ToString(), txtInput4.Text));

            CDataControl.g_ResultFutureTotal.set도매_비용_부가세( CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail34.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail34.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text);
            CDataControl.g_ResultFuture.set도매_비용_부가세( CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail34.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text)).ToString(), txtInput4.Text));
            CDataControl.g_ResultFutureTotal.set도매_비용_법인세( CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail35.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail35.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text);
            CDataControl.g_ResultFuture.set도매_비용_법인세( CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail35.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text)).ToString(), txtInput4.Text));
            CDataControl.g_ResultFutureTotal.set도매_비용_기타판매관리비( CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text);
            CDataControl.g_ResultFuture.set도매_비용_기타판매관리비( CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput4.Text)).ToString(), txtInput4.Text));
            CDataControl.g_ResultFutureTotal.도매_비용_소계 = SumSubCostTotal;
            CDataControl.g_ResultFuture.도매_비용_소계 = CommonUtil.StringToIntVal(CommonUtil.Division(SumSubCostTotal.ToString(), txtInput4.Text));
            CDataControl.g_ResultFutureTotal.도매손익계 = (SumSubBenefitTotal - SumSubCostTotal);
            CDataControl.g_ResultFuture.도매손익계 = CommonUtil.StringToIntVal(CommonUtil.Division((SumSubBenefitTotal - SumSubCostTotal).ToString(), txtInput4.Text));


            //소매 당대리점
            SumSubBenefitTotal = 0;
            CDataControl.g_ResultStoreTotal.set소매_수익_업무취급수수료( txtDetail19.Text);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(txtDetail19.Text);
            CDataControl.g_ResultStore.set소매_수익_업무취급수수료( CommonUtil.Division(txtDetail19.Text, txtInput16.Text));
            CDataControl.g_ResultStoreTotal.set소매_수익_직영매장판매수익( txtDetail20.Text);
            SumSubBenefitTotal += CommonUtil.StringToIntVal(txtDetail20.Text);
            CDataControl.g_ResultStore.set소매_수익_직영매장판매수익( CommonUtil.Division(txtDetail20.Text, txtInput16.Text));
            CDataControl.g_ResultStoreTotal.소매_수익_소계 = CommonUtil.StringToIntVal( SumSubBenefitTotal.ToString());
            CDataControl.g_ResultStore.소매_수익_소계 = CommonUtil.StringToIntVal( CommonUtil.Division(SumSubBenefitTotal.ToString(), txtInput16.Text));

            //당대리점 결과(세부항목별 값 입력 결과) 비용계정
            SumSubCostTotal = 0;
            CDataControl.g_ResultStoreTotal.set소매_비용_인건비_급여_복리후생비( (CommonUtil.StringToIntVal(txtDetail23.Text) * CommonUtil.StringToIntVal(txtInput19.Text)
                + CommonUtil.StringToIntVal(txtDetail24.Text) * CommonUtil.StringToIntVal(txtInput20.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput21.Text)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(txtDetail23.Text) * CommonUtil.StringToIntVal(txtInput19.Text)
                + CommonUtil.StringToIntVal(txtDetail24.Text) * CommonUtil.StringToIntVal(txtInput20.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput21.Text));
            CDataControl.g_ResultStore.set소매_비용_인건비_급여_복리후생비( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail23.Text) * CommonUtil.StringToIntVal(txtInput19.Text)
                + CommonUtil.StringToIntVal(txtDetail24.Text) * CommonUtil.StringToIntVal(txtInput20.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput21.Text)).ToString(), txtInput16.Text));
            CDataControl.g_ResultStoreTotal.set소매_비용_임차료( (CommonUtil.StringToIntVal(txtDetail25.Text)));
            SumSubCostTotal += (CommonUtil.StringToIntVal(txtDetail25.Text));
            CDataControl.g_ResultStore.set소매_비용_임차료( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail25.Text)).ToString(), txtInput16.Text));
            CDataControl.g_ResultStoreTotal.set소매_비용_이자비용( CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail33.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail33.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text);
            CDataControl.g_ResultStore.set소매_비용_이자비용( CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail33.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text)).ToString(), txtInput16.Text));

            CDataControl.g_ResultStoreTotal.set소매_비용_부가세( CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail34.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail34.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text);
            CDataControl.g_ResultStore.set소매_비용_부가세( CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail34.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text)).ToString(), txtInput16.Text));
            CDataControl.g_ResultStoreTotal.set소매_비용_법인세( CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail35.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail35.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text);
            CDataControl.g_ResultStore.set소매_비용_법인세(CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail35.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text)).ToString(), txtInput16.Text));
            CDataControl.g_ResultStoreTotal.set소매_비용_기타판매관리비( CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text));
            SumSubCostTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text);
            CDataControl.g_ResultStore.set소매_비용_기타판매관리비( CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text)).ToString(), txtInput16.Text));
            CDataControl.g_ResultStoreTotal.소매_비용_소계  =  SumSubCostTotal;
            CDataControl.g_ResultStore.소매_비용_소계 = CommonUtil.StringToIntVal(CommonUtil.Division(SumSubCostTotal.ToString(), txtInput4.Text));
            CDataControl.g_ResultStoreTotal.소매손익계 = (SumSubBenefitTotal - SumSubCostTotal);
            CDataControl.g_ResultStore.소매손익계 = CommonUtil.StringToIntVal(CommonUtil.Division((SumSubBenefitTotal - SumSubCostTotal).ToString(), txtInput4.Text));

            CDataControl.g_ResultStoreTotal.점별손익추정 = CommonUtil.StringToIntVal(CommonUtil.Division((SumSubBenefitTotal - SumSubCostTotal).ToString(), txtInput30.Text));
            CDataControl.g_ResultStore.점별손익추정 = CommonUtil.StringToIntVal( txtInput30.Text);


            //소매 당대리점미래
            SumSubBenefitWillTotal = 0;
            CDataControl.g_ResultFutureTotal.set소매_수익_업무취급수수료( txtDetail19.Text);
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(txtDetail19.Text);
            CDataControl.g_ResultFuture.set소매_수익_업무취급수수료( CommonUtil.Division(txtDetail19.Text, txtInput16.Text));
            CDataControl.g_ResultFutureTotal.set소매_수익_직영매장판매수익( txtDetail20.Text);
            SumSubBenefitWillTotal += CommonUtil.StringToIntVal(txtDetail20.Text);
            CDataControl.g_ResultFuture.set소매_수익_직영매장판매수익( CommonUtil.Division(txtDetail20.Text, txtInput16.Text));
            CDataControl.g_ResultFutureTotal.소매_수익_소계 =  SumSubBenefitWillTotal;
            CDataControl.g_ResultFuture.소매_수익_소계 = CommonUtil.StringToIntVal(CommonUtil.Division(SumSubBenefitWillTotal.ToString(), txtInput16.Text));

            //당대리점 결과(세부항목별 값 입력 결과) 비용계정
            SumSubCostWillTotal = 0;
            CDataControl.g_ResultFutureTotal.set소매_비용_인건비_급여_복리후생비( (CommonUtil.StringToIntVal(txtDetail23.Text) * CommonUtil.StringToIntVal(txtInput19.Text)
                + CommonUtil.StringToIntVal(txtDetail24.Text) * CommonUtil.StringToIntVal(txtInput20.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput21.Text)));
            SumSubCostWillTotal += (CommonUtil.StringToIntVal(txtDetail23.Text) * CommonUtil.StringToIntVal(txtInput19.Text)
                + CommonUtil.StringToIntVal(txtDetail24.Text) * CommonUtil.StringToIntVal(txtInput20.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput21.Text));
            CDataControl.g_ResultFuture.set소매_비용_인건비_급여_복리후생비( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail23.Text) * CommonUtil.StringToIntVal(txtInput19.Text)
                + CommonUtil.StringToIntVal(txtDetail24.Text) * CommonUtil.StringToIntVal(txtInput20.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) * CommonUtil.StringToIntVal(txtInput21.Text)).ToString(), txtInput16.Text));
            CDataControl.g_ResultFutureTotal.set소매_비용_임차료( (CommonUtil.StringToIntVal(txtDetail25.Text)));
            SumSubCostWillTotal += (CommonUtil.StringToIntVal(txtDetail25.Text));
            CDataControl.g_ResultFuture.set소매_비용_임차료( CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail25.Text)).ToString(), txtInput16.Text));
            CDataControl.g_ResultFutureTotal.set소매_비용_이자비용( CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail33.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text));
            SumSubCostWillTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail33.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text);
            CDataControl.g_ResultFuture.set소매_비용_이자비용( CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail33.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text)).ToString(), txtInput16.Text));

            CDataControl.g_ResultFutureTotal.set소매_비용_부가세( CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail34.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text));
            SumSubCostWillTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail34.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text);
            CDataControl.g_ResultFuture.set소매_비용_부가세( CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail34.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text)).ToString(), txtInput16.Text));
            CDataControl.g_ResultFutureTotal.set소매_비용_법인세( CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail35.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text));
            SumSubCostWillTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail35.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text);
            CDataControl.g_ResultFuture.set소매_비용_법인세( CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail35.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text)).ToString(), txtInput16.Text));
            CDataControl.g_ResultFutureTotal.set소매_비용_기타판매관리비( CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text));
            SumSubCostWillTotal += CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text);
            CDataControl.g_ResultFuture.set소매_비용_기타판매관리비( CommonUtil.Division((CommonUtil.StringToIntVal(CommonUtil.Division((CommonUtil.StringToIntVal(txtDetail14.Text) + CommonUtil.StringToIntVal(txtDetail15.Text)
                + CommonUtil.StringToIntVal(txtDetail16.Text) + CommonUtil.StringToIntVal(txtDetail17.Text)
                + CommonUtil.StringToIntVal(txtDetail18.Text) + CommonUtil.StringToIntVal(txtDetail26.Text)
                + CommonUtil.StringToIntVal(txtDetail27.Text) + CommonUtil.StringToIntVal(txtDetail28.Text)
                + CommonUtil.StringToIntVal(txtDetail29.Text) + CommonUtil.StringToIntVal(txtDetail30.Text)
                + CommonUtil.StringToIntVal(txtDetail31.Text) + CommonUtil.StringToIntVal(txtDetail32.Text)
                + CommonUtil.StringToIntVal(txtDetail36.Text)).ToString(), txtInput25.Text)) * CommonUtil.StringToIntVal(txtInput16.Text)).ToString(), txtInput16.Text));
            CDataControl.g_ResultFutureTotal.소매_비용_소계 = SumSubCostWillTotal;
            CDataControl.g_ResultFuture.소매_비용_소계 = CommonUtil.StringToIntVal(CommonUtil.Division(SumSubCostWillTotal.ToString(), txtInput4.Text));
            CDataControl.g_ResultFutureTotal.소매손익계 =  (SumSubBenefitWillTotal - SumSubCostWillTotal);
            CDataControl.g_ResultFuture.소매손익계 = CommonUtil.StringToIntVal( CommonUtil.Division((SumSubBenefitWillTotal - SumSubCostWillTotal).ToString(), txtInput4.Text));

            CDataControl.g_ResultFutureTotal.점별손익추정 = CommonUtil.StringToIntVal( CommonUtil.Division((SumSubBenefitWillTotal - SumSubCostWillTotal).ToString(), txtInput30.Text));
            CDataControl.g_ResultFuture.점별손익추정 = CommonUtil.StringToIntVal( txtInput30.Text);


        }


        //도매입력시 합계치 변경
        //누적가입자수
        private void txtInput1_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput1.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput1.SelectionStart = txtInput1.Text.Length;
            }
            txtInput22.Text = txtInput1.Text;
        }
        //월평균 판매대수 신규
        private void txtInput2_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput2.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput2.SelectionStart = txtInput2.Text.Length;
            }


            txtInput4.Text = CommonUtil.Sum_Values(txtInput2.Text, txtInput3.Text);
            txtInput23.Text = CommonUtil.Sum_Values(txtInput2.Text, txtInput14.Text);
        }
        //월평균 판매대수 기변
        private void txtInput3_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput3.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput3.SelectionStart = txtInput3.Text.Length;
            }

            txtInput4.Text = CommonUtil.Sum_Values(txtInput2.Text, txtInput3.Text);
            txtInput24.Text = CommonUtil.Sum_Values(txtInput3.Text, txtInput15.Text);
        }
        //월평균 판매대수 계
        private void txtInput4_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput4.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput4.SelectionStart = txtInput4.Text.Length;
            }

            txtInput25.Text = CommonUtil.Sum_Values(txtInput4.Text, txtInput16.Text);
        }
        //월평균 유통모델 출고대수 LG
        private void txtInput5_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput5.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput5.SelectionStart = txtInput5.Text.Length;
            }


            txtInput7.Text = CommonUtil.Sum_Values(txtInput5.Text, txtInput6.Text);
            txtInput26.Text = txtInput5.Text;
        }
        //월평균 유통모델 출고대수 SS
        private void txtInput6_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput6.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput6.SelectionStart = txtInput6.Text.Length;
            }

            txtInput7.Text = CommonUtil.Sum_Values(txtInput5.Text, txtInput6.Text);
            txtInput27.Text = txtInput6.Text;
        }
        //월평균 유통모델 출고대수 계
        private void txtInput7_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput7.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput7.SelectionStart = txtInput7.Text.Length;
            }

            txtInput28.Text = txtInput7.Text;
        }

        //거래선 수 개통사무실
        private void txtInput8_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput8.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput8.SelectionStart = txtInput8.Text.Length;
            }

            txtInput10.Text = CommonUtil.Sum_Values(txtInput8.Text, txtInput9.Text);
            txtInput29.Text = txtInput8.Text;
        }
        //거래선 수 판매점
        private void txtInput9_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput9.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput9.SelectionStart = txtInput9.Text.Length;
            }

            txtInput10.Text = CommonUtil.Sum_Values(txtInput8.Text, txtInput9.Text);
            txtInput31.Text = txtInput9.Text;
        }
        //거래선 수 계
        private void txtInput10_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput10.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput10.SelectionStart = txtInput10.Text.Length;
            }

            txtInput32.Text = CommonUtil.Sum_Values(txtInput10.Text, txtInput18.Text);
        }

        //직원수 간부급
        private void txtInput11_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput11.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput11.SelectionStart = txtInput11.Text.Length;
            }

            txtInput13.Text = CommonUtil.Sum_Values(txtInput11.Text, txtInput12.Text);
            txtInput33.Text = CommonUtil.Sum_Values(txtInput11.Text, txtInput19.Text);
        }
        //직원수 평사원
        private void txtInput12_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput12.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput12.SelectionStart = txtInput12.Text.Length;
            }

            txtInput13.Text = CommonUtil.Sum_Values(txtInput11.Text, txtInput12.Text);
            txtInput34.Text = CommonUtil.Sum_Values(txtInput12.Text, txtInput20.Text);
        }
        //직원수 계
        private void txtInput13_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput13.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput13.SelectionStart = txtInput13.Text.Length;
            }

            txtInput35.Text = CommonUtil.Sum_Values(txtInput13.Text, txtInput21.Text);
        }

        //소매입력시 합계치 변경
        //월판매대수 신규
        private void txtInput14_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput14.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput14.SelectionStart = txtInput14.Text.Length;
            }

            txtInput16.Text = CommonUtil.Sum_Values(txtInput14.Text, txtInput15.Text);
            txtInput23.Text = CommonUtil.Sum_Values(txtInput2.Text, txtInput14.Text);
        }
        //월판매대수 신규
        private void txtInput15_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput15.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput15.SelectionStart = txtInput15.Text.Length;
            }

            txtInput16.Text = CommonUtil.Sum_Values(txtInput14.Text, txtInput15.Text);
            txtInput24.Text = CommonUtil.Sum_Values(txtInput3.Text, txtInput15.Text);
        }
        //월판매대수 신규
        private void txtInput16_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput16.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput16.SelectionStart = txtInput16.Text.Length;
            }

            txtInput25.Text = CommonUtil.Sum_Values(txtInput4.Text, txtInput16.Text);
        }

        //거래선 수 직영점
        private void txtInput17_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput17.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput17.SelectionStart = txtInput17.Text.Length;
            }

            txtInput18.Text = txtInput17.Text;
            txtInput30.Text = txtInput17.Text;
        }
        //거래선 수 계
        private void txtInput18_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput18.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput18.SelectionStart = txtInput18.Text.Length;
            }

            txtInput32.Text = CommonUtil.Sum_Values(txtInput10.Text, txtInput18.Text);
        }
        //직원수 간부급
        private void txtInput19_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput19.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput19.SelectionStart = txtInput19.Text.Length;
            }

            txtInput21.Text = CommonUtil.Sum_Values(txtInput19.Text, txtInput20.Text);
            txtInput33.Text = CommonUtil.Sum_Values(txtInput11.Text, txtInput19.Text);
        }
        //직원수 평사원
        private void txtInput20_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput20.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput20.SelectionStart = txtInput20.Text.Length;
            }

            txtInput21.Text = CommonUtil.Sum_Values(txtInput19.Text, txtInput20.Text);
            txtInput34.Text = CommonUtil.Sum_Values(txtInput12.Text, txtInput20.Text);
        }
        //직원수 계
        private void txtInput21_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput21.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput21.SelectionStart = txtInput21.Text.Length;
            }

            txtInput35.Text = CommonUtil.Sum_Values(txtInput13.Text, txtInput21.Text);
        }



        private void txtInput23_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput23.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput23.SelectionStart = txtInput23.Text.Length;
            }
        }

        private void txtInput24_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput24.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput24.SelectionStart = txtInput24.Text.Length;
            }
        }

        private void txtInput25_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput25.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput25.SelectionStart = txtInput25.Text.Length;
            }
        }

        private void txtInput32_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput32.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput32.SelectionStart = txtInput32.Text.Length;
            }
        }

        private void txtInput33_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput33.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput33.SelectionStart = txtInput33.Text.Length;
            }
        }

        private void txtInput34_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput34.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput34.SelectionStart = txtInput34.Text.Length;
            }
        }

        private void txtInput35_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtInput35.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtInput35.SelectionStart = txtInput35.Text.Length;
            }
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
                string filename = CommonUtil.defaultName;


                FileInfo fi2 = new FileInfo(filename);
                    fi2.CopyTo(saveFileDialog1.FileName, true);

                CommonUtil.saveAsName = saveFileDialog1.FileName;

                //excel.Workbook _Workbook = CommonUtil.GetExcel_WorkBook(saveFileDialog1.FileName);
                //excel.Worksheet _WorkSheet1 = _Workbook.Sheets[1] as excel.Worksheet;
                //excel.Worksheet _WorkSheet2 = _Workbook.Sheets[2] as excel.Worksheet;
                SaveAsInput();
                CommonUtil.WriteDataToExcelFile(CommonUtil.saveAsName, CDataControl.g_BasicInput, CDataControl.g_DetailInput);
            }
        }



        //CS 관리 수수료 처리용 시작
        private void txtDetail2_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail2.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail2.SelectionStart = txtDetail2.Text.Length;
            }

            if (radioButton5.Checked)
            {
                txtDetail3.Text = (CommonUtil.StringToIntVal(txtDetail2.Text.Replace(",","")) * CommonUtil.QUARTER).ToString();
            }
        }

        private void txtDetail3_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail3.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail3.SelectionStart = txtDetail3.Text.Length;
            }

            if (radioButton6.Checked)
            {
                txtDetail2.Text = CommonUtil.Division(txtDetail3.Text.Replace(",", ""), CommonUtil.QUARTER.ToString());
            }
        }

        private void txtInput1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && e.KeyChar != 8)
            {
                e.Handled = true;
            }
 
        }

        
        //CS 관리 수수료 처리용 끝
        //도매 직원급여 총액 처리용 시작
        private void txtDetail9_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail9.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail9.SelectionStart = txtDetail9.Text.Length;
            }

            if (radioButton1.Checked)
            {
                txtDetail11.Text = CommonUtil.Division(txtDetail9.Text.Replace(",", ""), txtInput11.Text.Replace(",", ""));
            }
        }

        private void txtDetail10_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail10.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail10.SelectionStart = txtDetail10.Text.Length;
            }

            if (radioButton1.Checked)
            {
                txtDetail12.Text = CommonUtil.Division(txtDetail10.Text.Replace(",", ""), txtInput12.Text.Replace(",", ""));
            }
        }

        private void txtDetail11_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail11.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail11.SelectionStart = txtDetail11.Text.Length;
            }

            if (radioButton2.Checked)
            {
                txtDetail9.Text = (CommonUtil.StringToIntVal(txtDetail11.Text.Replace(",", "")) * CommonUtil.StringToIntVal(txtInput11.Text.Replace(",", ""))).ToString();
            }

        }

        private void txtDetail12_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail12.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail12.SelectionStart = txtDetail12.Text.Length;
            }

            if (radioButton2.Checked)
            {
                txtDetail10.Text = (CommonUtil.StringToIntVal(txtDetail12.Text.Replace(",", "")) * CommonUtil.StringToIntVal(txtInput12.Text.Replace(",", ""))).ToString();
            }

        }
        //도매 직원급여 총액 처리용 끝

        //소매 직원급여 총액 처리용 시작
        private void txtDetail21_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail21.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail21.SelectionStart = txtDetail21.Text.Length;
            }

            if (radioButton4.Checked)
            {
                txtDetail23.Text = CommonUtil.Division(txtDetail21.Text.Replace(",", ""), txtInput19.Text.Replace(",", ""));
            }

        }

        private void txtDetail22_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail22.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail22.SelectionStart = txtDetail22.Text.Length;
            }
            
            if (radioButton4.Checked)
            {
                txtDetail24.Text = CommonUtil.Division(txtDetail22.Text.Replace(",", ""), txtInput20.Text.Replace(",", ""));
            }

        }

        private void txtDetail23_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail23.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail23.SelectionStart = txtDetail23.Text.Length;
            }
            
            if (radioButton3.Checked)
            {
                txtDetail21.Text = (CommonUtil.StringToIntVal(txtDetail23.Text.Replace(",", "")) * CommonUtil.StringToIntVal(txtInput19.Text.Replace(",", ""))).ToString();
            }

        }

        private void txtDetail24_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail24.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail24.SelectionStart = txtDetail24.Text.Length;
            }
            
            if (radioButton3.Checked)
            {
                txtDetail22.Text = (CommonUtil.StringToIntVal(txtDetail24.Text.Replace(",", "")) * CommonUtil.StringToIntVal(txtInput20.Text.Replace(",", ""))).ToString();
            }

        }




        //소매 직원급여 총액 처리용 끝


        private void txtDetail1_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail1.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail1.SelectionStart = txtDetail1.Text.Length;
            }
        }

        private void txtDetail4_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail4.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail4.SelectionStart = txtDetail4.Text.Length;
            }

        }

        private void txtDetail5_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail5.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail5.SelectionStart = txtDetail5.Text.Length;
            }

        }

        private void txtDetail6_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail6.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail6.SelectionStart = txtDetail6.Text.Length;
            }

        }

        private void txtDetail7_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail7.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail7.SelectionStart = txtDetail7.Text.Length;
            }

        }

        private void txtDetail8_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail8.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail8.SelectionStart = txtDetail8.Text.Length;
            }

        }

        private void txtDetail13_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail13.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail13.SelectionStart = txtDetail13.Text.Length;
            }

        }

        private void txtDetail14_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail14.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail14.SelectionStart = txtDetail14.Text.Length;
            }

        }

        private void txtDetail15_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail15.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail15.SelectionStart = txtDetail15.Text.Length;
            }

        }

        private void txtDetail16_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail16.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail16.SelectionStart = txtDetail16.Text.Length;
            }

        }

        private void txtDetail17_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail17.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail17.SelectionStart = txtDetail17.Text.Length;
            }

        }

        private void txtDetail18_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail18.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail18.SelectionStart = txtDetail18.Text.Length;
            }

        }

        private void txtDetail19_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail19.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail19.SelectionStart = txtDetail19.Text.Length;
            }

        }

        private void txtDetail20_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail20.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail20.SelectionStart = txtDetail20.Text.Length;
            }

        }

        private void txtDetail25_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail25.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail25.SelectionStart = txtDetail25.Text.Length;
            }

        }

        private void txtDetail26_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail26.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail26.SelectionStart = txtDetail26.Text.Length;
            }

        }

        private void txtDetail27_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail27.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail27.SelectionStart = txtDetail27.Text.Length;
            }

        }

        private void txtDetail28_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail28.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail28.SelectionStart = txtDetail28.Text.Length;
            }

        }

        private void txtDetail29_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail29.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail29.SelectionStart = txtDetail29.Text.Length;
            }

        }

        private void txtDetail30_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail30.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail30.SelectionStart = txtDetail30.Text.Length;
            }

        }

        private void txtDetail31_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail31.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail31.SelectionStart = txtDetail31.Text.Length;
            }

        }

        private void txtDetail32_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail32.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail32.SelectionStart = txtDetail32.Text.Length;
            }

        }

        private void txtDetail33_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail33.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail33.SelectionStart = txtDetail33.Text.Length;
            }

        }

        private void txtDetail34_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail34.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail34.SelectionStart = txtDetail34.Text.Length;
            }

        }

        private void txtDetail35_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail35.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail35.SelectionStart = txtDetail35.Text.Length;
            }

        }

        private void txtDetail36_TextChanged(object sender, EventArgs e)
        {
            if ((sender as TextBox).Text.Contains(",") || (sender as TextBox).Text.Length > 0)
            {
                txtDetail36.Text = String.Format("{0:#,###}", Convert.ToInt64((sender as TextBox).Text.Replace(",", "")));
                txtDetail36.SelectionStart = txtDetail36.Text.Length;
            }

        }


        
    }
}
