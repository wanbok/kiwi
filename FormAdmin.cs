
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using excel = Microsoft.Office.Interop.Excel;

namespace KIWI
{

    public partial class FormAdmin : Form
    {
        public const string CHECK = "√";

        private CAdminDataController adminDC = null;
        private ListViewColumnSorter lvwColumnSorter = null;

        private TextBox[] txtOut = null;        //기존 업계평균
        private TextBox[] txtIAOut = null;      // 업계평균
        private TextBox[] txtInput = null;      //보정 계수
        private TextBox[] txtAOut = null;       //보정 계수 업계평균

        private Int64[] nOut = null;
        private Int64[] nIAOut = null;
        private Int64[] nInput = null;
        private Int64[] nAOut = null;

        public FormAdmin()
        {
            InitializeComponent();
            adminDC = new CAdminDataController();
            lvwColumnSorter = new ListViewColumnSorter();
            this.listView1.ListViewItemSorter = lvwColumnSorter;

            txtOut = new TextBox[31] { txtOut1, txtOut2, txtOut3, txtOut4, txtOut5, txtOut6, txtOut7, txtOut8, txtOut9, txtOut10,
            txtOut11, txtOut12, txtOut13, txtOut14, txtOut15, txtOut16, txtOut17, txtOut18, txtOut19, txtOut20,
            txtOut21, txtOut22, txtOut23, txtOut24, txtOut25, txtOut26, txtOut27, txtOut28, txtOut29, txtOut30,
            txtOut31
            };

            txtIAOut = new TextBox[31] { txtOut32, txtOut33, txtOut34, txtOut35, txtOut36, txtOut37, txtOut38, txtOut39, txtOut40,
            txtOut41, txtOut42, txtOut43, txtOut44, txtOut45, txtOut46, txtOut47, txtOut48, txtOut49, txtOut50,
            txtOut51, txtOut52, txtOut53, txtOut54, txtOut55, txtOut56, txtOut57, txtOut58, txtOut59, txtOut60,
            txtOut61, txtOut62 
            };

            txtInput = new TextBox[31] { txtInput1, txtInput2, txtInput3, txtInput4, txtInput5, txtInput6, txtInput7, txtInput8, txtInput9, txtInput10,
            txtInput11, txtInput12, txtInput13, txtInput14, txtInput15, txtInput16, txtInput17, txtInput18, txtInput19, txtInput20,
            txtInput21, txtInput22, txtInput23, txtInput24, txtInput25, txtInput26, txtInput27, txtInput28, txtInput29, txtInput30,
            txtInput31
            };

            txtAOut = new TextBox[31] { txtAOut1, txtAOut2, txtAOut3, txtAOut4, txtAOut5, txtAOut6, txtAOut7, txtAOut8, txtAOut9, txtAOut10,
            txtAOut11, txtAOut12, txtAOut13, txtAOut14, txtAOut15, txtAOut16, txtAOut17, txtAOut18, txtAOut19, txtAOut20,
            txtAOut21, txtAOut22, txtAOut23, txtAOut24, txtAOut25, txtAOut26, txtAOut27, txtAOut28, txtAOut29, txtAOut30,
            txtAOut31
            };

            nOut = new Int64[31];
            nIAOut = new Int64[31];
            nInput = new Int64[31];
            nAOut = new Int64[31];

            readFileOfExistedAverage();

            refreshList();
        }



        //기존 업계평균
        private void getBeforeIndustryAverage()
        {

        }
        private void setBeforeIndustryAverage()
        {
            for (int i = 0; i < 31; i++)
                txtOut[i].Text = i.ToString();
        }

        //업계평균
        private void getIndustryAverage()
        {

        }
        private void setIndustryAverage()
        {
            for (int i = 0; i < 31; i++)
                txtIAOut[i].Text = i.ToString();
        }

        //보정계수
        private void getCorrectionFactor()
        {

        }
        private void setCorrectionFactor()
        {
            for (int i = 0; i < 31; i++)
                txtInput[i].Text = "";
        }

        //보정계수 평균
        private void getCorrectionFactorIndustryAverage()
        {

        }

        private void setCorrectionFactorIndustryAverage()
        {
            for (int i = 0; i < 31; i++)
                txtAOut[i].Text = (Convert.ToInt32(txtIAOut[i].Text) + Convert.ToInt32(txtInput[i].Text) / 2).ToString();
        }


        //업계평균과 보정계수의 평균을 구하고 보정계수 평균에 적용. 
        private void button2_Click(object sender, EventArgs e)
        {
            saveFileOfExistedAverage(txtAOut);
        }

        // 파일열기
        public void openFileDialog(object sender, EventArgs e)
        {
            // Displays an OpenFileDialog so the user can select a Cursor.
            // OpenFileDialog openFileDialog1 = new OpenFileDialog();
            //openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "Excel File|*.xlsx";
            openFileDialog1.Title = "Select a Excel File";
            openFileDialog1.RestoreDirectory = true;

            // Show the Dialog.
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Read the files
                foreach (String file in openFileDialog1.FileNames)
                {
                    try
                    {
                        setExcelFileToXML(file);
                    }
                    catch (Exception ex)
                    {
                        // Could not load the image - probably related to Windows file system permissions.
                        MessageBox.Show("파일을 열 수 없습니다.\n\nReported error: " + ex.Message);
                    }
                }
                refreshList();

            }
        }

        private void refreshList()
        {
            String key = null;
            String 지역 = null;
            String 대리점명 = null;
            String 마케터 = null;
            String 단위당손익 = null;
            String 월capa = null;
            String 가입자수 = null;
            String 직영점판매수익 = null;
            String 선택여부 = null;
            String mExcelFileName = null;
            CBasicInput mBI = null;
            CBusinessData mDI = null;
            CResultData mRD = null;

            listView1.Items.Clear();
            int indexForListViewId = 0;
            for (int i = 0; i < adminDC.getDataLength(); i++)
            {
                adminDC.GetData(i, out key, out 지역, out 대리점명, out 마케터, out 단위당손익, out 월capa, out 가입자수, out 직영점판매수익, out 선택여부, out mExcelFileName, out mBI, out mDI, out mRD);
                ListViewItem item = new ListViewItem();
                item.Tag = key;
                item.SubItems[0].Text = (++indexForListViewId).ToString();
                item.SubItems.Add(지역);
                item.SubItems.Add(대리점명);
                item.SubItems.Add(마케터);
                item.SubItems.Add(단위당손익);
                item.SubItems.Add(월capa);
                item.SubItems.Add(가입자수);
                item.SubItems.Add(직영점판매수익);
                if (선택여부 == "Y")
                {
                    item.SubItems.Add(CHECK);
                    item.ForeColor = Color.Green;
                }
                else
                {
                    item.SubItems.Add("");
                    item.ForeColor = Color.Red;
                }
                listView1.Items.Add(item);
            }

            for (int i = 0; i < nIAOut.Length; i++)
            {
                nIAOut[i] = 0;
            }

            for (int i = 0; i < adminDC.getDataLength(); i++)
            {
                adminDC.GetData(i, out key, out 지역, out 대리점명, out 마케터, out 단위당손익, out 월capa, out 가입자수, out 직영점판매수익, out 선택여부, out mExcelFileName, out mBI, out mDI, out mRD);

                if (선택여부 == "N") continue;
                int k = 0;
                if (mBI.get도매_월평균판매대수_소계() == 0)
                {
                    nIAOut[k++] += mDI.get도매_수익_월평균관리수수료();
                    nIAOut[k++] += mDI.get도매_수익_CS관리수수료();
                    nIAOut[k++] += mDI.get도매_수익_사업자모델매입관련추가수익();
                    nIAOut[k++] += mDI.get도매_수익_유통모델매입관련추가수익_현금DC();
                    nIAOut[k++] += mDI.get도매_수익_유통모델매입관련추가수익_VolumeDC();
                    nIAOut[k++] += mDI.get도매_비용_대리점투자금액_신규();
                    nIAOut[k++] += mDI.get도매_비용_대리점투자금액_기변();
                    nIAOut[k++] += mDI.get도매_비용_직원급여_간부급_총액(mBI.get도매_직원수_간부급());
                    nIAOut[k++] += mDI.get도매_비용_직원급여_평사원_총액(mBI.get도매_직원수_평사원());
                    nIAOut[k++] += mDI.get도매_비용_지급임차료();
                    nIAOut[k++] += mDI.get도매_비용_운반비();
                    nIAOut[k++] += mDI.get도매_비용_차량유지비();
                    nIAOut[k++] += mDI.get도매_비용_지급수수료();
                    nIAOut[k++] += mDI.get도매_비용_건물관리비();
                }
                else
                {
                    nIAOut[k++] += mDI.get도매_수익_월평균관리수수료() / mBI.get도매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get도매_수익_CS관리수수료() / mBI.get도매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get도매_수익_사업자모델매입관련추가수익() / mBI.get도매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get도매_수익_유통모델매입관련추가수익_현금DC() / mBI.get도매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get도매_수익_유통모델매입관련추가수익_VolumeDC() / mBI.get도매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get도매_비용_대리점투자금액_신규() / mBI.get도매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get도매_비용_대리점투자금액_기변() / mBI.get도매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get도매_비용_직원급여_간부급_총액(mBI.get도매_직원수_간부급()) / mBI.get도매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get도매_비용_직원급여_평사원_총액(mBI.get도매_직원수_평사원()) / mBI.get도매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get도매_비용_지급임차료() / mBI.get도매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get도매_비용_운반비() / mBI.get도매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get도매_비용_차량유지비() / mBI.get도매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get도매_비용_지급수수료() / mBI.get도매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get도매_비용_건물관리비() / mBI.get도매_월평균판매대수_소계();
                }
                if (mBI.get소매_월평균판매대수_소계() == 0)
                {
                    nIAOut[k++] += mDI.get소매_수익_월평균업무취급수수료();
                    nIAOut[k++] += mDI.get소매_수익_직영매장판매수익();
                    nIAOut[k++] += mDI.get소매_비용_직원급여_간부급_총액(mBI.get소매_직원수_간부급());
                    nIAOut[k++] += mDI.get소매_비용_직원급여_평사원_총액(mBI.get소매_직원수_평사원());
                    nIAOut[k++] += mDI.get소매_비용_지급임차료();
                    nIAOut[k++] += mDI.get소매_비용_지급수수료();
                    nIAOut[k++] += mDI.get소매_비용_판매촉진비();
                    nIAOut[k++] += mDI.get소매_비용_건물관리비();
                }
                else
                {
                    nIAOut[k++] += mDI.get소매_수익_월평균업무취급수수료() / mBI.get소매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get소매_수익_직영매장판매수익() / mBI.get소매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get소매_비용_직원급여_간부급_총액(mBI.get소매_직원수_간부급()) / mBI.get소매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get소매_비용_직원급여_평사원_총액(mBI.get소매_직원수_평사원()) / mBI.get소매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get소매_비용_지급임차료() / mBI.get소매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get소매_비용_지급수수료() / mBI.get소매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get소매_비용_판매촉진비() / mBI.get소매_월평균판매대수_소계();
                    nIAOut[k++] += mDI.get소매_비용_건물관리비() / mBI.get소매_월평균판매대수_소계();
                }
                if (mBI.get월평균판매대수_소계_합계() == 0)
                {
                    nIAOut[k++] += mDI.get도소매_비용_복리후생비();
                    nIAOut[k++] += mDI.get도소매_비용_통신비();
                    nIAOut[k++] += mDI.get도소매_비용_공과금();
                    nIAOut[k++] += mDI.get도소매_비용_소모품비();
                    nIAOut[k++] += mDI.get도소매_비용_이자비용();
                    nIAOut[k++] += mDI.get도소매_비용_부가세();
                    nIAOut[k++] += mDI.get도소매_비용_법인세();
                    nIAOut[k++] += mDI.get도소매_비용_기타();
                }
                else
                {
                    nIAOut[k++] += mDI.get도소매_비용_복리후생비() / mBI.get월평균판매대수_소계_합계();
                    nIAOut[k++] += mDI.get도소매_비용_통신비() / mBI.get월평균판매대수_소계_합계();
                    nIAOut[k++] += mDI.get도소매_비용_공과금() / mBI.get월평균판매대수_소계_합계();
                    nIAOut[k++] += mDI.get도소매_비용_소모품비() / mBI.get월평균판매대수_소계_합계();
                    nIAOut[k++] += mDI.get도소매_비용_이자비용() / mBI.get월평균판매대수_소계_합계();
                    nIAOut[k++] += mDI.get도소매_비용_부가세() / mBI.get월평균판매대수_소계_합계();
                    nIAOut[k++] += mDI.get도소매_비용_법인세() / mBI.get월평균판매대수_소계_합계();
                    nIAOut[k++] += mDI.get도소매_비용_기타() / mBI.get월평균판매대수_소계_합계();
                }
            }

            Int64 분모 = adminDC.getDataLength() > 0 ? adminDC.getDataLength() : 1;
            for (int i = 0; i < nIAOut.Length; i++)
            {
                txtIAOut[i].Text = (nIAOut[i] / 분모).ToString();
            }

        }

        private void setExcelFileToXML(string file)
        {
            excel.Worksheet worksheet1 = CommonUtil.GetExcelWorksheet(file, 1);
            excel.Worksheet worksheet2 = CommonUtil.GetExcelWorksheet(file, 2);

            CBasicInput mBI = null;
            CBusinessData mDI = null;
            CResultData mRD = null;

            setDataForUse(worksheet1, worksheet2, out mBI, out mDI, out mRD);

            CommonUtil.GetExcel_WorkBook_CLOSE();

            adminDC.AddSaveData(
                mBI.get지역(),
                mBI.get대리점(),
                mBI.get마케터(),
                mDI.get소매_수익_직영매장판매수익().ToString(),
                mBI.get월평균판매대수_소계_합계().ToString(),
                mBI.get누적가입자수_합계().ToString(),
                mBI.get거래선수_직영점_합계().ToString(),
                "Y",
                file,
                mBI, mDI, mRD
            );

        }

        private void setDataForUse(excel.Worksheet worksheet1, excel.Worksheet worksheet2, out CBasicInput mBI, out CBusinessData mDI, out CResultData mRD)
        {
            mBI = new CBasicInput();
            mDI = new CBusinessData();
            mRD = new CResultData();

            if (worksheet1 == null || worksheet2 == null || mBI == null || mDI == null || mRD == null) return;

            //*******CBasicInput
            mBI.set지역(worksheet1.get_Range("C63", Type.Missing).Value2.ToString());
            mBI.set대리점(worksheet1.get_Range("E63", Type.Missing).Value2.ToString());
            mBI.set마케터(worksheet1.get_Range("G63", Type.Missing).Value2.ToString());

            //도매
            mBI.set도매_누적가입자수(worksheet1.get_Range("F7", Type.Missing).Value2.ToString());

            mBI.set도매_월평균판매대수_신규(worksheet1.get_Range("F8", Type.Missing).Value2.ToString());
            mBI.set도매_월평균판매대수_기변(worksheet1.get_Range("F9", Type.Missing).Value2.ToString());

            mBI.set도매_월평균유통모델출고대수_LG(worksheet1.get_Range("F11", Type.Missing).Value2.ToString());
            mBI.set도매_월평균유통모델출고대수_SS(worksheet1.get_Range("F12", Type.Missing).Value2.ToString());

            mBI.set도매_거래선수_개통사무실(worksheet1.get_Range("F14", Type.Missing).Value2.ToString());
            mBI.set도매_거래선수_판매점(worksheet1.get_Range("F16", Type.Missing).Value2.ToString());

            mBI.set도매_직원수_간부급(worksheet1.get_Range("F18", Type.Missing).Value2.ToString());
            mBI.set도매_직원수_평사원(worksheet1.get_Range("F19", Type.Missing).Value2.ToString());

            //소매
            mBI.set소매_월평균판매대수_신규(worksheet1.get_Range("G8", Type.Missing).Value2.ToString());
            mBI.set소매_월평균판매대수_기변(worksheet1.get_Range("G9", Type.Missing).Value2.ToString());

            mBI.set소매_거래선수_직영점(worksheet1.get_Range("G15", Type.Missing).Value2.ToString());

            mBI.set소매_직원수_간부급(worksheet1.get_Range("G18", Type.Missing).Value2.ToString());
            mBI.set소매_직원수_평사원(worksheet1.get_Range("G19", Type.Missing).Value2.ToString());

            //*******CBusinessData
            //도매
            mDI.set도매_수익_월평균관리수수료(worksheet1.get_Range("G26", Type.Missing).Value2.ToString());
            mDI.set도매_수익_CS관리수수료(worksheet1.get_Range("G27", Type.Missing).Value2.ToString());//월총액
            mDI.set도매_수익_사업자모델매입관련추가수익(worksheet1.get_Range("G29", Type.Missing).Value2.ToString());
            mDI.set도매_수익_유통모델매입관련추가수익_현금DC(worksheet1.get_Range("G30", Type.Missing).Value2.ToString());
            mDI.set도매_수익_유통모델매입관련추가수익_VolumeDC(worksheet1.get_Range("G31", Type.Missing).Value2.ToString());
            mDI.set도매_비용_대리점투자금액_신규(worksheet1.get_Range("G32", Type.Missing).Value2.ToString());
            mDI.set도매_비용_대리점투자금액_기변(worksheet1.get_Range("G33", Type.Missing).Value2.ToString());
            mDI.set도매_비용_직원급여_간부급(worksheet1.get_Range("G34", Type.Missing).Value2.ToString());//총액
            mDI.set도매_비용_직원급여_평사원(worksheet1.get_Range("G35", Type.Missing).Value2.ToString());//총액
            mDI.set도매_비용_지급임차료(worksheet1.get_Range("G38", Type.Missing).Value2.ToString());
            mDI.set도매_비용_운반비(worksheet1.get_Range("G39", Type.Missing).Value2.ToString());
            mDI.set도매_비용_차량유지비(worksheet1.get_Range("G40", Type.Missing).Value2.ToString());
            mDI.set도매_비용_지급수수료(worksheet1.get_Range("G41", Type.Missing).Value2.ToString());
            mDI.set도매_비용_판매촉진비(worksheet1.get_Range("G42", Type.Missing).Value2.ToString());
            mDI.set도매_비용_건물관리비(worksheet1.get_Range("G43", Type.Missing).Value2.ToString());

            mDI.set소매_수익_월평균업무취급수수료(worksheet1.get_Range("G44", Type.Missing).Value2.ToString());
            mDI.set소매_수익_직영매장판매수익(worksheet1.get_Range("G45", Type.Missing).Value2.ToString());
            mDI.set소매_비용_직원급여_간부급(worksheet1.get_Range("G46", Type.Missing).Value2.ToString());//총액
            mDI.set소매_비용_직원급여_평사원(worksheet1.get_Range("G47", Type.Missing).Value2.ToString());//총액
            mDI.set소매_비용_지급임차료(worksheet1.get_Range("G50", Type.Missing).Value2.ToString());
            mDI.set소매_비용_지급수수료(worksheet1.get_Range("G51", Type.Missing).Value2.ToString());
            mDI.set소매_비용_판매촉진비(worksheet1.get_Range("G52", Type.Missing).Value2.ToString());
            mDI.set소매_비용_건물관리비(worksheet1.get_Range("G53", Type.Missing).Value2.ToString());

            mDI.set도소매_비용_복리후생비(worksheet1.get_Range("G54", Type.Missing).Value2.ToString());
            mDI.set도소매_비용_통신비(worksheet1.get_Range("G55", Type.Missing).Value2.ToString());
            mDI.set도소매_비용_공과금(worksheet1.get_Range("G56", Type.Missing).Value2.ToString());
            mDI.set도소매_비용_소모품비(worksheet1.get_Range("G57", Type.Missing).Value2.ToString());
            mDI.set도소매_비용_이자비용(worksheet1.get_Range("G58", Type.Missing).Value2.ToString());
            mDI.set도소매_비용_부가세(worksheet1.get_Range("G59", Type.Missing).Value2.ToString());
            mDI.set도소매_비용_법인세(worksheet1.get_Range("G60", Type.Missing).Value2.ToString());
            mDI.set도소매_비용_기타(worksheet1.get_Range("G61", Type.Missing).Value2.ToString());

            //*******CResultData
            mRD.set도매_수익_가입자관리수수료(worksheet2.get_Range("E7", Type.Missing).Value2.ToString());
            mRD.set도매_수익_CS관리수수료(worksheet2.get_Range("E8", Type.Missing).Value2.ToString());
            mRD.set소매_수익_업무취급수수료(worksheet2.get_Range("E9", Type.Missing).Value2.ToString());
            mRD.set도매_수익_사업자모델매입에따른추가수익(worksheet2.get_Range("E10", Type.Missing).Value2.ToString());
            mRD.set도매_수익_유통모델매입에따른추가수익_현금_Volume(worksheet2.get_Range("E11", Type.Missing).Value2.ToString());
            mRD.set소매_수익_직영매장판매수익(worksheet2.get_Range("E12", Type.Missing).Value2.ToString());
            mRD.set전체_비용_대리점투자비용(worksheet2.get_Range("E14", Type.Missing).Value2.ToString());
            mRD.set전체_비용_인건비_급여_복리후생비(worksheet2.get_Range("E15", Type.Missing).Value2.ToString());
            mRD.set전체_비용_임차료(worksheet2.get_Range("E16", Type.Missing).Value2.ToString());
            mRD.set전체_비용_이자비용(worksheet2.get_Range("E17", Type.Missing).Value2.ToString());
            mRD.set전체_비용_부가세(worksheet2.get_Range("E18", Type.Missing).Value2.ToString());
            mRD.set전체_비용_법인세(worksheet2.get_Range("E19", Type.Missing).Value2.ToString());
            mRD.set전체_비용_기타판매관리비(worksheet2.get_Range("E20", Type.Missing).Value2.ToString());

            mRD.set도매_비용_대리점투자비용(worksheet2.get_Range("E33", Type.Missing).Value2.ToString());
            mRD.set도매_비용_인건비_급여_복리후생비(worksheet2.get_Range("E34", Type.Missing).Value2.ToString());
            mRD.set도매_비용_임차료(worksheet2.get_Range("E35", Type.Missing).Value2.ToString());
            mRD.set도매_비용_이자비용(worksheet2.get_Range("E36", Type.Missing).Value2.ToString());
            mRD.set도매_비용_부가세(worksheet2.get_Range("E37", Type.Missing).Value2.ToString());
            mRD.set도매_비용_법인세(worksheet2.get_Range("E38", Type.Missing).Value2.ToString());
            mRD.set도매_비용_기타판매관리비(worksheet2.get_Range("E39", Type.Missing).Value2.ToString());

            mRD.set소매_비용_인건비_급여_복리후생비(worksheet2.get_Range("E49", Type.Missing).Value2.ToString());
            mRD.set소매_비용_임차료(worksheet2.get_Range("E50", Type.Missing).Value2.ToString());
            mRD.set소매_비용_이자비용(worksheet2.get_Range("E51", Type.Missing).Value2.ToString());
            mRD.set소매_비용_부가세(worksheet2.get_Range("E52", Type.Missing).Value2.ToString());
            mRD.set소매_비용_법인세(worksheet2.get_Range("E53", Type.Missing).Value2.ToString());
            mRD.set소매_비용_기타판매관리비(worksheet2.get_Range("E54", Type.Missing).Value2.ToString());
        }

        private Int32 getExcelResultAsInt32(excel.Worksheet workSheet, string columnName)
        {
            object result = workSheet.get_Range(columnName, Type.Missing).Value2;
            return result == null ? 0 : Convert.ToInt32(result);
        }

        private string getExcelResult(excel.Worksheet workSheet, string columnName)
        {
            object result = workSheet.get_Range(columnName, Type.Missing).Value2;
            return result == null ? "" : (string)result;
        }

        private string getExcelResultAsInt64(excel.Worksheet workSheet, string columnName)
        {
            return CommonUtil.NullToString0(workSheet.get_Range(columnName, Type.Missing).Value2);
        }

        private void readFileOfExistedAverage()
        {
            try
            {
                string csv = System.IO.File.ReadAllText(CommonUtil.adminName);
                csv = CommonUtil.Base64Decode(csv);
                string[] splitedCsv = csv.Split(',');
                for (int i = 0; i < txtOut.Length; i++)
                {
                    txtOut[i].Text = splitedCsv[i];
                }
            }
            catch (Exception ex)
            {
                // 파일이 없음
                for (int i = 0; i < txtOut.Length; i++)
                {
                    txtOut[i].Text = 0.ToString();
                }
            }
        }

        private void saveFileOfExistedAverage(TextBox[] txtBoxes)
        {
            string csv = "";
            for (int i = 0; i < txtOut.Length; i++)
            {
                csv += txtBoxes[i].Text + ",";
            }
            System.IO.File.WriteAllText(CommonUtil.adminName, CommonUtil.Base64Encode(csv));
            readFileOfExistedAverage();
        }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ListView listView = (ListView)sender;
            ListViewItem item = listView.GetItemAt(e.X, e.Y);
            adminDC.toggle선택여부((string)item.Tag);

            refreshList();
        }

        private void txtOut_TextChanged(object sender, EventArgs e)
        {
            int index = -1;
            index = Array.IndexOf(txtIAOut, (sender as TextBox));
            if (index < 0)
            {
                index = Array.IndexOf(txtInput, (sender as TextBox));
                if (index < 0) return;
            }
            Int64 convertedA;
            Int64 convertedB;
            Int64 result;
            try
            {
                convertedA = Convert.ToInt64(txtIAOut[index].Text);
            }
            catch (FormatException eFormat)
            {
                txtIAOut[index].Text = "0";
                convertedA = Convert.ToInt64(txtIAOut[index].Text);
                MessageBox.Show("문서에 숫자가 아닌 문자가 있습니다.");
            }
            try
            {
                convertedB = Convert.ToInt64(txtInput[index].Text);
            }
            catch (FormatException eFormat)
            {
                txtInput[index].Text = "0";
                convertedB = Convert.ToInt64(txtInput[index].Text);
            }
            result = convertedA;
            if (convertedB > 0)
            {
                result = convertedB;//(convertedA + convertedB) / 2;
            }
            txtAOut[index].Text = result.ToString();
        }

        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            this.listView1.Sort();
        }

    }
}
