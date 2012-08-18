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
using System.Security.Permissions;

namespace KIWI
{

    public partial class FormAdmin : Form
    {
        public const string CHECK = "√";

        private CAdminDataController adminDC = null;
        private ListViewColumnSorter lvwColumnSorter = null;

        private TextBox[] txt기존업계평균 = null;        //기존 업계평균
        private TextBox[] txt업계평균 = null;      // 업계평균
        private TextBox[] txtInput = null;      //보정 계수
        private TextBox[] txtAOut = null;       //보정 계수 업계평균

        private TextBox[] txtExistedAsp = null;        //기존 ASP
        private TextBox[] txtInputAsp = null;        //ASP 입력창
        private TextBox[] txtAInputAsp = null;        //ASP 입력창

        private Double[] 기존업계평균 = null;
        private Double[] 업계평균 = null;
        private Double[] nInput = null;
        private Double[] nAOut = null;

        public FormAdmin()
        {
            InitializeComponent();
            adminDC = new CAdminDataController();
            lvwColumnSorter = new ListViewColumnSorter();
            this.listView1.ListViewItemSorter = lvwColumnSorter;

            txt기존업계평균 = new TextBox[26] { 
                txtOut1, txtOut2, txtOut6, txtOut7, txtOut8, txtOut9, txtOut10,
                txtOut11, txtOut12, txtOut13, txtOut14, txtOut15, txtOut16, txtOut17, txtOut18, txtOut19, txtOut20,
                txtOut21, txtOut22, txtOut23, txtOut24, txtOut25, txtOut26, txtOut27, txtOut28, 
                txtOut31
            };

            txt업계평균 = new TextBox[26] {
                txtOut32, txtOut33, txtOut37, txtOut38, txtOut39, txtOut40,
                txtOut41, txtOut42, txtOut43, txtOut44, txtOut45, txtOut46, txtOut47, txtOut48, txtOut49, txtOut50,
                txtOut51, txtOut52, txtOut53, txtOut54, txtOut55, txtOut56, txtOut57, txtOut58, txtOut59, 
                txtOut62
            };

            txtInput = new TextBox[26] { 
                    txtInput1, txtInput2, txtInput6, txtInput7, txtInput8, txtInput9, txtInput10,
                    txtInput11, txtInput12, txtInput13, txtInput14, txtInput15, txtInput16, txtInput17, txtInput18, txtInput19, txtInput20,
                    txtInput21, txtInput22, txtInput23, txtInput24, txtInput25, txtInput26, txtInput27, txtInput28, 
                    txtInput31, 
            };

            txtAOut = new TextBox[26] { 
                txtAOut1, txtAOut2, txtAOut6, txtAOut7, txtAOut8, txtAOut9, txtAOut10,
                txtAOut11, txtAOut12, txtAOut13, txtAOut14, txtAOut15, txtAOut16, txtAOut17, txtAOut18, txtAOut19, txtAOut20,
                txtAOut21, txtAOut22, txtAOut23, txtAOut24, txtAOut25, txtAOut26, txtAOut27, txtAOut28, 
                txtAOut31
            };

            txtExistedAsp = new TextBox[8] { textBox1, textBox2, textBox3, textBox4, textBox5, textBox6, textBox7, textBox8 };

            txtInputAsp = new TextBox[8] { 유통모델_LG, 유통모델_SS, 유통모델_소계, 사업자모델_LG, 사업자모델_SS, 사업자모델_소계, ASP_전체계, 리베이트 };
            txtAInputAsp = new TextBox[8] { out유통모델_LG, out유통모델_SS, out유통모델_소계, out사업자모델_LG, out사업자모델_SS, out사업자모델_소계, outASP_전체계, out리베이트 };



            기존업계평균 = new Double[26];
            업계평균 = new Double[26];
            nInput = new Double[26];
            nAOut = new Double[26];

            readFileOfExistedAverage();

            refreshList();

            applyFileNameLabel();
        }

        //업계평균과 보정계수의 평균을 구하고 보정계수 평균에 적용. 
        private void button2_Click(object sender, EventArgs e)
        {
            saveFileOfExistedAverage();
        }

        // 파일열기
        public void openFileDialog(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "LGE File|*.lge|Excel File|*.xlsx|All File|*.*";
            openFileDialog1.Title = "Select a File";
            openFileDialog1.InitialDirectory = CommonUtil.dataDirectory;
            openFileDialog1.DefaultExt = "lge";
            openFileDialog1.AutoUpgradeEnabled = true;
            openFileDialog1.AddExtension = true;
            openFileDialog1.RestoreDirectory = true;

            // If the directory doesn't exist, create it.
            if (!Directory.Exists(openFileDialog1.InitialDirectory))
            {
                Directory.CreateDirectory(openFileDialog1.InitialDirectory);
            }

            // Show the Dialog.
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Read the files
                foreach (String file in openFileDialog1.FileNames)
                {
                    try
                    {
                        if (file.EndsWith("lge"))
                        {
                            setLGEFileToXML(file);
                        }
                        else if (file.EndsWith("xlsx"))
                        {
                            setExcelFileToXML(file);
                        }
                        else
                        {
                            throw new Exception("지원하지 않는 파일형식입니다.");
                        }
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
                item.SubItems.Add(String.Format("{0:#,###}", Convert.ToDouble(단위당손익)));
                item.SubItems.Add(String.Format("{0:#,###}", Convert.ToDouble(월capa)));
                item.SubItems.Add(String.Format("{0:#,###}", Convert.ToDouble(가입자수)));
                item.SubItems.Add(String.Format("{0:#,###}", Convert.ToDouble(직영점판매수익)));
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

            for (int i = 0; i < 업계평균.Length; i++)
            {
                업계평균[i] = 0;
            }

            int 분모 = 0;
            for (int i = 0; i < adminDC.getDataLength(); i++)
            {
                adminDC.GetData(i, out key, out 지역, out 대리점명, out 마케터, out 단위당손익, out 월capa, out 가입자수, out 직영점판매수익, out 선택여부, out mExcelFileName, out mBI, out mDI, out mRD);

                if (선택여부 == "N") continue;
                분모++;

                int k = 0;
                업계평균[k++] += CommonUtil.Division(mDI.get도매_수익_월평균관리수수료(), mBI.get누적가입자수_합계());
                업계평균[k++] += CommonUtil.Division(mDI.get도매_수익_CS관리수수료(), mBI.get누적가입자수_합계());

                //nIAOut[k++] += CommonUtil.Division(mDI.get도매_비용_대리점투자금액_신규() , mBI.get도매_월평균판매대수_신규());
                업계평균[k++] += mDI.get도매_비용_대리점투자금액_신규();// 이미 단위금액임;
                //nIAOut[k++] += CommonUtil.Division(mDI.get도매_비용_대리점투자금액_기변() , mBI.get도매_월평균판매대수_기변());
                업계평균[k++] += mDI.get도매_비용_대리점투자금액_기변();// 이미 단위금액임;
                업계평균[k++] += mDI.get도매_비용_직원급여_간부급(); // 단위금액
                업계평균[k++] += mDI.get도매_비용_직원급여_평사원(); // 단위금액
                업계평균[k++] += CommonUtil.Division(mDI.get도매_비용_지급임차료(), mBI.get도매_거래선수_개통사무실());
                업계평균[k++] += CommonUtil.Division(mDI.get도매_비용_운반비(), mBI.get도매_월평균판매대수_소계());
                업계평균[k++] += CommonUtil.Division(mDI.get도매_비용_차량유지비(), mBI.get도매_직원수_소계());
                업계평균[k++] += CommonUtil.Division(mDI.get도매_비용_지급수수료(), mBI.get도매_월평균판매대수_소계());
                업계평균[k++] += CommonUtil.Division(mDI.get도매_비용_판매촉진비(), mBI.get도매_월평균판매대수_소계());
                업계평균[k++] += CommonUtil.Division(mDI.get도매_비용_건물관리비(), mBI.get도매_거래선수_개통사무실());

                업계평균[k++] += CommonUtil.Division(mDI.get소매_수익_월평균업무취급수수료(), mBI.get월평균판매대수_소계_합계());
                업계평균[k++] += CommonUtil.Division(mDI.get소매_수익_직영매장판매수익(), mBI.get소매_월평균판매대수_소계());
                업계평균[k++] += mDI.get소매_비용_직원급여_간부급(); // 단위금액
                업계평균[k++] += mDI.get소매_비용_직원급여_평사원(); // 단위금액
                업계평균[k++] += CommonUtil.Division(mDI.get소매_비용_지급임차료(), mBI.get소매_거래선수_소계());
                업계평균[k++] += CommonUtil.Division(mDI.get소매_비용_지급수수료(), mBI.get소매_월평균판매대수_소계());
                업계평균[k++] += CommonUtil.Division(mDI.get소매_비용_판매촉진비(), mBI.get소매_월평균판매대수_소계());
                업계평균[k++] += CommonUtil.Division(mDI.get소매_비용_건물관리비(), mBI.get소매_거래선수_소계());

                업계평균[k++] += CommonUtil.Division(mDI.get도소매_비용_복리후생비(), mBI.get직원수_소계_합계());
                업계평균[k++] += CommonUtil.Division(mDI.get도소매_비용_통신비(), mBI.get직원수_소계_합계());
                업계평균[k++] += CommonUtil.Division(mDI.get도소매_비용_공과금(), mBI.get직원수_소계_합계());
                업계평균[k++] += CommonUtil.Division(mDI.get도소매_비용_소모품비(), mBI.get월평균판매대수_소계_합계());
                업계평균[k++] += CommonUtil.Division(mDI.get도소매_비용_이자비용(), mBI.get월평균판매대수_소계_합계());
                //nIAOut[k++] += mDI.get도소매_비용_이자비용();// 이미 평균금액이라 단위금액으로 판단
                업계평균[k++] += CommonUtil.Division(mDI.get도소매_비용_기타(), mBI.get월평균판매대수_소계_합계());
            }

            for (int i = 0; i < 업계평균.Length; i++)
            {
                txt업계평균[i].Text = CommonUtil.Division(업계평균[i] , 분모).ToString();
            }

        }

        private void setLGEFileToXML(string file)
        {
            CBasicInput mBI = null;
            CBusinessData mDI = null;
            CResultData mRD = null;

            setDataForUse(file, "|", out mBI, out mDI, out mRD);

            adminDC.AddSaveData(
                mBI.get지역(),
                mBI.get대리점(),
                mBI.get마케터(),
                mRD.전체손익계.ToString(),
                mBI.get월평균판매대수_소계_합계().ToString(),
                mBI.get누적가입자수_합계().ToString(),
                mDI.get소매_수익_직영매장판매수익().ToString(),
                "Y",
                file,
                mBI, mDI, mRD
            );

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
                mRD.전체손익계.ToString(),
                mBI.get월평균판매대수_소계_합계().ToString(),
                mBI.get누적가입자수_합계().ToString(),
                mDI.get소매_수익_직영매장판매수익().ToString(),
                "Y",
                file,
                mBI, mDI, mRD
            );

        }

        private void setDataForUse(String filepath, String spliter, out CBasicInput mBI, out CBusinessData mDI, out CResultData mRD)
        {
            mBI = new CBasicInput();
            mDI = new CBusinessData();
            mRD = new CResultData();

            if (filepath == null || mBI == null || mDI == null || mRD == null) return;

            CReportData reportData = new CReportData();
            String lge = CommonUtil.Base64Decode(System.IO.File.ReadAllText(filepath));
            String[] splittedLge = lge.Split(spliter.ToCharArray());

            int startIndex = 0;
            int length = 6;
            String[] param = splittedLge.Take(length).ToArray<String>();
            reportData.setArrData(param);

            mBI.set지역(reportData.get지역());
            mBI.set대리점(reportData.get대리점());
            mBI.set마케터(reportData.get마케터());

            startIndex += length;
            length = 14;
            param = splittedLge.Skip(startIndex).Take(length).ToArray<String>();
            mBI.setArrData(param);
            startIndex += length;
            length = 31;
            param = splittedLge.Skip(startIndex).Take(length).ToArray<String>();
            mDI.setArrData(param);
            startIndex += length;
            length = 42;
            //param = splittedLge.Skip(startIndex).Take(length).ToArray<String>();
            //CDataControl.g_FileResultBusinessTotal.setArrayOutput전체(param);
            startIndex += length;
            //param = splittedLge.Skip(startIndex).Take(length).ToArray<String>();
            //CDataControl.g_FileResultBusiness.setArrayOutput전체(param);
            startIndex += length;
            //param = splittedLge.Skip(startIndex).Take(length).ToArray<String>();
            //CDataControl.g_FileResultStoreTotal.setArrayOutput전체(param);
            startIndex += length;
            param = splittedLge.Skip(startIndex).Take(length).ToArray<String>();
            mRD.setArrayOutput전체(param);
            //startIndex += length;
            //param = splittedLge.Skip(startIndex).Take(length).ToArray<String>();
            //CDataControl.g_FileResultFutureTotal.setArrayOutput전체(param);
            //startIndex += length;
            //param = splittedLge.Skip(startIndex).Take(length).ToArray<String>();
            //CDataControl.g_FileResultFuture.setArrayOutput전체(param);
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
            //mDI.set도매_수익_사업자모델매입관련추가수익(worksheet1.get_Range("G29", Type.Missing).Value2.ToString());
            //mDI.set도매_수익_유통모델매입관련추가수익_현금DC(worksheet1.get_Range("G30", Type.Missing).Value2.ToString());
            //mDI.set도매_수익_유통모델매입관련추가수익_VolumeDC(worksheet1.get_Range("G31", Type.Missing).Value2.ToString());
            mDI.set도매_비용_대리점투자금액_신규(worksheet1.get_Range("G32", Type.Missing).Value2.ToString());
            mDI.set도매_비용_대리점투자금액_기변(worksheet1.get_Range("G33", Type.Missing).Value2.ToString());
            mDI.set도매_비용_직원급여_간부급(worksheet1.get_Range("G36", Type.Missing).Value2.ToString());
            mDI.set도매_비용_직원급여_평사원(worksheet1.get_Range("G37", Type.Missing).Value2.ToString());
            mDI.set도매_비용_지급임차료(worksheet1.get_Range("G38", Type.Missing).Value2.ToString());
            mDI.set도매_비용_운반비(worksheet1.get_Range("G39", Type.Missing).Value2.ToString());
            mDI.set도매_비용_차량유지비(worksheet1.get_Range("G40", Type.Missing).Value2.ToString());
            mDI.set도매_비용_지급수수료(worksheet1.get_Range("G41", Type.Missing).Value2.ToString());
            mDI.set도매_비용_판매촉진비(worksheet1.get_Range("G42", Type.Missing).Value2.ToString());
            mDI.set도매_비용_건물관리비(worksheet1.get_Range("G43", Type.Missing).Value2.ToString());

            mDI.set소매_수익_월평균업무취급수수료(worksheet1.get_Range("G44", Type.Missing).Value2.ToString());
            mDI.set소매_수익_직영매장판매수익(worksheet1.get_Range("G45", Type.Missing).Value2.ToString());
            mDI.set소매_비용_직원급여_간부급(worksheet1.get_Range("G48", Type.Missing).Value2.ToString());
            mDI.set소매_비용_직원급여_평사원(worksheet1.get_Range("G49", Type.Missing).Value2.ToString());
            mDI.set소매_비용_지급임차료(worksheet1.get_Range("G50", Type.Missing).Value2.ToString());
            mDI.set소매_비용_지급수수료(worksheet1.get_Range("G51", Type.Missing).Value2.ToString());
            mDI.set소매_비용_판매촉진비(worksheet1.get_Range("G52", Type.Missing).Value2.ToString());
            mDI.set소매_비용_건물관리비(worksheet1.get_Range("G53", Type.Missing).Value2.ToString());

            mDI.set도소매_비용_복리후생비(worksheet1.get_Range("G54", Type.Missing).Value2.ToString());
            mDI.set도소매_비용_통신비(worksheet1.get_Range("G55", Type.Missing).Value2.ToString());
            mDI.set도소매_비용_공과금(worksheet1.get_Range("G56", Type.Missing).Value2.ToString());
            mDI.set도소매_비용_소모품비(worksheet1.get_Range("G57", Type.Missing).Value2.ToString());
            mDI.set도소매_비용_이자비용(worksheet1.get_Range("G58", Type.Missing).Value2.ToString());
            //mDI.set도소매_비용_부가세(worksheet1.get_Range("G59", Type.Missing).Value2.ToString());
            //mDI.set도소매_비용_법인세(worksheet1.get_Range("G60", Type.Missing).Value2.ToString());
            mDI.set도소매_비용_기타(worksheet1.get_Range("G61", Type.Missing).Value2.ToString());

            //*******CResultData
            mRD.set도매_수익_가입자관리수수료(worksheet2.get_Range("J7", Type.Missing).Value2.ToString());
            mRD.set도매_수익_CS관리수수료(worksheet2.get_Range("J8", Type.Missing).Value2.ToString());
            mRD.set소매_수익_업무취급수수료(worksheet2.get_Range("J9", Type.Missing).Value2.ToString());
            mRD.set도매_수익_사업자모델매입에따른추가수익(worksheet2.get_Range("J10", Type.Missing).Value2.ToString());
            mRD.set도매_수익_유통모델매입에따른추가수익_현금_Volume(worksheet2.get_Range("J11", Type.Missing).Value2.ToString());
            mRD.set소매_수익_직영매장판매수익(worksheet2.get_Range("J12", Type.Missing).Value2.ToString());
            mRD.set전체_비용_대리점투자비용(worksheet2.get_Range("J14", Type.Missing).Value2.ToString());
            mRD.set전체_비용_인건비_급여_복리후생비(worksheet2.get_Range("J15", Type.Missing).Value2.ToString());
            mRD.set전체_비용_임차료(worksheet2.get_Range("J16", Type.Missing).Value2.ToString());
            mRD.set전체_비용_이자비용(worksheet2.get_Range("J17", Type.Missing).Value2.ToString());
            mRD.set전체_비용_부가세(worksheet2.get_Range("J18", Type.Missing).Value2.ToString());
            mRD.set전체_비용_법인세(worksheet2.get_Range("J19", Type.Missing).Value2.ToString());
            mRD.set전체_비용_기타판매관리비(worksheet2.get_Range("J20", Type.Missing).Value2.ToString());

            mRD.set도매_비용_대리점투자비용(worksheet2.get_Range("J33", Type.Missing).Value2.ToString());
            mRD.set도매_비용_인건비_급여_복리후생비(worksheet2.get_Range("J34", Type.Missing).Value2.ToString());
            mRD.set도매_비용_임차료(worksheet2.get_Range("J35", Type.Missing).Value2.ToString());
            mRD.set도매_비용_이자비용(worksheet2.get_Range("J36", Type.Missing).Value2.ToString());
            mRD.set도매_비용_부가세(worksheet2.get_Range("J37", Type.Missing).Value2.ToString());
            mRD.set도매_비용_법인세(worksheet2.get_Range("J38", Type.Missing).Value2.ToString());
            mRD.set도매_비용_기타판매관리비(worksheet2.get_Range("J39", Type.Missing).Value2.ToString());

            mRD.set소매_비용_인건비_급여_복리후생비(worksheet2.get_Range("J49", Type.Missing).Value2.ToString());
            mRD.set소매_비용_임차료(worksheet2.get_Range("J50", Type.Missing).Value2.ToString());
            mRD.set소매_비용_이자비용(worksheet2.get_Range("J51", Type.Missing).Value2.ToString());
            mRD.set소매_비용_부가세(worksheet2.get_Range("J52", Type.Missing).Value2.ToString());
            mRD.set소매_비용_법인세(worksheet2.get_Range("J53", Type.Missing).Value2.ToString());
            mRD.set소매_비용_기타판매관리비(worksheet2.get_Range("J54", Type.Missing).Value2.ToString());
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
                string csv = System.IO.File.ReadAllText(CommonUtil.defaultManagerFileName);
                csv = CommonUtil.Base64Decode(csv);
                string[] splitedCsv = csv.Split(',');
                for (int i = 0; i < txt기존업계평균.Length; i++)
                {
                    txt기존업계평균[i].Text = splitedCsv[i];
                }
                for (int i = 0; i < txtExistedAsp.Length; i++)
                {
                    txtExistedAsp[i].Text = splitedCsv[txt기존업계평균.Length + i];
                }
            }
            catch (Exception ex)
            {
                // 파일이 없음
                for (int i = 0; i < txt기존업계평균.Length; i++)
                {
                    txt기존업계평균[i].Text = 0.ToString();
                }
                for (int i = 0; i < txtExistedAsp.Length; i++)
                {
                    txtExistedAsp[i].Text = 0.ToString();
                }
            }
        }

        private void saveFileOfExistedAverage()
        {
            string csv = "";
            for (int i = 0; i < txtAOut.Length; i++)
            {
                csv += txtAOut[i].Text.Replace(",", "") + ",";
            }
            for (int i = 0; i < txtInputAsp.Length; i++)
            {
                csv += txtInputAsp[i].Text.Replace(",", "") + ",";
            }
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "LGM File|*.lgm";
            saveFileDialog1.Title = "Select a File";
            saveFileDialog1.InitialDirectory = CommonUtil.업계평균Directory;
            saveFileDialog1.DefaultExt = "lgm";
            saveFileDialog1.AutoUpgradeEnabled = true;
            saveFileDialog1.AddExtension = true;
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = CommonUtil.datedManagerFileName;

            // If the directory doesn't exist, create it.
            if (!Directory.Exists(saveFileDialog1.InitialDirectory))
            {
                Directory.CreateDirectory(saveFileDialog1.InitialDirectory);
            }
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                FileIOPermission permission = new FileIOPermission(FileIOPermissionAccess.AllAccess, CommonUtil.defaultManagerFileName);
                permission.Demand();
                System.IO.File.WriteAllText(CommonUtil.defaultManagerFileName, CommonUtil.Base64Encode(csv));
                permission = new FileIOPermission(FileIOPermissionAccess.AllAccess, saveFileDialog1.FileName);
                permission.Demand();
                System.IO.File.WriteAllText(saveFileDialog1.FileName, CommonUtil.Base64Encode(csv));
                readFileOfExistedAverage();
                setFileNameLabel(saveFileDialog1.FileName);
            }
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
            setTxtInput_TextChanged(sender);

            int index = -1;
            index = Array.IndexOf(txt업계평균, (sender as TextBox));
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
                convertedA = Convert.ToInt64(txt업계평균[index].Text.Replace(",", ""));
            }
            catch (FormatException eFormat)
            {
                txt업계평균[index].Text = "0";
                convertedA = Convert.ToInt64(txt업계평균[index].Text);
                MessageBox.Show("문서에 숫자가 아닌 문자가 있습니다.");
            }
            try
            {
                convertedB = Convert.ToInt64(txtInput[index].Text.Replace(",", ""));
            }
            catch (FormatException eFormat)
            {
                txtInput[index].Text = "0";
                convertedB = Convert.ToInt64(txtInput[index].Text);
            }
            result = convertedA;
            if (convertedB != 0)
            {
                result = convertedB;
            }
            txtAOut[index].Text = result.ToString();
            setTxtInput_TextChanged(txtAOut[index]);
        }

        private void txtASPOut_TextChanged(object sender, EventArgs e)
        {
            setTxtInput_TextChanged(sender);

            int index = -1;
            index = Array.IndexOf(txtInputAsp, (sender as TextBox));
            if (index < 0) return;

            Int64 convertedA;
            Int64 result;
            try
            {
                convertedA = Convert.ToInt64(txtInputAsp[index].Text.Replace(",", ""));
            }
            catch (FormatException eFormat)
            {
                txtInputAsp[index].Text = "0";
                convertedA = Convert.ToInt64(txtInputAsp[index].Text);
                MessageBox.Show("문서에 숫자가 아닌 문자가 있습니다.");
            }
            result = convertedA;
            txtAInputAsp[index].Text = result.ToString();
            setTxtInput_TextChanged(txtAInputAsp[index]);
        }

        

        private string setTxtInput_TextChanged(object sender)
        {
            TextBox _TextBox = (sender as TextBox);

            try
            {
                Double num = Convert.ToDouble(_TextBox.Text.Replace(",", ""));

                if (_TextBox.Text.Length < 24 && _TextBox.Text.Length > 1)
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

        private void listView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete) {
                ListView listView = (ListView)sender;
                ListView.SelectedListViewItemCollection items = listView.SelectedItems;
                foreach (ListViewItem item in items)
                {
                    adminDC.deleteData((string)item.Tag);
                }

                refreshList();
            }
        }

        private void txtInput1_Click(object sender, EventArgs e)
        {
            TextBox _TextBox = (sender as TextBox);
            if (_TextBox.Text == "0")
            {
                _TextBox.SelectAll();
            }

        }

        private void txtInput2_KeyPress(object sender, KeyPressEventArgs e)
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

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (TextBox txt in txtInput)
            {
                txt.Text = "0";
            }
            foreach (TextBox txt in txtInputAsp)
            {
                txt.Text = "0";
            }
        }

        private void setFileNameLabel(String fileName)
        {
            String[] splitedFileName = fileName.Split('\\');
            System.IO.File.WriteAllText(CommonUtil.fileNameLabelFileName, CommonUtil.Base64Encode(splitedFileName[splitedFileName.Length - 1]));
            applyFileNameLabel();
        }

        private void applyFileNameLabel()
        {
            try
            {
                파일명.Text = CommonUtil.Base64Decode(System.IO.File.ReadAllText(CommonUtil.fileNameLabelFileName));
            }
            catch (FileNotFoundException fnfe)
            {
                // FileNotFoundException 무시
            }
        }
    }
}