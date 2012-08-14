using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Collections;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Security.Permissions;
namespace KIWI
{
    public class CommonUtil
    {

        public const Int64 QUARTER = 3;
        public const int 파일종류_기본 = 0;
        public const int 파일종류_관리자 = 1;
        public const int 파일종류_시뮬레이션 = 2;
        private static excel.ApplicationClass application = null;
        private static excel.Workbook workBook = null;
        private static excel.ApplicationClass applicationForSimul = null;
        private static excel.Workbook workBookForSimul = null;
        public static Boolean isLoadedFromFile = false;
        public static Boolean isSimulatedOnce = false;
        public static string defaultName = AppDomain.CurrentDomain.BaseDirectory + "files\\default.xlsx";
        //public static string openAsName = null;
        public static string dataDirectory = AppDomain.CurrentDomain.BaseDirectory + "data\\";
        public static string 업계평균Directory = AppDomain.CurrentDomain.BaseDirectory + "업계평균\\";
        public static string saveAsSimulName = null;
        public static string saveAsName = null;
        public static string defaultManagerFileName = AppDomain.CurrentDomain.BaseDirectory + "files\\manager.lgm";
        public static string datedManagerFileName = "업계평균_" + DateTime.Now.ToString("yyyyMMdd") + ".lgm";

        /// <summary>
        /// 
        /// </summary>
        /// <param name="text1"></param>
        /// <param name="text2"></param>
        /// <returns></returns>
        public static string Sum_Values(string text1, string text2)
        {
            Double sumManager = StringToDoubleVal(text1) + StringToDoubleVal(text2);

            return sumManager.ToString();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="text1"></param>
        /// <returns></returns>
        public static Double StringToDoubleVal(object text1)
        {
            string returnValue = "";
            if (text1 != null)
            {
                if (text1 is string)
                {
                    returnValue = (text1 as string);
                }
            }

            bool result = returnValue == "";
            if (text1 is string)
                result = returnValue.Length < 1;
            try{
                return result ? 0 : Convert.ToDouble(Convert.ToDouble(returnValue));
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="text1"></param>
        /// <returns></returns>
        public static string NullToString0(object obj)
        {
            return obj == null ? 0.ToString() : Convert.ToDouble(obj).ToString();
        }


        /// <summary>
        /// 분모가 0이거나, 분자가 0일경우 0을 반환, 이외의 경우 나눈 몫을 반환
        /// </summary>
        /// <param name="string분자">분모값</param>
        /// <param name="string분모">분자값</param>
        /// <returns>분모가 0이거나, 분자가 0일경우 0을 반환, 이외의 경우 나눈 몫을 반환</returns>
        public static string Division(string string분자, string string분모)
        {
            return Division(StringToDoubleVal(string분자),StringToDoubleVal(string분모)).ToString();
        }


        /// <summary>
        /// 분모가 0이거나, 분자가 0일경우 0을 반환, 이외의 경우 나눈 몫을 반환
        /// </summary>
        /// <param name="string분자">분모값</param>
        /// <param name="string분모">분자값</param>
        /// <returns>분모가 0이거나, 분자가 0일경우 0을 반환, 이외의 경우 나눈 몫을 반환</returns>
        public static Int64 Division(Int64 분자, Int64 분모)
        {
            return Convert.ToInt64(Division(Convert.ToDouble(분자), Convert.ToDouble(분모)));
        }


        /// <summary>
        /// 분모가 0이거나, 분자가 0일경우 0을 반환, 이외의 경우 나눈 몫을 반환
        /// </summary>
        /// <param name="string분자">분모값</param>
        /// <param name="string분모">분자값</param>
        /// <returns>분모가 0이거나, 분자가 0일경우 0을 반환, 이외의 경우 나눈 몫을 반환</returns>
        public static Double Division(Double 분자, Double 분모)
        {
            return 분자 != 0 ? (분모 == 0 ? 0 : 분자 / 분모) : 0;
        }

        /// <summary>
        /// 분모가 0이거나, 분자가 0일경우 0을 반환, 이외의 경우 나눈 몫을 반환
        /// </summary>
        /// <param name="textDenominator">분모값</param>
        /// <param name="textNumerator">분자값</param>
        /// <returns>분모가 0이거나, 분자가 0일경우 0을 반환, 이외의 경우 나눈 몫을 반환</returns>
        public static excel.Worksheet GetExcelWorksheet(string fileName, int sheetNo)
        {
            // 엑셀문서(워크북)를 읽기 전용으로 열기.
            //Microsoft.Office.Interop.Excel.Workbook _WorkBook = _Application.Workbooks.Open(문서경로,0,true,5,Missing.Value,Missing.Value,false,Missing.Value,
            excel.Workbook _WorkBook = GetExcel_WorkBook(fileName);

            // sheets 생성
            excel.Sheets _Sheets = _WorkBook.Sheets;

            // 작업할 sheet 선택
            excel.Worksheet _WorkSheet = _Sheets[sheetNo] as excel.Worksheet;
            return _WorkSheet;
        }

        /// <summary>
        /// 분모가 0이거나, 분자가 0일경우 0을 반환, 이외의 경우 나눈 몫을 반환
        /// </summary>
        /// <param name="textDenominator">분모값</param>
        /// <param name="textNumerator">분자값</param>
        /// <returns>분모가 0이거나, 분자가 0일경우 0을 반환, 이외의 경우 나눈 몫을 반환</returns>
        public static excel.Workbook GetExcel_WorkBook(string fileName)
        {
            //if (CommonUtil.openAsName != fileName)
            //{
            //    application = null;
            //    workBook = null;
            //    CommonUtil.openAsName = fileName;
            //}
            // 엑셀 프로세스 생성            
            if (application == null)
                application = new excel.ApplicationClass();
            //application.Visible = false;
            // 엑셀문서(워크북)를 읽기 전용으로 열기.
            //Microsoft.Office.Interop.Excel.Workbook _WorkBook = _Application.Workbooks.Open(문서경로,0,true,5,Missing.Value,Missing.Value,false,Missing.Value,
            if (workBook == null)
            {
                workBook = application.Workbooks.Open(fileName, 0, false, 5, Missing.Value, Missing.Value, false
                                       , Missing.Value, Missing.Value, true, false, Missing.Value, false, false, false);
            }
            return workBook;
        }

        public static excel.Worksheet GetExcelWorksheetForSimul(string fileName, int sheetNo)
        {
            // 엑셀문서(워크북)를 읽기 전용으로 열기.
            //Microsoft.Office.Interop.Excel.Workbook _WorkBook = _Application.Workbooks.Open(문서경로,0,true,5,Missing.Value,Missing.Value,false,Missing.Value,
            excel.Workbook _WorkBook = GetExcel_WorkBookForSimul(fileName);

            // sheets 생성
            excel.Sheets _Sheets = _WorkBook.Sheets;

            // 작업할 sheet 선택
            excel.Worksheet _WorkSheet = _Sheets[sheetNo] as excel.Worksheet;
            return _WorkSheet;
        }

        public static excel.Workbook GetExcel_WorkBookForSimul(string fileName)
        {
            if (CommonUtil.saveAsSimulName != fileName)
            {
                applicationForSimul = null;
                workBookForSimul = null;
                CommonUtil.saveAsSimulName = fileName;
            }
            // 엑셀 프로세스 생성            
            if (applicationForSimul == null)
                applicationForSimul = new excel.ApplicationClass();
            //application.Visible = false;
            // 엑셀문서(워크북)를 읽기 전용으로 열기.
            //Microsoft.Office.Interop.Excel.Workbook _WorkBook = _Application.Workbooks.Open(문서경로,0,true,5,Missing.Value,Missing.Value,false,Missing.Value,
            if (workBookForSimul == null)
            {
                workBookForSimul = applicationForSimul.Workbooks.Open(fileName, 0, false, 5, Missing.Value, Missing.Value, false
                                       , Missing.Value, Missing.Value, true, false, Missing.Value, false, false, false);
            }
            return workBookForSimul;
        }

        /// <summary>
        /// 분모가 0이거나, 분자가 0일경우 0을 반환, 이외의 경우 나눈 몫을 반환
        /// </summary>
        /// <param name="textDenominator">분모값</param>
        /// <param name="textNumerator">분자값</param>
        /// <returns>분모가 0이거나, 분자가 0일경우 0을 반환, 이외의 경우 나눈 몫을 반환</returns>
        public static void GetExcel_WorkBook_CLOSE(excel.Workbook excelWorkbook)
        {
            GetExcel_WorkBook_CLOSE();

            // 엑셀종료
        }

        public static void GetExcel_WorkBook_CLOSE()
        {
            if (workBook != null)
            {
                workBook.Close(true, Type.Missing, Type.Missing); //닫기 
                Marshal.ReleaseComObject(workBook);
                workBook = null;
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
            }

            if (workBookForSimul != null)
            {
                workBookForSimul.Close(true, Type.Missing, Type.Missing); //닫기 
                Marshal.ReleaseComObject(workBookForSimul);
                workBookForSimul = null;
                if (applicationForSimul != null)
                {
                    applicationForSimul.Quit();
                    applicationForSimul = null;
                }
            }
        }


        public static void clearTextBox(Control panel)
        {
            foreach (Control obj in panel.Controls)
            {
                if (obj is TextBox)
                {
                    (obj as TextBox).Text = "0";
                }
                else if (obj is RichTextBox)
                {
                    (obj as RichTextBox).Text = "0";
                }
                clearTextBox(obj);
            }
        }

        public static void ReadExcelFileToData(string file)
        {
            GetExcel_WorkBook(file);
            excel.Worksheet workSheet1 = workBook.Sheets[1] as excel.Worksheet;
            ReadExcelFileToDataBasicInput(workSheet1);
            ReadExcelFileToDataDetailInput(workSheet1);
            excel.Worksheet workSheet2 = workBook.Sheets[2] as excel.Worksheet;
            ReadExcelFileToDataResultBusiness(workSheet2);
            ReadExcelFileToDataResultStore(workSheet2);
            ReadExcelFileToDataResultFuture(workSheet2);
            excel.Worksheet workSheet3 = workBook.Sheets[3] as excel.Worksheet;
            ReadExcelFileToDataReport(workSheet3);
            GetExcel_WorkBook_CLOSE();
        }
        public static void WriteDataToExcelFile(string fileName, bool isSimul)
        {
            GetExcel_WorkBook(fileName);
            excel.Worksheet workSheet1 = workBook.Sheets[1] as excel.Worksheet;
            excel.Worksheet workSheet2 = workBook.Sheets[2] as excel.Worksheet;
            excel.Worksheet workSheet3 = workBook.Sheets[3] as excel.Worksheet;
            if (isSimul)
            {
                WriteDataToExcelFileDasicInput(workSheet1, CDataControl.g_SimBasicInput);
                WriteDataToExcelFileDetailInput(workSheet1, CDataControl.g_SimBasicInput, CDataControl.g_SimDetailInput);
                WriteExcelFileToDataResultBusiness(workSheet2, CDataControl.g_SimResultBusinessTotal, CDataControl.g_SimResultBusiness);
                WriteExcelFileToDataResultStore(workSheet2, CDataControl.g_SimResultStoreTotal, CDataControl.g_SimResultFutureTotal);
                WriteExcelFileToDataResultFuture(workSheet2, CDataControl.g_SimResultFutureTotal, CDataControl.g_SimResultFuture);
            }
            else
            {
                WriteDataToExcelFileDasicInput(workSheet1, CDataControl.g_BasicInput);
                WriteDataToExcelFileDetailInput(workSheet1, CDataControl.g_BasicInput, CDataControl.g_DetailInput);
                WriteExcelFileToDataResultBusiness(workSheet2, CDataControl.g_ResultBusinessTotal, CDataControl.g_ResultBusiness);
                WriteExcelFileToDataResultStore(workSheet2, CDataControl.g_ResultStoreTotal, CDataControl.g_ResultFutureTotal);
                WriteExcelFileToDataResultFuture(workSheet2, CDataControl.g_ResultFutureTotal, CDataControl.g_ResultFuture);
            }
            WriteExcelFileToDataReport(workSheet3);
            GetExcel_WorkBook_CLOSE();

        }

        public static void ReadFileManagerToData()
        {
            //관리자 파일을 읽어 넣는다
            try
            {
                string csv = System.IO.File.ReadAllText(defaultManagerFileName);
                csv = CommonUtil.Base64Decode(csv);
                string[] splitedCsv = csv.Split(',');
                CDataControl.g_BusinessAvg.setArrData_관리자데이터(splitedCsv);
            }
            catch (Exception ex)
            {
                String[] txtWrite2 = new String[34];
                // 파일이 없음
                for (int i = 0; i < txtWrite2.Length; i++)
                {
                    txtWrite2[i] = 0.ToString();
                }
                CDataControl.g_BusinessAvg.setArrData_관리자데이터(txtWrite2);
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="_WorkSheet"></param>
        public static void ReadExcelFileToDataBasicInput(excel.Worksheet _WorkSheet)
        {
            CDataControl.g_FileBasicInput.set지역(NullToEmpty(_WorkSheet.get_Range("C63", Type.Missing).Value2));
            CDataControl.g_FileBasicInput.set대리점(NullToEmpty(_WorkSheet.get_Range("E63", Type.Missing).Value2));
            CDataControl.g_FileBasicInput.set마케터(NullToEmpty(_WorkSheet.get_Range("G63", Type.Missing).Value2));
            CDataControl.g_ReportData.set지역(NullToEmpty(_WorkSheet.get_Range("C63", Type.Missing).Value2));
            CDataControl.g_ReportData.set대리점(NullToEmpty(_WorkSheet.get_Range("E63", Type.Missing).Value2));
            CDataControl.g_ReportData.set마케터(NullToEmpty(_WorkSheet.get_Range("G63", Type.Missing).Value2));

            //도매
            CDataControl.g_FileBasicInput.set도매_누적가입자수(NullToEmpty(_WorkSheet.get_Range("F7", Type.Missing).Value2));

            CDataControl.g_FileBasicInput.set도매_월평균판매대수_신규(NullToEmpty(_WorkSheet.get_Range("F8", Type.Missing).Value2));
            CDataControl.g_FileBasicInput.set도매_월평균판매대수_기변(NullToEmpty(_WorkSheet.get_Range("F9", Type.Missing).Value2));
            //CDataControl.g_FileBasicInput.set도매_월평균판매대수_소계(NullToEmpty(_WorkSheet.get_Range("F10", Type.Missing).Value2));

            CDataControl.g_FileBasicInput.set도매_월평균유통모델출고대수_LG(NullToEmpty(_WorkSheet.get_Range("F11", Type.Missing).Value2));
            CDataControl.g_FileBasicInput.set도매_월평균유통모델출고대수_SS(NullToEmpty(_WorkSheet.get_Range("F12", Type.Missing).Value2));
            //CDataControl.g_FileBasicInput.set도매_월평균유통모델출고대수_소계(NullToEmpty(_WorkSheet.get_Range("F13", Type.Missing).Value2));

            CDataControl.g_FileBasicInput.set도매_거래선수_개통사무실(NullToEmpty(_WorkSheet.get_Range("F14", Type.Missing).Value2));
            CDataControl.g_FileBasicInput.set도매_거래선수_판매점(NullToEmpty(_WorkSheet.get_Range("F16", Type.Missing).Value2));
            //CDataControl.g_FileBasicInput.set도매_거래선수_소계(NullToEmpty(_WorkSheet.get_Range("F17", Type.Missing).Value2));

            CDataControl.g_FileBasicInput.set도매_직원수_간부급(NullToEmpty(_WorkSheet.get_Range("F18", Type.Missing).Value2));
            CDataControl.g_FileBasicInput.set도매_직원수_평사원(NullToEmpty(_WorkSheet.get_Range("F19", Type.Missing).Value2));
            //CDataControl.g_FileBasicInput.set도매_직원수_소계(NullToEmpty(_WorkSheet.get_Range("F20", Type.Missing).Value2));

            //소매
            CDataControl.g_FileBasicInput.set소매_월평균판매대수_신규(NullToEmpty(_WorkSheet.get_Range("G8", Type.Missing).Value2));
            CDataControl.g_FileBasicInput.set소매_월평균판매대수_기변(NullToEmpty(_WorkSheet.get_Range("G9", Type.Missing).Value2));
            //CDataControl.g_FileBasicInput.set소매_월평균판매대수_소계(NullToEmpty(_WorkSheet.get_Range("G10", Type.Missing).Value2));

            CDataControl.g_FileBasicInput.set소매_거래선수_직영점(NullToEmpty(_WorkSheet.get_Range("G15", Type.Missing).Value2));
            //CDataControl.g_FileBasicInput.set소매_거래선수_소계(NullToEmpty(_WorkSheet.get_Range("G17", Type.Missing).Value2));

            CDataControl.g_FileBasicInput.set소매_직원수_간부급(NullToEmpty(_WorkSheet.get_Range("G18", Type.Missing).Value2));
            CDataControl.g_FileBasicInput.set소매_직원수_평사원(NullToEmpty(_WorkSheet.get_Range("G19", Type.Missing).Value2));
            //CDataControl.g_FileBasicInput.set소매_직원수_소계(NullToEmpty(_WorkSheet.get_Range("G20", Type.Missing).Value2));

            //합계
            //CDataControl.g_BasicInput.set누적가입자수_합계(_WorkSheet.get_Range("H7", Type.Missing).Value2.ToString());

            //CDataControl.g_BasicInput.set월평균판매대수_신규_합계(_WorkSheet.get_Range("H8", Type.Missing).Value2.ToString());
            //CDataControl.g_BasicInput.set월평균판매대수_기변_합계(_WorkSheet.get_Range("H9", Type.Missing).Value2.ToString());
            //CDataControl.g_BasicInput.set월평균판매대수_소계_합계(_WorkSheet.get_Range("H10", Type.Missing).Value2.ToString());

            //CDataControl.g_BasicInput.set월평균유통모델출고대수_LG_합계(_WorkSheet.get_Range("H11", Type.Missing).Value2.ToString());
            //CDataControl.g_BasicInput.set월평균유통모델출고대수_SS_합계(_WorkSheet.get_Range("H12", Type.Missing).Value2.ToString());
            //CDataControl.g_BasicInput.set월평균유통모델출고대수_소계_합계(_WorkSheet.get_Range("H13", Type.Missing).Value2.ToString());

            //CDataControl.g_BasicInput.set거래선수_개통사무실_합계(_WorkSheet.get_Range("H14", Type.Missing).Value2.ToString());
            //CDataControl.g_BasicInput.set거래선수_직영점_합계(_WorkSheet.get_Range("H15", Type.Missing).Value2.ToString());
            //CDataControl.g_BasicInput.set거래선수_판매점_합계(_WorkSheet.get_Range("H16", Type.Missing).Value2.ToString());
            //CDataControl.g_BasicInput.set거래선수_소계_합계(_WorkSheet.get_Range("H17", Type.Missing).Value2.ToString());

            //CDataControl.g_BasicInput.set직원수_간부급_합계(_WorkSheet.get_Range("H18", Type.Missing).Value2.ToString());
            //CDataControl.g_BasicInput.set직원수_평사원_합계(_WorkSheet.get_Range("H19", Type.Missing).Value2.ToString());
            //CDataControl.g_BasicInput.set직원수_소계_합계(_WorkSheet.get_Range("H20", Type.Missing).Value2.ToString());

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="_WorkSheet"></param>
        public static void ReadExcelFileToDataDetailInput(excel.Worksheet _WorkSheet)
        {
            //도매
            CDataControl.g_FileDetailInput.set도매_수익_월평균관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G26", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도매_수익_CS관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G27", Type.Missing).Value2)));//월총액
            CDataControl.g_FileDetailInput.set도매_수익_사업자모델매입관련추가수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G29", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도매_수익_유통모델매입관련추가수익_현금DC(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G30", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도매_수익_유통모델매입관련추가수익_VolumeDC(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G31", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도매_비용_대리점투자금액_신규(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G32", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도매_비용_대리점투자금액_기변(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G33", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도매_비용_직원급여_간부급(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G36", Type.Missing).Value2)));//월단위
            CDataControl.g_FileDetailInput.set도매_비용_직원급여_평사원(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G37", Type.Missing).Value2)));//월단위
            CDataControl.g_FileDetailInput.set도매_비용_지급임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G38", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도매_비용_운반비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G39", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도매_비용_차량유지비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G40", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도매_비용_지급수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G41", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도매_비용_판매촉진비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G42", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도매_비용_건물관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G43", Type.Missing).Value2)));

            CDataControl.g_FileDetailInput.set소매_수익_월평균업무취급수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G44", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set소매_수익_직영매장판매수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G45", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set소매_비용_직원급여_간부급(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G48", Type.Missing).Value2)));//월단위
            CDataControl.g_FileDetailInput.set소매_비용_직원급여_평사원(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G49", Type.Missing).Value2)));//월단위
            CDataControl.g_FileDetailInput.set소매_비용_지급임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G50", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set소매_비용_지급수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G51", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set소매_비용_판매촉진비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G52", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set소매_비용_건물관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G53", Type.Missing).Value2)));

            CDataControl.g_FileDetailInput.set도소매_비용_복리후생비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G54", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도소매_비용_통신비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G55", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도소매_비용_공과금(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G56", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도소매_비용_소모품비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G57", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도소매_비용_이자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G58", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도소매_비용_부가세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G59", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도소매_비용_법인세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G60", Type.Missing).Value2)));
            CDataControl.g_FileDetailInput.set도소매_비용_기타(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("G61", Type.Missing).Value2)));

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="_WorkSheet"></param>
        public static void ReadExcelFileToDataResultBusiness(excel.Worksheet _WorkSheet)
        {
            CDataControl.g_FileResultBusinessTotal.set전체_수익_가입자관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D7", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set전체_수익_CS관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D8", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set전체_수익_업무취급수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D9", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set전체_수익_사업자모델매입에따른추가수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D10", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set전체_수익_유통모델매입에따른추가수익_현금_Volume(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D11", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set전체_수익_직영매장판매수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D12", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.전체_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D13", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set전체_비용_대리점투자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D14", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set전체_비용_인건비_급여_복리후생비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D15", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set전체_비용_임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D16", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set전체_비용_이자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D17", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set전체_비용_부가세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D18", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set전체_비용_법인세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D19", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set전체_비용_기타판매관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D20", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.전체_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D21", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.전체손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D22", Type.Missing).Value2));

            CDataControl.g_FileResultBusinessTotal.set도매_수익_가입자관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D28", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set도매_수익_CS관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D29", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set도매_수익_사업자모델매입에따른추가수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D30", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set도매_수익_유통모델매입에따른추가수익_현금_Volume(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D31", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.도매_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D32", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set도매_비용_대리점투자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D33", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set도매_비용_인건비_급여_복리후생비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D34", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set도매_비용_임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D35", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set도매_비용_이자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D36", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set도매_비용_부가세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D37", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set도매_비용_법인세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D38", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set도매_비용_기타판매관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D39", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.도매_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D40", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.도매손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D41", Type.Missing).Value2));

            CDataControl.g_FileResultBusinessTotal.set소매_수익_업무취급수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D46", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set소매_수익_직영매장판매수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D47", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.소매_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D48", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set소매_비용_인건비_급여_복리후생비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D49", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set소매_비용_임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D50", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set소매_비용_이자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D51", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set소매_비용_부가세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D52", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set소매_비용_법인세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D53", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.set소매_비용_기타판매관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D54", Type.Missing).Value2)));
            CDataControl.g_FileResultBusinessTotal.소매_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D55", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.소매손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D56", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.점별손익추정 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("D57", Type.Missing).Value2));



            CDataControl.g_FileResultBusiness.set전체_수익_가입자관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E7", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set전체_수익_CS관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E8", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set전체_수익_업무취급수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E9", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set전체_수익_사업자모델매입에따른추가수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E10", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set전체_수익_유통모델매입에따른추가수익_현금_Volume(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E11", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set전체_수익_직영매장판매수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E12", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.전체_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E13", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set전체_비용_대리점투자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E14", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set전체_비용_인건비_급여_복리후생비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E15", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set전체_비용_임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E16", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set전체_비용_이자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E17", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set전체_비용_부가세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E18", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set전체_비용_법인세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E19", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set전체_비용_기타판매관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E20", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.전체_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E21", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.전체손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E22", Type.Missing).Value2));

            CDataControl.g_FileResultBusiness.set도매_수익_가입자관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E28", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set도매_수익_CS관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E29", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set도매_수익_사업자모델매입에따른추가수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E30", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set도매_수익_유통모델매입에따른추가수익_현금_Volume(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E31", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.도매_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E32", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set도매_비용_대리점투자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E33", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set도매_비용_인건비_급여_복리후생비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E34", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set도매_비용_임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E35", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set도매_비용_이자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E36", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set도매_비용_부가세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E37", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set도매_비용_법인세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E38", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set도매_비용_기타판매관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E39", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.도매_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E40", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.도매손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E41", Type.Missing).Value2));

            CDataControl.g_FileResultBusiness.set소매_수익_업무취급수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E46", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set소매_수익_직영매장판매수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E47", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.소매_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E48", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set소매_비용_인건비_급여_복리후생비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E49", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set소매_비용_임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E50", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set소매_비용_이자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E51", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set소매_비용_부가세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E52", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set소매_비용_법인세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E53", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.set소매_비용_기타판매관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E54", Type.Missing).Value2)));
            CDataControl.g_FileResultBusiness.소매_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E55", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.소매손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E56", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.점별손익추정 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("E57", Type.Missing).Value2));

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="_WorkSheet"></param>
        public static void ReadExcelFileToDataResultStore(excel.Worksheet _WorkSheet)
        {
            //데이터 저장
            CDataControl.g_FileResultStoreTotal.set전체_수익_가입자관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I7", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set전체_수익_CS관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I8", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set전체_수익_업무취급수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I9", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set전체_수익_사업자모델매입에따른추가수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I10", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set전체_수익_유통모델매입에따른추가수익_현금_Volume(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I11", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set전체_수익_직영매장판매수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I12", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.전체_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I13", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set전체_비용_대리점투자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I14", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set전체_비용_인건비_급여_복리후생비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I15", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set전체_비용_임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I16", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set전체_비용_이자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I17", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set전체_비용_부가세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I18", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set전체_비용_법인세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I19", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set전체_비용_기타판매관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I20", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.전체_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I21", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.전체손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I22", Type.Missing).Value2));

            CDataControl.g_FileResultStoreTotal.set도매_수익_가입자관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I28", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set도매_수익_CS관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I29", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set도매_수익_사업자모델매입에따른추가수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I30", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set도매_수익_유통모델매입에따른추가수익_현금_Volume(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I31", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.도매_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I32", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set도매_비용_대리점투자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I33", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set도매_비용_인건비_급여_복리후생비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I34", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set도매_비용_임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I35", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set도매_비용_이자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I36", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set도매_비용_부가세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I37", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set도매_비용_법인세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I38", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set도매_비용_기타판매관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I39", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.도매_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I40", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.도매손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I41", Type.Missing).Value2));

            CDataControl.g_FileResultStoreTotal.set소매_수익_업무취급수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I46", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set소매_수익_직영매장판매수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I47", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.소매_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I48", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set소매_비용_인건비_급여_복리후생비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I49", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set소매_비용_임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I50", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set소매_비용_이자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I51", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set소매_비용_부가세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I52", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set소매_비용_법인세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I53", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.set소매_비용_기타판매관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I54", Type.Missing).Value2)));
            CDataControl.g_FileResultStoreTotal.소매_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I55", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.소매손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I56", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.점별손익추정 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("I57", Type.Missing).Value2));



            CDataControl.g_FileResultStore.set전체_수익_가입자관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J7", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set전체_수익_CS관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J8", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set전체_수익_업무취급수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J9", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set전체_수익_사업자모델매입에따른추가수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J10", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set전체_수익_유통모델매입에따른추가수익_현금_Volume(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J11", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set전체_수익_직영매장판매수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J12", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.전체_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J13", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set전체_비용_대리점투자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J14", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set전체_비용_인건비_급여_복리후생비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J15", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set전체_비용_임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J16", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set전체_비용_이자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J17", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set전체_비용_부가세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J18", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set전체_비용_법인세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J19", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set전체_비용_기타판매관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J20", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.전체_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J21", Type.Missing).Value2));
            CDataControl.g_FileResultStore.전체손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J22", Type.Missing).Value2));

            CDataControl.g_FileResultStore.set도매_수익_가입자관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J28", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set도매_수익_CS관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J29", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set도매_수익_사업자모델매입에따른추가수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J30", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set도매_수익_유통모델매입에따른추가수익_현금_Volume(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J31", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.도매_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J32", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set도매_비용_대리점투자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J33", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set도매_비용_인건비_급여_복리후생비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J34", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set도매_비용_임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J35", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set도매_비용_이자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J36", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set도매_비용_부가세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J37", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set도매_비용_법인세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J38", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set도매_비용_기타판매관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J39", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.도매_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J40", Type.Missing).Value2));
            CDataControl.g_FileResultStore.도매손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J41", Type.Missing).Value2));

            CDataControl.g_FileResultStore.set소매_수익_업무취급수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J46", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set소매_수익_직영매장판매수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J47", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.소매_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J48", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set소매_비용_인건비_급여_복리후생비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J49", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set소매_비용_임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J50", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set소매_비용_이자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J51", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set소매_비용_부가세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J52", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set소매_비용_법인세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J53", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.set소매_비용_기타판매관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J54", Type.Missing).Value2)));
            CDataControl.g_FileResultStore.소매_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J55", Type.Missing).Value2));
            CDataControl.g_FileResultStore.소매손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J56", Type.Missing).Value2));
            CDataControl.g_FileResultStore.점별손익추정 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("J57", Type.Missing).Value2));

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="_WorkSheet"></param>
        public static void ReadExcelFileToDataResultFuture(excel.Worksheet _WorkSheet)
        {
            CDataControl.g_FileResultFutureTotal.set전체_수익_가입자관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("N7", Type.Missing).Value2)));
            CDataControl.g_FileResultFutureTotal.set전체_수익_CS관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("N8", Type.Missing).Value2)));
            CDataControl.g_FileResultFutureTotal.set전체_수익_업무취급수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("N9", Type.Missing).Value2)));
            CDataControl.g_FileResultFutureTotal.set전체_수익_사업자모델매입에따른추가수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("N10", Type.Missing).Value2)));
            CDataControl.g_FileResultFutureTotal.set전체_수익_유통모델매입에따른추가수익_현금_Volume(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("N11", Type.Missing).Value2)));
            CDataControl.g_FileResultFutureTotal.set전체_수익_직영매장판매수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("N12", Type.Missing).Value2)));
            CDataControl.g_FileResultFutureTotal.전체_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("N13", Type.Missing).Value2));
            CDataControl.g_FileResultFutureTotal.set전체_비용_대리점투자비용(_WorkSheet.get_Range("N14", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set전체_비용_인건비_급여_복리후생비(_WorkSheet.get_Range("N15", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set전체_비용_임차료(_WorkSheet.get_Range("N16", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set전체_비용_이자비용(_WorkSheet.get_Range("N17", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set전체_비용_부가세(_WorkSheet.get_Range("N18", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set전체_비용_법인세(_WorkSheet.get_Range("N19", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set전체_비용_기타판매관리비(_WorkSheet.get_Range("N20", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.전체_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("N21", Type.Missing).Value2.ToString()));
            CDataControl.g_FileResultFutureTotal.전체손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("N22", Type.Missing).Value2.ToString()));

            CDataControl.g_FileResultFutureTotal.set도매_수익_가입자관리수수료(_WorkSheet.get_Range("N28", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_수익_CS관리수수료(_WorkSheet.get_Range("N29", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_수익_사업자모델매입에따른추가수익(_WorkSheet.get_Range("N30", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_수익_유통모델매입에따른추가수익_현금_Volume(_WorkSheet.get_Range("N31", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.도매_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("N32", Type.Missing).Value2.ToString()));
            CDataControl.g_FileResultFutureTotal.set도매_비용_대리점투자비용(_WorkSheet.get_Range("N33", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_비용_인건비_급여_복리후생비(_WorkSheet.get_Range("N34", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_비용_임차료(_WorkSheet.get_Range("N35", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_비용_이자비용(_WorkSheet.get_Range("N36", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_비용_부가세(_WorkSheet.get_Range("N37", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_비용_법인세(_WorkSheet.get_Range("N38", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_비용_기타판매관리비(_WorkSheet.get_Range("N39", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.도매_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("N40", Type.Missing).Value2.ToString()));
            CDataControl.g_FileResultFutureTotal.도매손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("N41", Type.Missing).Value2.ToString()));

            CDataControl.g_FileResultFutureTotal.set소매_수익_업무취급수수료(_WorkSheet.get_Range("N46", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set소매_수익_직영매장판매수익(_WorkSheet.get_Range("N47", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.소매_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("N48", Type.Missing).Value2.ToString()));
            CDataControl.g_FileResultFutureTotal.set소매_비용_인건비_급여_복리후생비(_WorkSheet.get_Range("N49", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set소매_비용_임차료(_WorkSheet.get_Range("N50", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set소매_비용_이자비용(_WorkSheet.get_Range("N51", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set소매_비용_부가세(_WorkSheet.get_Range("N52", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set소매_비용_법인세(_WorkSheet.get_Range("N53", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set소매_비용_기타판매관리비(_WorkSheet.get_Range("N54", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.소매_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("N55", Type.Missing).Value2.ToString()));
            CDataControl.g_FileResultFutureTotal.소매손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("N56", Type.Missing).Value2.ToString()));
            CDataControl.g_FileResultFutureTotal.점별손익추정 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("N57", Type.Missing).Value2.ToString()));



            CDataControl.g_FileResultFuture.set전체_수익_가입자관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O7", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set전체_수익_CS관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O8", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set전체_수익_업무취급수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O9", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set전체_수익_사업자모델매입에따른추가수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O10", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set전체_수익_유통모델매입에따른추가수익_현금_Volume(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O11", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set전체_수익_직영매장판매수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O12", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.전체_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O13", Type.Missing).Value2));
            CDataControl.g_FileResultFuture.set전체_비용_대리점투자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O14", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set전체_비용_인건비_급여_복리후생비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O15", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set전체_비용_임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O16", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set전체_비용_이자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O17", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set전체_비용_부가세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O18", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set전체_비용_법인세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O19", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set전체_비용_기타판매관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O20", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.전체_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O21", Type.Missing).Value2));
            CDataControl.g_FileResultFuture.전체손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O22", Type.Missing).Value2));

            CDataControl.g_FileResultFuture.set도매_수익_가입자관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O28", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set도매_수익_CS관리수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O29", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set도매_수익_사업자모델매입에따른추가수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O30", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set도매_수익_유통모델매입에따른추가수익_현금_Volume(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O31", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.도매_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O32", Type.Missing).Value2));
            CDataControl.g_FileResultFuture.set도매_비용_대리점투자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O33", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set도매_비용_인건비_급여_복리후생비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O34", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set도매_비용_임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O35", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set도매_비용_이자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O36", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set도매_비용_부가세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O37", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set도매_비용_법인세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O38", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set도매_비용_기타판매관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O39", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.도매_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O40", Type.Missing).Value2));
            CDataControl.g_FileResultFuture.도매손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O41", Type.Missing).Value2));

            CDataControl.g_FileResultFuture.set소매_수익_업무취급수수료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O46", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set소매_수익_직영매장판매수익(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O47", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.소매_수익_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O48", Type.Missing).Value2));
            CDataControl.g_FileResultFuture.set소매_비용_인건비_급여_복리후생비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O49", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set소매_비용_임차료(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O50", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set소매_비용_이자비용(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O51", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set소매_비용_부가세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O52", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set소매_비용_법인세(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O53", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.set소매_비용_기타판매관리비(StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O54", Type.Missing).Value2)));
            CDataControl.g_FileResultFuture.소매_비용_소계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O55", Type.Missing).Value2));
            CDataControl.g_FileResultFuture.소매손익계 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O56", Type.Missing).Value2));
            CDataControl.g_FileResultFuture.점별손익추정 = StringToDoubleVal(NullToEmpty(_WorkSheet.get_Range("O57", Type.Missing).Value2));

        }
        private static void ReadExcelFileToDataReport(excel.Worksheet workSheet3)
        {
            object 분석내용 = workSheet3.get_Range("B4", Type.Missing).Value2;
            object 지원활동 = workSheet3.get_Range("C4", Type.Missing).Value2;
            object 배경및이슈 = workSheet3.get_Range("D4", Type.Missing).Value2;
            CDataControl.g_ReportData.set분석내용_및_대리점_활동방향(분석내용==null?"":분석내용.ToString());
            CDataControl.g_ReportData.setLG_지원_활동(지원활동 == null ? "" : 지원활동.ToString());
            CDataControl.g_ReportData.set배경_및_이슈(배경및이슈 == null ? "" : 배경및이슈.ToString());
        }






        /// <summary>
        /// 
        /// </summary>
        /// <param name="_WorkSheet"></param>
        public static void WriteExcelFileToDataResultBusiness(excel.Worksheet _WorkSheet, CResultData businessTotal, CResultData business)
        {
            _WorkSheet.get_Range("D7", Type.Missing).Value2 = businessTotal.전체_수익_가입자관리수수료;
            _WorkSheet.get_Range("D8", Type.Missing).Value2 = businessTotal.전체_수익_CS관리수수료;
            _WorkSheet.get_Range("D9", Type.Missing).Value2 = businessTotal.전체_수익_업무취급수수료;
            _WorkSheet.get_Range("D10", Type.Missing).Value2 = businessTotal.전체_수익_사업자모델매입에따른추가수익;
            _WorkSheet.get_Range("D11", Type.Missing).Value2 = businessTotal.전체_수익_유통모델매입에따른추가수익_현금_Volume;
            _WorkSheet.get_Range("D12", Type.Missing).Value2 = businessTotal.전체_수익_직영매장판매수익;
            _WorkSheet.get_Range("D13", Type.Missing).Value2 = businessTotal.전체_수익_소계;
            _WorkSheet.get_Range("D14", Type.Missing).Value2 = businessTotal.get전체_비용_대리점투자비용();
            _WorkSheet.get_Range("D15", Type.Missing).Value2 = businessTotal.get전체_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("D16", Type.Missing).Value2 = businessTotal.get전체_비용_임차료();
            _WorkSheet.get_Range("D17", Type.Missing).Value2 = businessTotal.get전체_비용_이자비용();
            _WorkSheet.get_Range("D18", Type.Missing).Value2 = businessTotal.get전체_비용_부가세();
            _WorkSheet.get_Range("D19", Type.Missing).Value2 = businessTotal.get전체_비용_법인세();
            _WorkSheet.get_Range("D20", Type.Missing).Value2 = businessTotal.get전체_비용_기타판매관리비();
            _WorkSheet.get_Range("D21", Type.Missing).Value2 = businessTotal.전체_비용_소계;
            _WorkSheet.get_Range("D22", Type.Missing).Value2 = businessTotal.전체손익계;

            _WorkSheet.get_Range("D28", Type.Missing).Value2 = businessTotal.get도매_수익_가입자관리수수료();
            _WorkSheet.get_Range("D29", Type.Missing).Value2 = businessTotal.get도매_수익_CS관리수수료();
            _WorkSheet.get_Range("D30", Type.Missing).Value2 = businessTotal.get도매_수익_사업자모델매입에따른추가수익();
            _WorkSheet.get_Range("D31", Type.Missing).Value2 = businessTotal.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
            _WorkSheet.get_Range("D32", Type.Missing).Value2 = businessTotal.도매_수익_소계;
            _WorkSheet.get_Range("D33", Type.Missing).Value2 = businessTotal.get도매_비용_대리점투자비용();
            _WorkSheet.get_Range("D34", Type.Missing).Value2 = businessTotal.get도매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("D35", Type.Missing).Value2 = businessTotal.get도매_비용_임차료();
            _WorkSheet.get_Range("D36", Type.Missing).Value2 = businessTotal.get도매_비용_이자비용();
            _WorkSheet.get_Range("D37", Type.Missing).Value2 = businessTotal.get도매_비용_부가세();
            _WorkSheet.get_Range("D38", Type.Missing).Value2 = businessTotal.get도매_비용_법인세();
            _WorkSheet.get_Range("D39", Type.Missing).Value2 = businessTotal.get도매_비용_기타판매관리비();
            _WorkSheet.get_Range("D40", Type.Missing).Value2 = businessTotal.도매_비용_소계;
            _WorkSheet.get_Range("D41", Type.Missing).Value2 = businessTotal.도매손익계;

            _WorkSheet.get_Range("D46", Type.Missing).Value2 = businessTotal.get소매_수익_업무취급수수료();
            _WorkSheet.get_Range("D47", Type.Missing).Value2 = businessTotal.get소매_수익_직영매장판매수익();
            _WorkSheet.get_Range("D48", Type.Missing).Value2 = businessTotal.소매_수익_소계;
            _WorkSheet.get_Range("D49", Type.Missing).Value2 = businessTotal.get소매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("D50", Type.Missing).Value2 = businessTotal.get소매_비용_임차료();
            _WorkSheet.get_Range("D51", Type.Missing).Value2 = businessTotal.get소매_비용_이자비용();
            _WorkSheet.get_Range("D52", Type.Missing).Value2 = businessTotal.get소매_비용_부가세();
            _WorkSheet.get_Range("D53", Type.Missing).Value2 = businessTotal.get소매_비용_법인세();
            _WorkSheet.get_Range("D54", Type.Missing).Value2 = businessTotal.get소매_비용_기타판매관리비();
            _WorkSheet.get_Range("D55", Type.Missing).Value2 = businessTotal.소매_비용_소계;
            _WorkSheet.get_Range("D56", Type.Missing).Value2 = businessTotal.소매손익계;
            _WorkSheet.get_Range("D57", Type.Missing).Value2 = businessTotal.점별손익추정;



            _WorkSheet.get_Range("E7", Type.Missing).Value2 = business.전체_수익_가입자관리수수료;
            _WorkSheet.get_Range("E8", Type.Missing).Value2 = business.전체_수익_CS관리수수료;
            _WorkSheet.get_Range("E9", Type.Missing).Value2 = business.전체_수익_업무취급수수료;
            _WorkSheet.get_Range("E10", Type.Missing).Value2 = business.전체_수익_사업자모델매입에따른추가수익;
            _WorkSheet.get_Range("E11", Type.Missing).Value2 = business.전체_수익_유통모델매입에따른추가수익_현금_Volume;
            _WorkSheet.get_Range("E12", Type.Missing).Value2 = business.전체_수익_직영매장판매수익;
            _WorkSheet.get_Range("E13", Type.Missing).Value2 = business.전체_수익_소계;
            _WorkSheet.get_Range("E14", Type.Missing).Value2 = business.get전체_비용_대리점투자비용();
            _WorkSheet.get_Range("E15", Type.Missing).Value2 = business.get전체_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("E16", Type.Missing).Value2 = business.get전체_비용_임차료();
            _WorkSheet.get_Range("E17", Type.Missing).Value2 = business.get전체_비용_이자비용();
            _WorkSheet.get_Range("E18", Type.Missing).Value2 = business.get전체_비용_부가세();
            _WorkSheet.get_Range("E19", Type.Missing).Value2 = business.get전체_비용_법인세();
            _WorkSheet.get_Range("E20", Type.Missing).Value2 = business.get전체_비용_기타판매관리비();
            _WorkSheet.get_Range("E21", Type.Missing).Value2 = business.전체_비용_소계;
            _WorkSheet.get_Range("E22", Type.Missing).Value2 = business.전체손익계;

            _WorkSheet.get_Range("E28", Type.Missing).Value2 = business.get도매_수익_가입자관리수수료();
            _WorkSheet.get_Range("E29", Type.Missing).Value2 = business.get도매_수익_CS관리수수료();
            _WorkSheet.get_Range("E30", Type.Missing).Value2 = business.get도매_수익_사업자모델매입에따른추가수익();
            _WorkSheet.get_Range("E31", Type.Missing).Value2 = business.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
            _WorkSheet.get_Range("E32", Type.Missing).Value2 = business.도매_수익_소계;
            _WorkSheet.get_Range("E33", Type.Missing).Value2 = business.get도매_비용_대리점투자비용();
            _WorkSheet.get_Range("E34", Type.Missing).Value2 = business.get도매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("E35", Type.Missing).Value2 = business.get도매_비용_임차료();
            _WorkSheet.get_Range("E36", Type.Missing).Value2 = business.get도매_비용_이자비용();
            _WorkSheet.get_Range("E37", Type.Missing).Value2 = business.get도매_비용_부가세();
            _WorkSheet.get_Range("E38", Type.Missing).Value2 = business.get도매_비용_법인세();
            _WorkSheet.get_Range("E39", Type.Missing).Value2 = business.get도매_비용_기타판매관리비();
            _WorkSheet.get_Range("E40", Type.Missing).Value2 = business.도매_비용_소계;
            _WorkSheet.get_Range("E41", Type.Missing).Value2 = business.도매손익계;

            _WorkSheet.get_Range("E46", Type.Missing).Value2 = business.get소매_수익_업무취급수수료();
            _WorkSheet.get_Range("E47", Type.Missing).Value2 = business.get소매_수익_직영매장판매수익();
            _WorkSheet.get_Range("E48", Type.Missing).Value2 = business.소매_수익_소계;
            _WorkSheet.get_Range("E49", Type.Missing).Value2 = business.get소매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("E50", Type.Missing).Value2 = business.get소매_비용_임차료();
            _WorkSheet.get_Range("E51", Type.Missing).Value2 = business.get소매_비용_이자비용();
            _WorkSheet.get_Range("E52", Type.Missing).Value2 = business.get소매_비용_부가세();
            _WorkSheet.get_Range("E53", Type.Missing).Value2 = business.get소매_비용_법인세();
            _WorkSheet.get_Range("E54", Type.Missing).Value2 = business.get소매_비용_기타판매관리비();
            _WorkSheet.get_Range("E55", Type.Missing).Value2 = business.소매_비용_소계;
            _WorkSheet.get_Range("E56", Type.Missing).Value2 = business.소매손익계;
            _WorkSheet.get_Range("E57", Type.Missing).Value2 = business.점별손익추정;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="_WorkSheet"></param>
        public static void WriteExcelFileToDataResultStore(excel.Worksheet _WorkSheet, CResultData storeTotal, CResultData store)
        {
            _WorkSheet.get_Range("I7", Type.Missing).Value2 = storeTotal.전체_수익_가입자관리수수료;
            _WorkSheet.get_Range("I8", Type.Missing).Value2 = storeTotal.전체_수익_CS관리수수료;
            _WorkSheet.get_Range("I9", Type.Missing).Value2 = storeTotal.전체_수익_업무취급수수료;
            _WorkSheet.get_Range("I10", Type.Missing).Value2 = storeTotal.전체_수익_사업자모델매입에따른추가수익;
            _WorkSheet.get_Range("I11", Type.Missing).Value2 = storeTotal.전체_수익_유통모델매입에따른추가수익_현금_Volume;
            _WorkSheet.get_Range("I12", Type.Missing).Value2 = storeTotal.전체_수익_직영매장판매수익;
            _WorkSheet.get_Range("I13", Type.Missing).Value2 = storeTotal.전체_수익_소계;
            _WorkSheet.get_Range("I14", Type.Missing).Value2 = storeTotal.get전체_비용_대리점투자비용();
            _WorkSheet.get_Range("I15", Type.Missing).Value2 = storeTotal.get전체_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("I16", Type.Missing).Value2 = storeTotal.get전체_비용_임차료();
            _WorkSheet.get_Range("I17", Type.Missing).Value2 = storeTotal.get전체_비용_이자비용();
            _WorkSheet.get_Range("I18", Type.Missing).Value2 = storeTotal.get전체_비용_부가세();
            _WorkSheet.get_Range("I19", Type.Missing).Value2 = storeTotal.get전체_비용_법인세();
            _WorkSheet.get_Range("I20", Type.Missing).Value2 = storeTotal.get전체_비용_기타판매관리비();
            _WorkSheet.get_Range("I21", Type.Missing).Value2 = storeTotal.전체_비용_소계;
            _WorkSheet.get_Range("I22", Type.Missing).Value2 = storeTotal.전체손익계;

            _WorkSheet.get_Range("I28", Type.Missing).Value2 = storeTotal.get도매_수익_가입자관리수수료();
            _WorkSheet.get_Range("I29", Type.Missing).Value2 = storeTotal.get도매_수익_CS관리수수료();
            _WorkSheet.get_Range("I30", Type.Missing).Value2 = storeTotal.get도매_수익_사업자모델매입에따른추가수익();
            _WorkSheet.get_Range("I31", Type.Missing).Value2 = storeTotal.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
            _WorkSheet.get_Range("I32", Type.Missing).Value2 = storeTotal.도매_수익_소계;
            _WorkSheet.get_Range("I33", Type.Missing).Value2 = storeTotal.get도매_비용_대리점투자비용();
            _WorkSheet.get_Range("I34", Type.Missing).Value2 = storeTotal.get도매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("I35", Type.Missing).Value2 = storeTotal.get도매_비용_임차료();
            _WorkSheet.get_Range("I36", Type.Missing).Value2 = storeTotal.get도매_비용_이자비용();
            _WorkSheet.get_Range("I37", Type.Missing).Value2 = storeTotal.get도매_비용_부가세();
            _WorkSheet.get_Range("I38", Type.Missing).Value2 = storeTotal.get도매_비용_법인세();
            _WorkSheet.get_Range("I39", Type.Missing).Value2 = storeTotal.get도매_비용_기타판매관리비();
            _WorkSheet.get_Range("I40", Type.Missing).Value2 = storeTotal.도매_비용_소계;
            _WorkSheet.get_Range("I41", Type.Missing).Value2 = storeTotal.도매손익계;

            _WorkSheet.get_Range("I46", Type.Missing).Value2 = storeTotal.get소매_수익_업무취급수수료();
            _WorkSheet.get_Range("I47", Type.Missing).Value2 = storeTotal.get소매_수익_직영매장판매수익();
            _WorkSheet.get_Range("I48", Type.Missing).Value2 = storeTotal.소매_수익_소계;
            _WorkSheet.get_Range("I49", Type.Missing).Value2 = storeTotal.get소매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("I50", Type.Missing).Value2 = storeTotal.get소매_비용_임차료();
            _WorkSheet.get_Range("I51", Type.Missing).Value2 = storeTotal.get소매_비용_이자비용();
            _WorkSheet.get_Range("I52", Type.Missing).Value2 = storeTotal.get소매_비용_부가세();
            _WorkSheet.get_Range("I53", Type.Missing).Value2 = storeTotal.get소매_비용_법인세();
            _WorkSheet.get_Range("I54", Type.Missing).Value2 = storeTotal.get소매_비용_기타판매관리비();
            _WorkSheet.get_Range("I55", Type.Missing).Value2 = storeTotal.소매_비용_소계;
            _WorkSheet.get_Range("I56", Type.Missing).Value2 = storeTotal.소매손익계;
            _WorkSheet.get_Range("I57", Type.Missing).Value2 = storeTotal.점별손익추정;



            _WorkSheet.get_Range("J7", Type.Missing).Value2 = store.전체_수익_가입자관리수수료;
            _WorkSheet.get_Range("J8", Type.Missing).Value2 = store.전체_수익_CS관리수수료;
            _WorkSheet.get_Range("J9", Type.Missing).Value2 = store.전체_수익_업무취급수수료;
            _WorkSheet.get_Range("J10", Type.Missing).Value2 = store.전체_수익_사업자모델매입에따른추가수익;
            _WorkSheet.get_Range("J11", Type.Missing).Value2 = store.전체_수익_유통모델매입에따른추가수익_현금_Volume;
            _WorkSheet.get_Range("J12", Type.Missing).Value2 = store.전체_수익_직영매장판매수익;
            _WorkSheet.get_Range("J13", Type.Missing).Value2 = store.전체_수익_소계;
            _WorkSheet.get_Range("J14", Type.Missing).Value2 = store.get전체_비용_대리점투자비용();
            _WorkSheet.get_Range("J15", Type.Missing).Value2 = store.get전체_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("J16", Type.Missing).Value2 = store.get전체_비용_임차료();
            _WorkSheet.get_Range("J17", Type.Missing).Value2 = store.get전체_비용_이자비용();
            _WorkSheet.get_Range("J18", Type.Missing).Value2 = store.get전체_비용_부가세();
            _WorkSheet.get_Range("J19", Type.Missing).Value2 = store.get전체_비용_법인세();
            _WorkSheet.get_Range("J20", Type.Missing).Value2 = store.get전체_비용_기타판매관리비();
            _WorkSheet.get_Range("J21", Type.Missing).Value2 = store.전체_비용_소계;
            _WorkSheet.get_Range("J22", Type.Missing).Value2 = store.전체손익계;

            _WorkSheet.get_Range("J28", Type.Missing).Value2 = store.get도매_수익_가입자관리수수료();
            _WorkSheet.get_Range("J29", Type.Missing).Value2 = store.get도매_수익_CS관리수수료();
            _WorkSheet.get_Range("J30", Type.Missing).Value2 = store.get도매_수익_사업자모델매입에따른추가수익();
            _WorkSheet.get_Range("J31", Type.Missing).Value2 = store.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
            _WorkSheet.get_Range("J32", Type.Missing).Value2 = store.도매_수익_소계;
            _WorkSheet.get_Range("J33", Type.Missing).Value2 = store.get도매_비용_대리점투자비용();
            _WorkSheet.get_Range("J34", Type.Missing).Value2 = store.get도매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("J35", Type.Missing).Value2 = store.get도매_비용_임차료();
            _WorkSheet.get_Range("J36", Type.Missing).Value2 = store.get도매_비용_이자비용();
            _WorkSheet.get_Range("J37", Type.Missing).Value2 = store.get도매_비용_부가세();
            _WorkSheet.get_Range("J38", Type.Missing).Value2 = store.get도매_비용_법인세();
            _WorkSheet.get_Range("J39", Type.Missing).Value2 = store.get도매_비용_기타판매관리비();
            _WorkSheet.get_Range("J40", Type.Missing).Value2 = store.도매_비용_소계;
            _WorkSheet.get_Range("J41", Type.Missing).Value2 = store.도매손익계;

            _WorkSheet.get_Range("J46", Type.Missing).Value2 = store.get소매_수익_업무취급수수료();
            _WorkSheet.get_Range("J47", Type.Missing).Value2 = store.get소매_수익_직영매장판매수익();
            _WorkSheet.get_Range("J48", Type.Missing).Value2 = store.소매_수익_소계;
            _WorkSheet.get_Range("J49", Type.Missing).Value2 = store.get소매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("J50", Type.Missing).Value2 = store.get소매_비용_임차료();
            _WorkSheet.get_Range("J51", Type.Missing).Value2 = store.get소매_비용_이자비용();
            _WorkSheet.get_Range("J52", Type.Missing).Value2 = store.get소매_비용_부가세();
            _WorkSheet.get_Range("J53", Type.Missing).Value2 = store.get소매_비용_법인세();
            _WorkSheet.get_Range("J54", Type.Missing).Value2 = store.get소매_비용_기타판매관리비();
            _WorkSheet.get_Range("J55", Type.Missing).Value2 = store.소매_비용_소계;
            _WorkSheet.get_Range("J56", Type.Missing).Value2 = store.소매손익계;
            _WorkSheet.get_Range("J57", Type.Missing).Value2 = store.점별손익추정;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="_WorkSheet"></param>
        public static void WriteExcelFileToDataResultFuture(excel.Worksheet _WorkSheet, CResultData futureTotal, CResultData future)
        {
            _WorkSheet.get_Range("N7", Type.Missing).Value2 = futureTotal.전체_수익_가입자관리수수료;
            _WorkSheet.get_Range("N8", Type.Missing).Value2 = futureTotal.전체_수익_CS관리수수료;
            _WorkSheet.get_Range("N9", Type.Missing).Value2 = futureTotal.전체_수익_업무취급수수료;
            _WorkSheet.get_Range("N10", Type.Missing).Value2 = futureTotal.전체_수익_사업자모델매입에따른추가수익;
            _WorkSheet.get_Range("N11", Type.Missing).Value2 = futureTotal.전체_수익_유통모델매입에따른추가수익_현금_Volume;
            _WorkSheet.get_Range("N12", Type.Missing).Value2 = futureTotal.전체_수익_직영매장판매수익;
            _WorkSheet.get_Range("N13", Type.Missing).Value2 = futureTotal.전체_수익_소계;
            _WorkSheet.get_Range("N14", Type.Missing).Value2 = futureTotal.get전체_비용_대리점투자비용();
            _WorkSheet.get_Range("N15", Type.Missing).Value2 = futureTotal.get전체_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("N16", Type.Missing).Value2 = futureTotal.get전체_비용_임차료();
            _WorkSheet.get_Range("N17", Type.Missing).Value2 = futureTotal.get전체_비용_이자비용();
            _WorkSheet.get_Range("N18", Type.Missing).Value2 = futureTotal.get전체_비용_부가세();
            _WorkSheet.get_Range("N19", Type.Missing).Value2 = futureTotal.get전체_비용_법인세();
            _WorkSheet.get_Range("N20", Type.Missing).Value2 = futureTotal.get전체_비용_기타판매관리비();
            _WorkSheet.get_Range("N21", Type.Missing).Value2 = futureTotal.전체_비용_소계;
            _WorkSheet.get_Range("N22", Type.Missing).Value2 = futureTotal.전체손익계;

            _WorkSheet.get_Range("N28", Type.Missing).Value2 = futureTotal.get도매_수익_가입자관리수수료();
            _WorkSheet.get_Range("N29", Type.Missing).Value2 = futureTotal.get도매_수익_CS관리수수료();
            _WorkSheet.get_Range("N30", Type.Missing).Value2 = futureTotal.get도매_수익_사업자모델매입에따른추가수익();
            _WorkSheet.get_Range("N31", Type.Missing).Value2 = futureTotal.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
            _WorkSheet.get_Range("N32", Type.Missing).Value2 = futureTotal.도매_수익_소계;
            _WorkSheet.get_Range("N33", Type.Missing).Value2 = futureTotal.get도매_비용_대리점투자비용();
            _WorkSheet.get_Range("N34", Type.Missing).Value2 = futureTotal.get도매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("N35", Type.Missing).Value2 = futureTotal.get도매_비용_임차료();
            _WorkSheet.get_Range("N36", Type.Missing).Value2 = futureTotal.get도매_비용_이자비용();
            _WorkSheet.get_Range("N37", Type.Missing).Value2 = futureTotal.get도매_비용_부가세();
            _WorkSheet.get_Range("N38", Type.Missing).Value2 = futureTotal.get도매_비용_법인세();
            _WorkSheet.get_Range("N39", Type.Missing).Value2 = futureTotal.get도매_비용_기타판매관리비();
            _WorkSheet.get_Range("N40", Type.Missing).Value2 = futureTotal.도매_비용_소계;
            _WorkSheet.get_Range("N41", Type.Missing).Value2 = futureTotal.도매손익계;

            _WorkSheet.get_Range("N46", Type.Missing).Value2 = futureTotal.get소매_수익_업무취급수수료();
            _WorkSheet.get_Range("N47", Type.Missing).Value2 = futureTotal.get소매_수익_직영매장판매수익();
            _WorkSheet.get_Range("N48", Type.Missing).Value2 = futureTotal.소매_수익_소계;
            _WorkSheet.get_Range("N49", Type.Missing).Value2 = futureTotal.get소매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("N50", Type.Missing).Value2 = futureTotal.get소매_비용_임차료();
            _WorkSheet.get_Range("N51", Type.Missing).Value2 = futureTotal.get소매_비용_이자비용();
            _WorkSheet.get_Range("N52", Type.Missing).Value2 = futureTotal.get소매_비용_부가세();
            _WorkSheet.get_Range("N53", Type.Missing).Value2 = futureTotal.get소매_비용_법인세();
            _WorkSheet.get_Range("N54", Type.Missing).Value2 = futureTotal.get소매_비용_기타판매관리비();
            _WorkSheet.get_Range("N55", Type.Missing).Value2 = futureTotal.소매_비용_소계;
            _WorkSheet.get_Range("N56", Type.Missing).Value2 = futureTotal.소매손익계;
            _WorkSheet.get_Range("N57", Type.Missing).Value2 = futureTotal.점별손익추정;



            _WorkSheet.get_Range("O7", Type.Missing).Value2 = future.전체_수익_가입자관리수수료;
            _WorkSheet.get_Range("O8", Type.Missing).Value2 = future.전체_수익_CS관리수수료;
            _WorkSheet.get_Range("O9", Type.Missing).Value2 = future.전체_수익_업무취급수수료;
            _WorkSheet.get_Range("O10", Type.Missing).Value2 = future.전체_수익_사업자모델매입에따른추가수익;
            _WorkSheet.get_Range("O11", Type.Missing).Value2 = future.전체_수익_유통모델매입에따른추가수익_현금_Volume;
            _WorkSheet.get_Range("O12", Type.Missing).Value2 = future.전체_수익_직영매장판매수익;
            _WorkSheet.get_Range("O13", Type.Missing).Value2 = future.전체_수익_소계;
            _WorkSheet.get_Range("O14", Type.Missing).Value2 = future.get전체_비용_대리점투자비용();
            _WorkSheet.get_Range("O15", Type.Missing).Value2 = future.get전체_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("O16", Type.Missing).Value2 = future.get전체_비용_임차료();
            _WorkSheet.get_Range("O17", Type.Missing).Value2 = future.get전체_비용_이자비용();
            _WorkSheet.get_Range("O18", Type.Missing).Value2 = future.get전체_비용_부가세();
            _WorkSheet.get_Range("O19", Type.Missing).Value2 = future.get전체_비용_법인세();
            _WorkSheet.get_Range("O20", Type.Missing).Value2 = future.get전체_비용_기타판매관리비();
            _WorkSheet.get_Range("O21", Type.Missing).Value2 = future.전체_비용_소계;
            _WorkSheet.get_Range("O22", Type.Missing).Value2 = future.전체손익계;

            _WorkSheet.get_Range("O28", Type.Missing).Value2 = future.get도매_수익_가입자관리수수료();
            _WorkSheet.get_Range("O29", Type.Missing).Value2 = future.get도매_수익_CS관리수수료();
            _WorkSheet.get_Range("O30", Type.Missing).Value2 = future.get도매_수익_사업자모델매입에따른추가수익();
            _WorkSheet.get_Range("O31", Type.Missing).Value2 = future.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
            _WorkSheet.get_Range("O32", Type.Missing).Value2 = future.도매_수익_소계;
            _WorkSheet.get_Range("O33", Type.Missing).Value2 = future.get도매_비용_대리점투자비용();
            _WorkSheet.get_Range("O34", Type.Missing).Value2 = future.get도매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("O35", Type.Missing).Value2 = future.get도매_비용_임차료();
            _WorkSheet.get_Range("O36", Type.Missing).Value2 = future.get도매_비용_이자비용();
            _WorkSheet.get_Range("O37", Type.Missing).Value2 = future.get도매_비용_부가세();
            _WorkSheet.get_Range("O38", Type.Missing).Value2 = future.get도매_비용_법인세();
            _WorkSheet.get_Range("O39", Type.Missing).Value2 = future.get도매_비용_기타판매관리비();
            _WorkSheet.get_Range("O40", Type.Missing).Value2 = future.도매_비용_소계;
            _WorkSheet.get_Range("O41", Type.Missing).Value2 = future.도매손익계;

            _WorkSheet.get_Range("O46", Type.Missing).Value2 = future.get소매_수익_업무취급수수료();
            _WorkSheet.get_Range("O47", Type.Missing).Value2 = future.get소매_수익_직영매장판매수익();
            _WorkSheet.get_Range("O48", Type.Missing).Value2 = future.소매_수익_소계;
            _WorkSheet.get_Range("O49", Type.Missing).Value2 = future.get소매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("O50", Type.Missing).Value2 = future.get소매_비용_임차료();
            _WorkSheet.get_Range("O51", Type.Missing).Value2 = future.get소매_비용_이자비용();
            _WorkSheet.get_Range("O52", Type.Missing).Value2 = future.get소매_비용_부가세();
            _WorkSheet.get_Range("O53", Type.Missing).Value2 = future.get소매_비용_법인세();
            _WorkSheet.get_Range("O54", Type.Missing).Value2 = future.get소매_비용_기타판매관리비();
            _WorkSheet.get_Range("O55", Type.Missing).Value2 = future.소매_비용_소계;
            _WorkSheet.get_Range("O56", Type.Missing).Value2 = future.소매손익계;
            _WorkSheet.get_Range("O57", Type.Missing).Value2 = future.점별손익추정;


        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="_WorkSheet"></param>
        public static void WriteDataToExcelFileDasicInput(excel.Worksheet _WorkSheet, CBasicInput g_BasicInput)
        {
            _WorkSheet.get_Range("C63", Type.Missing).Value2 = g_BasicInput.get지역();
            _WorkSheet.get_Range("E63", Type.Missing).Value2 = g_BasicInput.get대리점();
            _WorkSheet.get_Range("G63", Type.Missing).Value2 = g_BasicInput.get마케터();

            //도매
            _WorkSheet.get_Range("F7", Type.Missing).Value2 = g_BasicInput.get도매_누적가입자수();

            _WorkSheet.get_Range("F8", Type.Missing).Value2 = g_BasicInput.get도매_월평균판매대수_신규();
            _WorkSheet.get_Range("F9", Type.Missing).Value2 = g_BasicInput.get도매_월평균판매대수_기변();
            _WorkSheet.get_Range("F10", Type.Missing).Value2 = g_BasicInput.get도매_월평균판매대수_소계();
                                                               
            _WorkSheet.get_Range("F11", Type.Missing).Value2 = g_BasicInput.get도매_월평균유통모델출고대수_LG();
            _WorkSheet.get_Range("F12", Type.Missing).Value2 = g_BasicInput.get도매_월평균유통모델출고대수_SS();
            _WorkSheet.get_Range("F13", Type.Missing).Value2 = g_BasicInput.get도매_월평균유통모델출고대수_소계();
                                                               
            _WorkSheet.get_Range("F14", Type.Missing).Value2 = g_BasicInput.get도매_거래선수_개통사무실();
            _WorkSheet.get_Range("F16", Type.Missing).Value2 = g_BasicInput.get도매_거래선수_판매점();
            _WorkSheet.get_Range("F17", Type.Missing).Value2 = g_BasicInput.get도매_거래선수_소계();
                                                               
            _WorkSheet.get_Range("F18", Type.Missing).Value2 = g_BasicInput.get도매_직원수_간부급();
            _WorkSheet.get_Range("F19", Type.Missing).Value2 = g_BasicInput.get도매_직원수_평사원();
            _WorkSheet.get_Range("F20", Type.Missing).Value2 = g_BasicInput.get도매_직원수_소계();

            //소매
            _WorkSheet.get_Range("G8", Type.Missing).Value2 = g_BasicInput.get소매_월평균판매대수_신규();
            _WorkSheet.get_Range("G9", Type.Missing).Value2 = g_BasicInput.get소매_월평균판매대수_기변();
            _WorkSheet.get_Range("G10", Type.Missing).Value2 = g_BasicInput.get소매_월평균판매대수_소계();
                                                               
            _WorkSheet.get_Range("G15", Type.Missing).Value2 = g_BasicInput.get소매_거래선수_직영점();
            _WorkSheet.get_Range("G17", Type.Missing).Value2 = g_BasicInput.get소매_거래선수_소계();
                                                               
            _WorkSheet.get_Range("G18", Type.Missing).Value2 = g_BasicInput.get소매_직원수_간부급();
            _WorkSheet.get_Range("G19", Type.Missing).Value2 = g_BasicInput.get소매_직원수_평사원();
            _WorkSheet.get_Range("G20", Type.Missing).Value2 = g_BasicInput.get소매_직원수_소계();

            //합계
            _WorkSheet.get_Range("H7", Type.Missing).Value2 = g_BasicInput.get누적가입자수_합계();

            _WorkSheet.get_Range("H8", Type.Missing).Value2 = g_BasicInput.get월평균판매대수_신규_합계();
            _WorkSheet.get_Range("H9", Type.Missing).Value2 = g_BasicInput.get월평균판매대수_기변_합계();
            _WorkSheet.get_Range("H10", Type.Missing).Value2 = g_BasicInput.get월평균판매대수_소계_합계();
                                                               
            _WorkSheet.get_Range("H11", Type.Missing).Value2 = g_BasicInput.get월평균유통모델출고대수_LG_합계();
            _WorkSheet.get_Range("H12", Type.Missing).Value2 = g_BasicInput.get월평균유통모델출고대수_SS_합계();
            _WorkSheet.get_Range("H13", Type.Missing).Value2 = g_BasicInput.get월평균유통모델출고대수_소계_합계();
                                                               
            _WorkSheet.get_Range("H14", Type.Missing).Value2 = g_BasicInput.get거래선수_개통사무실_합계();
            _WorkSheet.get_Range("H15", Type.Missing).Value2 = g_BasicInput.get거래선수_직영점_합계();
            _WorkSheet.get_Range("H16", Type.Missing).Value2 = g_BasicInput.get거래선수_판매점_합계();
            _WorkSheet.get_Range("H17", Type.Missing).Value2 = g_BasicInput.get거래선수_소계_합계();
                                                               
            _WorkSheet.get_Range("H18", Type.Missing).Value2 = g_BasicInput.get직원수_간부급_합계();
            _WorkSheet.get_Range("H19", Type.Missing).Value2 = g_BasicInput.get직원수_평사원_합계();
            _WorkSheet.get_Range("H20", Type.Missing).Value2 = g_BasicInput.get직원수_소계_합계();


        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="_WorkSheet"></param>
        /// <param name="g_BasicInput"></param>
        public static void WriteDataToExcelFileDetailInput(excel.Worksheet _WorkSheet, CBasicInput g_BasicInput, CBusinessData g_DetailInput)
        {
            //도매
            _WorkSheet.get_Range("G26", Type.Missing).Value2 = g_DetailInput.get도매_수익_월평균관리수수료();
            _WorkSheet.get_Range("G27", Type.Missing).Value2 = g_DetailInput.get도매_수익_CS관리수수료();//월총액
            _WorkSheet.get_Range("G28", Type.Missing).Value2 = g_DetailInput.get도매_수익_CS관리수수료_분기();//분기총액
            _WorkSheet.get_Range("G29", Type.Missing).Value2 = g_DetailInput.get도매_수익_사업자모델매입관련추가수익();
            _WorkSheet.get_Range("G30", Type.Missing).Value2 = g_DetailInput.get도매_수익_유통모델매입관련추가수익_현금DC();
            _WorkSheet.get_Range("G31", Type.Missing).Value2 = g_DetailInput.get도매_수익_유통모델매입관련추가수익_VolumeDC();
            _WorkSheet.get_Range("G32", Type.Missing).Value2 = g_DetailInput.get도매_비용_대리점투자금액_신규();
            _WorkSheet.get_Range("G33", Type.Missing).Value2 = g_DetailInput.get도매_비용_대리점투자금액_기변();

            _WorkSheet.get_Range("G34", Type.Missing).Value2 = g_DetailInput.get도매_비용_직원급여_간부급_총액(g_BasicInput.get도매_직원수_간부급());//총액
            _WorkSheet.get_Range("G35", Type.Missing).Value2 = g_DetailInput.get도매_비용_직원급여_평사원_총액(g_BasicInput.get도매_직원수_평사원());//총액
            _WorkSheet.get_Range("G36", Type.Missing).Value2 = g_DetailInput.get도매_비용_직원급여_간부급();//월평균
            _WorkSheet.get_Range("G37", Type.Missing).Value2 = g_DetailInput.get도매_비용_직원급여_평사원();//월평균

            _WorkSheet.get_Range("G38", Type.Missing).Value2 = g_DetailInput.get도매_비용_지급임차료();
            _WorkSheet.get_Range("G39", Type.Missing).Value2 = g_DetailInput.get도매_비용_운반비();
            _WorkSheet.get_Range("G40", Type.Missing).Value2 = g_DetailInput.get도매_비용_차량유지비();
            _WorkSheet.get_Range("G41", Type.Missing).Value2 = g_DetailInput.get도매_비용_지급수수료();
            _WorkSheet.get_Range("G42", Type.Missing).Value2 = g_DetailInput.get도매_비용_판매촉진비();
            _WorkSheet.get_Range("G43", Type.Missing).Value2 = g_DetailInput.get도매_비용_건물관리비();

            _WorkSheet.get_Range("G44", Type.Missing).Value2 = g_DetailInput.get소매_수익_월평균업무취급수수료();
            _WorkSheet.get_Range("G45", Type.Missing).Value2 = g_DetailInput.get소매_수익_직영매장판매수익();
            _WorkSheet.get_Range("G46", Type.Missing).Value2 = g_DetailInput.get소매_비용_직원급여_간부급_총액(g_BasicInput.get소매_직원수_간부급());//총액
            _WorkSheet.get_Range("G47", Type.Missing).Value2 = g_DetailInput.get소매_비용_직원급여_평사원_총액(g_BasicInput.get소매_직원수_평사원());//총액
            _WorkSheet.get_Range("G48", Type.Missing).Value2 = g_DetailInput.get소매_비용_직원급여_간부급();//월평균
            _WorkSheet.get_Range("G49", Type.Missing).Value2 = g_DetailInput.get소매_비용_직원급여_평사원();//월평균


            _WorkSheet.get_Range("G50", Type.Missing).Value2 = g_DetailInput.get소매_비용_지급임차료();
            _WorkSheet.get_Range("G51", Type.Missing).Value2 = g_DetailInput.get소매_비용_지급수수료();
            _WorkSheet.get_Range("G52", Type.Missing).Value2 = g_DetailInput.get소매_비용_판매촉진비();
            _WorkSheet.get_Range("G53", Type.Missing).Value2 = g_DetailInput.get소매_비용_건물관리비();

            _WorkSheet.get_Range("G54", Type.Missing).Value2 = g_DetailInput.get도소매_비용_복리후생비();
            _WorkSheet.get_Range("G55", Type.Missing).Value2 = g_DetailInput.get도소매_비용_통신비();
            _WorkSheet.get_Range("G56", Type.Missing).Value2 = g_DetailInput.get도소매_비용_공과금();
            _WorkSheet.get_Range("G57", Type.Missing).Value2 = g_DetailInput.get도소매_비용_소모품비();
            _WorkSheet.get_Range("G58", Type.Missing).Value2 = g_DetailInput.get도소매_비용_이자비용();
            _WorkSheet.get_Range("G59", Type.Missing).Value2 = g_DetailInput.get도소매_비용_부가세();
            _WorkSheet.get_Range("G60", Type.Missing).Value2 = g_DetailInput.get도소매_비용_법인세();
            _WorkSheet.get_Range("G61", Type.Missing).Value2 = g_DetailInput.get도소매_비용_기타();

        }

        private static void WriteExcelFileToDataReport(excel.Worksheet workSheet3)
        {
            workSheet3.get_Range("B4", Type.Missing).Value2 = CDataControl.g_ReportData.get분석내용_및_대리점_활동방향();
            workSheet3.get_Range("C4", Type.Missing).Value2 = CDataControl.g_ReportData.getLG_지원_활동();
            workSheet3.get_Range("D4", Type.Missing).Value2 = CDataControl.g_ReportData.get배경_및_이슈();
        }

        public static void writeLGEFile(String filepath, String spliter, int type = 파일종류_기본)
        {
            try
            {
                string lge = null;
                switch (type) {
                    case 파일종류_기본:
                        lge = CDataControl.getSplitedLGEFileFromData(spliter);
                        break;
                    case 파일종류_관리자:
                        lge = CDataControl.getAdminDataBySerialization(spliter);
                        break;
                    case 파일종류_시뮬레이션:
                        lge = CDataControl.getSplitedLGEFileFromSimulData(spliter);
                        break;
                    default :
                        break;
                }
                FileIOPermission permission = new FileIOPermission(FileIOPermissionAccess.AllAccess, CommonUtil.defaultManagerFileName);
                permission.Demand();
                System.IO.File.WriteAllText(filepath, CommonUtil.Base64Encode(lge));
            }
            catch (Exception ex)
            {
                MessageBox.Show("파일을 저장할 수 없습니다.\n\nReported error: " + ex.Message);
            }
        }

        public static void readLGEFile(String filepath, String spliter, int type = 파일종류_기본)
        {
            try
            {
                string lge = null;
                lge = CommonUtil.Base64Decode(System.IO.File.ReadAllText(filepath));
                switch (type)
                {
                    case 파일종류_기본:
                        CDataControl.setDataFromLGEFile(lge, spliter);
                        break;
                    case 파일종류_관리자:
                        CDataControl.setAdminDataFromLGEFile(lge, spliter);
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("파일을 열 수 없습니다.\n\nReported error: " + ex.Message);
            }
        }

        public static void deepCopyBasicInput(CBasicInput srcBasicInput, CBasicInput dstBasicInput)
        {
            dstBasicInput.set지역(srcBasicInput.get지역());
            dstBasicInput.set대리점(srcBasicInput.get대리점());
            dstBasicInput.set마케터(srcBasicInput.get마케터());
            dstBasicInput.setArrData(srcBasicInput.getArrData());
        }

        public static void deepCopyBusinessData(CBusinessData srcBusinessData, CBusinessData dstBusinessData)
        {
            dstBusinessData.setArrData(srcBusinessData.getArrData());
        }

        //public static void ReadExcelFileToDataResult()
        //{

        //}

        public static string Base64Encode(string src)
        {
            byte[] arr = System.Text.Encoding.UTF8.GetBytes(src);
            return Convert.ToBase64String(arr);
        }

        public static string Base64Decode(string src)
        {
            byte[] arr = Convert.FromBase64String(src);
            return System.Text.Encoding.UTF8.GetString(arr);
        }

        public static string NullToEmpty(object str)
        {
            string returnValue = "";

            if (str != null)
            {
                if (str is string)
                {
                    returnValue = (str as string);
                }
                else 
                    returnValue = str.ToString();
            }

            return returnValue;
        }

        internal static void setInputData(string[] txtWrite, string[] txtWrite2, CBasicInput bi, CBusinessData di, CResultData[] rdts, CResultData[] rds, CResultData rdt, CResultData rd, CResultData businessTotal, CResultData business)
        {
            bi.setArrData_BasicInput(txtWrite);
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
                rdt.set도매_비용_인건비_급여_복리후생비(
                    di.get도매_비용_직원급여_간부급() * bi.get도매_직원수_간부급() +
                    di.get도매_비용_직원급여_평사원() * bi.get도매_직원수_평사원() +
                    Convert.ToDouble(
                        CommonUtil.Division(
                            Convert.ToDouble(
                                di.get도소매_비용_복리후생비()
                            ),
                            Convert.ToDouble(
                                bi.get월평균판매대수_소계_합계()
                            )
                        ) * bi.get도매_월평균판매대수_소계()
                    )
                );
                rdt.set도매_비용_임차료(di.get도매_비용_지급임차료());
                rdt.set도매_비용_이자비용(Convert.ToDouble(CommonUtil.Division(Convert.ToDouble(di.get도소매_비용_이자비용()), Convert.ToDouble(bi.get월평균판매대수_소계_합계())) * Convert.ToDouble(bi.get도매_월평균판매대수_소계())));
                rdt.set도매_비용_부가세(Convert.ToDouble(CommonUtil.Division(Convert.ToDouble(di.get도소매_비용_부가세()), Convert.ToDouble(bi.get월평균판매대수_소계_합계())) * Convert.ToDouble(bi.get도매_월평균판매대수_소계())));
                rdt.set도매_비용_법인세(Convert.ToDouble(CommonUtil.Division(Convert.ToDouble(di.get도소매_비용_법인세()), Convert.ToDouble(bi.get월평균판매대수_소계_합계())) * Convert.ToDouble(bi.get도매_월평균판매대수_소계())));
                /* 기타판매관리비
                 *  SUM(
                 *      'Input(기본+세부항목)'!F35,
                 *      'Input(기본+세부항목)'!F36,
                 *      'Input(기본+세부항목)'!F37,
                 *      'Input(기본+세부항목)'!F38,
                 *      'Input(기본+세부항목)'!F39
                 *  )
                 *  +
                 *  (
                 *      SUM(
                 *          'Input(기본+세부항목)'!F48, // 복리후생비
                 *          'Input(기본+세부항목)'!F49,
                 *          'Input(기본+세부항목)'!F50,
                 *          'Input(기본+세부항목)'!F51,
                 *          'Input(기본+세부항목)'!F53
                 *      )
                 *      / 'Input(기본+세부항목)'!G10 * 'Input(기본+세부항목)'!E10
                 *  )
                 */
                rdt.set도매_비용_기타판매관리비(
                    di.get도매_비용_운반비() + di.get도매_비용_차량유지비() + di.get도매_비용_지급수수료() + di.get도매_비용_판매촉진비() + di.get도매_비용_건물관리비() +
                    Convert.ToDouble(
                        CommonUtil.Division(
                            Convert.ToDouble(
                                di.get도소매_비용_통신비() + di.get도소매_비용_공과금() + di.get도소매_비용_소모품비() + di.get도소매_비용_기타()
                            ),
                            Convert.ToDouble(
                                bi.get월평균판매대수_소계_합계()
                            )
                        ) * bi.get도매_월평균판매대수_소계()
                    )
                );
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
                rdt.set소매_비용_인건비_급여_복리후생비(
                    di.get소매_비용_직원급여_간부급() * bi.get소매_직원수_간부급() +
                    di.get소매_비용_직원급여_평사원() * bi.get소매_직원수_평사원() +
                    Convert.ToDouble(
                        CommonUtil.Division(
                            Convert.ToDouble(
                                di.get도소매_비용_복리후생비()
                            ),
                            Convert.ToDouble(
                                bi.get월평균판매대수_소계_합계()
                            )
                        ) * bi.get소매_월평균판매대수_소계()
                    )
                );
                rdt.set소매_비용_임차료(di.get소매_비용_지급임차료());
                rdt.set소매_비용_이자비용(CommonUtil.Division(di.get도소매_비용_이자비용(), bi.get월평균판매대수_소계_합계()) * bi.get소매_월평균판매대수_소계());
                rdt.set소매_비용_부가세(CommonUtil.Division(di.get도소매_비용_부가세(), bi.get월평균판매대수_소계_합계()) * bi.get소매_월평균판매대수_소계());
                rdt.set소매_비용_법인세(CommonUtil.Division(di.get도소매_비용_법인세(), bi.get월평균판매대수_소계_합계()) * bi.get소매_월평균판매대수_소계());
                rdt.set소매_비용_기타판매관리비(
                    (
                        di.get소매_비용_지급수수료() + di.get소매_비용_판매촉진비() + di.get소매_비용_건물관리비()
                    ) +
                    Convert.ToDouble(
                        CommonUtil.Division(
                            Convert.ToDouble(
                                di.get도소매_비용_통신비() + di.get도소매_비용_공과금() + di.get도소매_비용_소모품비() + di.get도소매_비용_기타()
                            ),
                            Convert.ToDouble(bi.get월평균판매대수_소계_합계())
                        ) *
                        bi.get소매_월평균판매대수_소계()
                    )
                );
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
                rd.점별손익추정 = bi.get거래선수_직영점_합계();
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
                //rdt.set전체_비용_기타판매관리비(di.get도매_비용_운반비() + di.get도매_비용_차량유지비() + di.get도매_비용_지급수수료() + di.get도매_비용_판매촉진비() + di.get도매_비용_건물관리비() + di.get소매_비용_지급수수료() + di.get소매_비용_판매촉진비() + di.get소매_비용_건물관리비() + di.get도소매_비용_통신비() + di.get도소매_비용_공과금() + di.get도소매_비용_소모품비() + di.get도소매_비용_기타());
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

            //  업계 평균적용 결과
            Boolean isOver2000 = bi.get월평균판매대수_소계_합계() >= 2000;
            rdt = businessTotal;
            rd = business;
            di = CDataControl.g_BusinessAvg;     // 관리자가 배포한 업계 단위비용
            //      도매
            //          총액
            //              수익
            rdt.set도매_수익_가입자관리수수료(di.get도매_수익_월평균관리수수료() * bi.get도매_누적가입자수());
            rdt.set도매_수익_CS관리수수료(di.get도매_수익_CS관리수수료() * bi.get도매_누적가입자수());
            Double 사업자모델매입에따른추가수익_단위금액 = Convert.ToDouble(Convert.ToDouble(di.ASP_사업자_소계) * (isOver2000 ? 0.01 : 0.005));
            di.set도매_수익_사업자모델매입관련추가수익(사업자모델매입에따른추가수익_단위금액);       // 프린트용 정보 저장
            rdt.set도매_수익_사업자모델매입에따른추가수익(
                사업자모델매입에따른추가수익_단위금액 * // 판매량이 2000대 이상일때는 asp의 1%, 미만일때는 asp의 0.5%
                (bi.get월평균판매대수_소계_합계() - bi.get월평균유통모델출고대수_소계_합계())
            );
            /* 유통모델관련 추가수익
             * 현금DC
             *  월평균유통모델출고대수_LG_소계*유통모델_LG*0.8%+
             *  월평균유통모델출고대수_SS_소계*유통모델_SS*0.6%
             * 
             * Volume DC
             *  IF('Input(기본+세부항목)'!G10>2000,
             *      월평균유통모델출고대수_LG_소계*3%*유통모델_LG+
             *      월평균유통모델출고대수_SS_소계*유통모델_SS*2.2%,
             *      
             *      월평균유통모델출고대수_LG_소계*1.5%*유통모델_LG+
             *      월평균유통모델출고대수_SS_소계*유통모델_SS*1%)
             */
            Double 유통모델매입에따른추가수익_단위금액 =
                    Convert.ToDouble(CommonUtil.Division((bi.get월평균유통모델출고대수_SS_합계() * Convert.ToDouble(di.ASP_유통_SS) * 0.006 + bi.get월평균유통모델출고대수_LG_합계() * Convert.ToDouble(di.ASP_유통_LG) * 0.008), bi.get월평균유통모델출고대수_소계_합계())) +    // 현금DC
                    Convert.ToDouble(CommonUtil.Division((bi.get월평균유통모델출고대수_SS_합계() * Convert.ToDouble(di.ASP_유통_SS) * (isOver2000 ? 0.022 : 0.01) + bi.get월평균유통모델출고대수_LG_합계() * Convert.ToDouble(di.ASP_유통_LG) * (isOver2000 ? 0.03 : 0.015)), bi.get월평균유통모델출고대수_소계_합계()));    // Volume DC
            di.set도매_수익_유통모델매입관련추가수익_현금DC(Convert.ToDouble(CommonUtil.Division((bi.get월평균유통모델출고대수_SS_합계() * Convert.ToDouble(di.ASP_유통_SS) * 0.006 + bi.get월평균유통모델출고대수_LG_합계() * Convert.ToDouble(di.ASP_유통_LG) * 0.008), bi.get월평균유통모델출고대수_소계_합계())));
            di.set도매_수익_유통모델매입관련추가수익_VolumeDC(Convert.ToDouble(CommonUtil.Division((bi.get월평균유통모델출고대수_SS_합계() * Convert.ToDouble(di.ASP_유통_SS) * (isOver2000 ? 0.022 : 0.01) + bi.get월평균유통모델출고대수_LG_합계() * Convert.ToDouble(di.ASP_유통_LG) * (isOver2000 ? 0.03 : 0.015)), bi.get월평균유통모델출고대수_소계_합계())));
            rdt.set도매_수익_유통모델매입에따른추가수익_현금_Volume(
               유통모델매입에따른추가수익_단위금액 * bi.get월평균유통모델출고대수_소계_합계()
            );
           
            rdt.도매_수익_소계 = rdt.get도매_수익_가입자관리수수료() + rdt.get도매_수익_CS관리수수료() + rdt.get도매_수익_사업자모델매입에따른추가수익() + rdt.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
            //              비용
            rdt.set도매_비용_대리점투자비용(di.get도매_비용_대리점투자금액_신규() * bi.get도매_월평균판매대수_신규() + di.get도매_비용_대리점투자금액_기변() * bi.get도매_월평균판매대수_기변());
            rdt.set도매_비용_인건비_급여_복리후생비(di.get도매_비용_직원급여_간부급() * bi.get도매_직원수_간부급() + di.get도매_비용_직원급여_평사원() * bi.get도매_직원수_평사원() + di.get도소매_비용_복리후생비() * bi.get도매_직원수_소계());
            rdt.set도매_비용_임차료(di.get도매_비용_지급임차료() * bi.get도매_거래선수_개통사무실());
            rdt.set도매_비용_이자비용(di.get도소매_비용_이자비용() * bi.get도매_월평균판매대수_소계());

            /* 부가세 및 법인세 단위금액의 수식
                    *  (
                    *      (
                    *          (
                    *              월평균관리수수료(CS관리수수료포함) * 누적가입자수+
                    *              (
                    *                  (
                    *                      ASP전체계-리베이트
                    *                  )
                    *                  *소매_월평균판매대수_계
                    *              )
                    *              +
                    *              월단위취급수수료(단위금액) * 전체_월평균판매대수_계+직영매장수익(단위금액) * 소매_월평균판매대수_계+리베이트*도매_월평균판매대수_계+
                    *              (
                    *                  (
                    *                      ASP전체계-리베이트
                    *                  )
                    *                  *전체_월평균판매대수_계
                    *              )
                    *              -
                    *              (
                    *                  전체_월평균판매대수_계*ASP전체계
                    *              )
                    *          )
                    *          *10%
                    *      )
                    *      +
                    *      (
                    *          SUM(
                    *              월평균관리수수료(CS관리수수료포함) * 누적가입자수,
                    *              사업자모델매입관련추가수익(단위금액) * (전체_월평균판매대수_계 - 전체_월평균유통모델출고대수_계),
                    *              유통모델매입관련추가수익(현금DC)(단위금액) * 전체_월평균유통모델출고대수_계,
                    *              유통모델매입관련추가수익(VolumeDC)(단위금액) * 전체_월평균유통모델출고대수_계,
                    *              월단위취급수수료(단위금액) * 전체_월평균판매대수_계,
                    *              직영매장수익(단위금액) * 소매_월평균판매대수_계
                    *          )
                    *          -
                    *          SUM(
                    *              대리점투자금액(신규) *도매_월평균판매대수_신규 +대리점투자금액(기변) *도매_월평균판매대수_기변,
                    *              직원급여(간부급) *도매_직원수_간부급 +직원급여(평사원) *도매_직원수_평사원,
                    *              복리후생비 *도매_직원수_소계,
                    *              통신비 *도매_직원수_소계,
                    *              세금과공과금 *도매_직원수_소계,
                    *              지급임차료 *도매_거래선수_개통사무실,
                    *              운반비 *도매_월평균판매대수_소계,
                    *              소모품비 *도매_월평균판매대수_소계,
                    *              지급수수료 *도매_월평균판매대수_소계,
                    *              판매촉진비 *도매_월평균판매대수_소계,
                    *              건물관리비 *도매_거래선수_개통사무실,
                    *              이자비용 *도매_월평균판매대수_소계,
                    *              차량유지비 *도매_직원수_소계,
                    *              기타 *도매_월평균판매대수_소계,
                    *              소매_직원급여(간부급) *소매_직원수_간부급 +소매_직원급여(평사원)*소매_직원수_평사원,
                    *              복리후생비 *소매_직원수_소계,
                    *              통신비 *소매_직원수_소계,
                    *              세금과공과금 *소매_직원수_소계,
                    *              지급임차료 *소매_거래선수_소계,
                    *              소모품비 *소매_월평균판매대수_소계,
                    *              지급수수료 *소매_월평균판매대수_소계,
                    *              판매촉진비 *소매_월평균판매대수_소계,
                    *              건물관리비 *소매_거래선수_소계,
                    *              이자비용 *소매_월평균판매대수_소계,
                    *              기타 *소매_월평균판매대수_소계,
                    *          )
                    *          -
                    *          (
                    *              (
                    *                  월평균관리수수료(CS관리수수료포함) * 누적가입자수+
                    *                  (
                    *                      (
                    *                          ASP전체계-리베이트
                    *                      )
                    *                      *
                    *                      소매_월평균판매대수_계
                    *                  )
                    *                  +
                    *                  월단위취급수수료(단위금액) * 전체_월평균판매대수_계+직영매장수익(단위금액) * 소매_월평균판매대수_계+리베이트*도매_월평균판매대수_계
                    *                  +
                    *                  (
                    *                      (
                    *                          ASP전체계-리베이트
                    *                      )
                    *                      *
                    *                      전체_월평균판매대수_계
                    *                  )
                    *                  -
                    *                  (
                    *                      전체_월평균판매대수_계*ASP전체계
                    *                  )
                    *              )
                    *              *
                    *              10%
                    *          )
                    *      )
                    *      *
                    *      22%
                    *  )
                    *  /
                    *  전체_월평균판매대수_계
                    */

            // 부가세
            /*  
             *      (
             *          (
             *              월평균관리수수료(CS관리수수료포함) * 누적가입자수+
             *              (
             *                  (
             *                      ASP전체계-리베이트
             *                  )
             *                  *소매_월평균판매대수_계
             *              )
             *              +
             *              월단위취급수수료(단위금액) * 전체_월평균판매대수_계+직영매장수익(단위금액) * 소매_월평균판매대수_계+리베이트*도매_월평균판매대수_계+
             *              (
             *                  (
             *                      ASP전체계-리베이트
             *                  )
             *                  *전체_월평균판매대수_계
             *              )
             *              -
             *              (
             *                  전체_월평균판매대수_계*ASP전체계
             *              )
             *          )
             *          *10%
             *      )
             *      /
             *      전체_월평균판매대수_계
             */

            //nIAOut[k++] += CommonUtil.Division(di.get도소매_비용_부가세() , bi.get월평균판매대수_소계_합계());
            Double Doubleasp전체계 = Convert.ToDouble(di.ASP_총계);
            Double Double리베이트 = Convert.ToDouble(di.Rebate);
            Double 부가세_단위금액 = Convert.ToDouble(CommonUtil.Division(
                    (
                        (di.get도매_수익_월평균관리수수료() + di.get도매_수익_CS관리수수료()) * bi.get누적가입자수_합계() +
                        (
                            (Doubleasp전체계 - Double리베이트) * bi.get소매_월평균판매대수_소계()
                        ) + di.get소매_수익_월평균업무취급수수료() * bi.get월평균판매대수_소계_합계() + di.get소매_수익_직영매장판매수익() * bi.get소매_월평균판매대수_소계() + Double리베이트 * bi.get도매_월평균판매대수_소계() +
                        (
                            (Doubleasp전체계 - Double리베이트) * bi.get월평균판매대수_소계_합계()
                        ) -
                        (
                            bi.get월평균판매대수_소계_합계() * Doubleasp전체계
                        )
                    ) * 0.1 , bi.get월평균판매대수_소계_합계())
                );
            di.set도소매_비용_부가세(부가세_단위금액);
            rdt.set도매_비용_부가세(부가세_단위금액 * bi.get도매_월평균판매대수_소계());

            // 법인세
            /*  (
            *      (
            *          SUM(
            *              월평균관리수수료(CS관리수수료포함) * 누적가입자수,
            *              사업자모델매입관련추가수익(단위금액) * (전체_월평균판매대수_계 - 전체_월평균유통모델출고대수_계),
            *              유통모델매입관련추가수익(현금DC)(단위금액) * 전체_월평균유통모델출고대수_계,
            *              유통모델매입관련추가수익(VolumeDC)(단위금액) * 전체_월평균유통모델출고대수_계,
            *              월단위취급수수료(단위금액) * 전체_월평균판매대수_계,
            *              직영매장수익(단위금액) * 소매_월평균판매대수_계
            *          )
            *          -
            *          SUM(
            *              대리점투자금액(신규) *도매_월평균판매대수_신규 +대리점투자금액(기변) *도매_월평균판매대수_기변,
            *              직원급여(간부급) *도매_직원수_간부급 +직원급여(평사원) *도매_직원수_평사원,
            *              복리후생비 *도매_직원수_소계,
            *              통신비 *도매_직원수_소계,
            *              세금과공과금 *도매_직원수_소계,
            *              지급임차료 *도매_거래선수_개통사무실,
            *              운반비 *도매_월평균판매대수_소계,
            *              소모품비 *도매_월평균판매대수_소계,
            *              지급수수료 *도매_월평균판매대수_소계,
            *              판매촉진비 *도매_월평균판매대수_소계,
            *              건물관리비 *도매_거래선수_개통사무실,
            *              이자비용 *도매_월평균판매대수_소계,
            *              차량유지비 *도매_직원수_소계,
            *              기타 *도매_월평균판매대수_소계,
            *              소매_직원급여(간부급) *소매_직원수_간부급 +소매_직원급여(평사원)*소매_직원수_평사원,
            *              복리후생비 *소매_직원수_소계,
            *              통신비 *소매_직원수_소계,
            *              세금과공과금 *소매_직원수_소계,
            *              지급임차료 *소매_거래선수_소계,
            *              소모품비 *소매_월평균판매대수_소계,
            *              지급수수료 *소매_월평균판매대수_소계,
            *              판매촉진비 *소매_월평균판매대수_소계,
            *              건물관리비 *소매_거래선수_소계,
            *              이자비용 *소매_월평균판매대수_소계,
            *              기타 *소매_월평균판매대수_소계,
            *          )
            *          -
            *          (
            *              (
            *                  월평균관리수수료(CS관리수수료포함) * 누적가입자수+
            *                  (
            *                      (
            *                          ASP전체계-리베이트
            *                      )
            *                      *
            *                      소매_월평균판매대수_계
            *                  )
            *                  +
            *                  월단위취급수수료(단위금액) * 전체_월평균판매대수_계+직영매장수익(단위금액) * 소매_월평균판매대수_계+리베이트*도매_월평균판매대수_계
            *                  +
            *                  (
            *                      (
            *                          ASP전체계-리베이트
            *                      )
            *                      *
            *                      전체_월평균판매대수_계
            *                  )
            *                  -
            *                  (
            *                      전체_월평균판매대수_계*ASP전체계
            *                  )
            *              )
            *              *
            *              10%
            *          )
            *      )
            *      *
            *      22%
            *  )
            *  /
            *  전체_월평균판매대수_계
            */

            /* 수익합계
            *          SUM(
            *              월평균관리수수료(CS관리수수료포함) * 누적가입자수,
            *              사업자모델매입관련추가수익(단위금액) * (전체_월평균판매대수_계 - 전체_월평균유통모델출고대수_계),
            *              유통모델매입관련추가수익(현금DC)(단위금액) * 전체_월평균유통모델출고대수_계,
            *              유통모델매입관련추가수익(VolumeDC)(단위금액) * 전체_월평균유통모델출고대수_계,
            *              월단위취급수수료(단위금액) * 전체_월평균판매대수_계,
            *              직영매장수익(단위금액) * 소매_월평균판매대수_계
            *          )
             */
            Double 수익합계 = 
                (di.get도매_수익_월평균관리수수료()+di.get도매_수익_CS관리수수료()) * bi.get누적가입자수_합계() +
                사업자모델매입에따른추가수익_단위금액 * (bi.get월평균판매대수_소계_합계() - bi.get월평균유통모델출고대수_소계_합계()) +
                유통모델매입에따른추가수익_단위금액 * bi.get월평균유통모델출고대수_소계_합계() +
                di.get소매_수익_월평균업무취급수수료() * bi.get월평균판매대수_소계_합계() +
                di.get소매_수익_직영매장판매수익() * bi.get소매_월평균판매대수_소계();

            /* 비용합계
            *          SUM(
            *              대리점투자금액(신규) *도매_월평균판매대수_신규 +대리점투자금액(기변) *도매_월평균판매대수_기변,
            *              직원급여(간부급) *도매_직원수_간부급 +직원급여(평사원) *도매_직원수_평사원,
            *              복리후생비 *도매_직원수_소계,
            *              복리후생비 *소매_직원수_소계,
            *              통신비 *도매_직원수_소계,
            *              통신비 *소매_직원수_소계,
            *              세금과공과금 *도매_직원수_소계,
            *              세금과공과금 *소매_직원수_소계,
            *              지급임차료 *도매_거래선수_개통사무실,
            *              운반비 *도매_월평균판매대수_소계,
            *              소모품비 *도매_월평균판매대수_소계,
            *              소모품비 *소매_월평균판매대수_소계,
            *              지급수수료 *도매_월평균판매대수_소계,
            *              판매촉진비 *도매_월평균판매대수_소계,
            *              건물관리비 *도매_거래선수_개통사무실,
            *              이자비용 *도매_월평균판매대수_소계,
            *              차량유지비 *도매_직원수_소계,
            *              기타 *도매_월평균판매대수_소계,
            *              소매_직원급여(간부급) *소매_직원수_간부급 +소매_직원급여(평사원)*소매_직원수_평사원,
            *              지급임차료 *소매_거래선수_소계,
            *              지급수수료 *소매_월평균판매대수_소계,
            *              판매촉진비 *소매_월평균판매대수_소계,
            *              건물관리비 *소매_거래선수_소계,
            *              이자비용 *소매_월평균판매대수_소계,
            *              기타 *소매_월평균판매대수_소계,
            *          )
             */
            Double 비용합계 =
                di.get도매_비용_대리점투자금액_신규() * bi.get도매_월평균판매대수_신규() + di.get도매_비용_대리점투자금액_기변() * bi.get도매_월평균판매대수_기변() +
                di.get도매_비용_직원급여_간부급_총액(bi.get도매_직원수_간부급()) + di.get도매_비용_직원급여_평사원_총액(bi.get도매_직원수_평사원()) +
                di.get도소매_비용_복리후생비()/*소매포함*/ * bi.get직원수_소계_합계() +
                di.get도소매_비용_통신비()/*소매포함*/ * bi.get직원수_소계_합계() +
                di.get도소매_비용_공과금()/*소매포함*/ * bi.get직원수_소계_합계() +
                di.get도매_비용_지급임차료() * bi.get도매_거래선수_개통사무실() +
                di.get도매_비용_운반비() * bi.get도매_월평균판매대수_소계() +
                di.get도소매_비용_소모품비() * bi.get월평균판매대수_소계_합계() +
                di.get도매_비용_지급수수료() * bi.get도매_월평균판매대수_소계() +
                di.get도매_비용_판매촉진비() * bi.get도매_월평균판매대수_소계() +
                di.get도매_비용_건물관리비() * bi.get도매_거래선수_개통사무실() +
                di.get도소매_비용_이자비용() * bi.get월평균판매대수_소계_합계() +
                di.get도매_비용_차량유지비() * bi.get도매_직원수_소계() +
                di.get도소매_비용_기타() * bi.get월평균판매대수_소계_합계() +
                di.get소매_비용_직원급여_간부급_총액(bi.get소매_직원수_간부급()) + di.get소매_비용_직원급여_평사원_총액(bi.get소매_직원수_평사원()) +
                di.get소매_비용_지급임차료() * bi.get소매_거래선수_소계() +
                di.get소매_비용_지급수수료() * bi.get소매_월평균판매대수_소계() +
                di.get소매_비용_판매촉진비() * bi.get소매_월평균판매대수_소계() +
                di.get소매_비용_건물관리비() * bi.get소매_거래선수_소계();
            Double 법인세_단위금액 = Convert.ToDouble(CommonUtil.Division((수익합계 - 비용합계 - 부가세_단위금액 * bi.get월평균판매대수_소계_합계()) * 0.22, bi.get월평균판매대수_소계_합계()));
            di.set도소매_비용_법인세(법인세_단위금액);

            rdt.set도매_비용_법인세(법인세_단위금액 * bi.get도매_월평균판매대수_소계());
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
            rdt.set소매_비용_부가세(부가세_단위금액 * bi.get소매_월평균판매대수_소계());
            rdt.set소매_비용_법인세(법인세_단위금액 * bi.get소매_월평균판매대수_소계());
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
            rd.점별손익추정 = bi.get거래선수_직영점_합계();
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
    }
}
