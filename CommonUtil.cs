using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Collections;
using System.Runtime.InteropServices;
using System.Windows.Forms;
namespace KIWI
{
    public class CommonUtil
    {

        public const Int64 QUARTER = 3;
        private static excel.ApplicationClass application = null;
        private static excel.Workbook workBook = null;
        private static excel.ApplicationClass applicationForSimul = null;
        private static excel.Workbook workBookForSimul = null;
        public static string defaultName = AppDomain.CurrentDomain.BaseDirectory + "default.xlsx";
        //public static string openAsName = null;
        public static string saveAsSimulName = null;
        public static string saveAsName = null;
        public static string defaultManagerName = AppDomain.CurrentDomain.BaseDirectory + "manager.csv";
        public static string adminName = AppDomain.CurrentDomain.BaseDirectory + "admin.csv";

        /// <summary>
        /// 
        /// </summary>
        /// <param name="text1"></param>
        /// <param name="text2"></param>
        /// <returns></returns>
        public static string Sum_Values(string text1, string text2)
        {
            Int64 sumManager = StringToIntVal(text1) + StringToIntVal(text2);

            return sumManager.ToString();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="text1"></param>
        /// <returns></returns>
        public static Int64 StringToIntVal(object text1)
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
            return result ? 0 : Convert.ToInt64(Convert.ToDouble(returnValue));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="text1"></param>
        /// <returns></returns>
        public static string NullToString0(object obj)
        {
            return obj == null ? 0.ToString() : Convert.ToInt64(obj).ToString();
        }


        /// <summary>
        /// 분모가 0이거나, 분자가 0일경우 0을 반환, 이외의 경우 나눈 몫을 반환
        /// </summary>
        /// <param name="textDenominator">분모값</param>
        /// <param name="textNumerator">분자값</param>
        /// <returns>분모가 0이거나, 분자가 0일경우 0을 반환, 이외의 경우 나눈 몫을 반환</returns>
        public static string Division(string textDenominator, string textNumerator)
        {
            return StringToIntVal(textDenominator) != 0 ? (StringToIntVal(textNumerator) == 0 ? 0.ToString() : (StringToIntVal(textDenominator) / StringToIntVal(textNumerator)).ToString()) : 0.ToString();
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
            GetExcel_WorkBook_CLOSE();
        }
        public static void WriteDataToExcelFile(string fileName, CBasicInput g_BasicInput, CBusinessData g_DetailInput)
        {
            GetExcel_WorkBook(fileName);
            excel.Worksheet workSheet1 = workBook.Sheets[1] as excel.Worksheet;
            WriteDataToExcelFileDasicInput(workSheet1, g_BasicInput);
            WriteDataToExcelFileDetailInput(workSheet1, g_BasicInput, g_DetailInput);
            excel.Worksheet workSheet2 = workBook.Sheets[2] as excel.Worksheet;
            WriteExcelFileToDataResultBusiness(workSheet2);
            WriteExcelFileToDataResultStore(workSheet2);
            WriteExcelFileToDataResultFuture(workSheet2);
            GetExcel_WorkBook_CLOSE();

        }

        public static void ReadFileManagerToData()
        {
            //관리자 파일을 읽어 넣는다
            try
            {
                string csv = System.IO.File.ReadAllText(defaultManagerName);
                csv = CommonUtil.Base64Decode(csv);
                string[] splitedCsv = csv.Split(',');
                CDataControl.g_BusinessAvg.setArrData_DetailInput(splitedCsv);
            }
            catch (Exception ex)
            {
                String[] txtWrite2 = new String[31];
                // 파일이 없음
                for (int i = 0; i < txtWrite2.Length; i++)
                {
                    txtWrite2[i] = 0.ToString();
                }
                CDataControl.g_BusinessAvg.setArrData_DetailInput(txtWrite2);
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
            CDataControl.g_FileDetailInput.set도매_수익_월평균관리수수료(NullToEmpty(_WorkSheet.get_Range("G26", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도매_수익_CS관리수수료(NullToEmpty(_WorkSheet.get_Range("G27", Type.Missing).Value2));//월총액
            CDataControl.g_FileDetailInput.set도매_수익_사업자모델매입관련추가수익(NullToEmpty(_WorkSheet.get_Range("G29", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도매_수익_유통모델매입관련추가수익_현금DC(NullToEmpty(_WorkSheet.get_Range("G30", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도매_수익_유통모델매입관련추가수익_VolumeDC(NullToEmpty(_WorkSheet.get_Range("G31", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도매_비용_대리점투자금액_신규(NullToEmpty(_WorkSheet.get_Range("G32", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도매_비용_대리점투자금액_기변(NullToEmpty(_WorkSheet.get_Range("G33", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도매_비용_직원급여_간부급(NullToEmpty(_WorkSheet.get_Range("G34", Type.Missing).Value2));//총액
            CDataControl.g_FileDetailInput.set도매_비용_직원급여_평사원(NullToEmpty(_WorkSheet.get_Range("G35", Type.Missing).Value2));//총액
            CDataControl.g_FileDetailInput.set도매_비용_지급임차료(NullToEmpty(_WorkSheet.get_Range("G38", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도매_비용_운반비(NullToEmpty(_WorkSheet.get_Range("G39", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도매_비용_차량유지비(NullToEmpty(_WorkSheet.get_Range("G40", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도매_비용_지급수수료(NullToEmpty(_WorkSheet.get_Range("G41", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도매_비용_판매촉진비(NullToEmpty(_WorkSheet.get_Range("G42", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도매_비용_건물관리비(NullToEmpty(_WorkSheet.get_Range("G43", Type.Missing).Value2));

            CDataControl.g_FileDetailInput.set소매_수익_월평균업무취급수수료(NullToEmpty(_WorkSheet.get_Range("G44", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set소매_수익_직영매장판매수익(NullToEmpty(_WorkSheet.get_Range("G45", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set소매_비용_직원급여_간부급(NullToEmpty(_WorkSheet.get_Range("G46", Type.Missing).Value2));//총액
            CDataControl.g_FileDetailInput.set소매_비용_직원급여_평사원(NullToEmpty(_WorkSheet.get_Range("G47", Type.Missing).Value2));//총액
            CDataControl.g_FileDetailInput.set소매_비용_지급임차료(NullToEmpty(_WorkSheet.get_Range("G50", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set소매_비용_지급수수료(NullToEmpty(_WorkSheet.get_Range("G51", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set소매_비용_판매촉진비(NullToEmpty(_WorkSheet.get_Range("G52", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set소매_비용_건물관리비(NullToEmpty(_WorkSheet.get_Range("G53", Type.Missing).Value2));

            CDataControl.g_FileDetailInput.set도소매_비용_복리후생비(NullToEmpty(_WorkSheet.get_Range("G54", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도소매_비용_통신비(NullToEmpty(_WorkSheet.get_Range("G55", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도소매_비용_공과금(NullToEmpty(_WorkSheet.get_Range("G56", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도소매_비용_소모품비(NullToEmpty(_WorkSheet.get_Range("G57", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도소매_비용_이자비용(NullToEmpty(_WorkSheet.get_Range("G58", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도소매_비용_부가세(NullToEmpty(_WorkSheet.get_Range("G59", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도소매_비용_법인세(NullToEmpty(_WorkSheet.get_Range("G60", Type.Missing).Value2));
            CDataControl.g_FileDetailInput.set도소매_비용_기타(NullToEmpty(_WorkSheet.get_Range("G61", Type.Missing).Value2));

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="_WorkSheet"></param>
        public static void ReadExcelFileToDataResultBusiness(excel.Worksheet _WorkSheet)
        {
            CDataControl.g_FileResultBusinessTotal.전체_수익_가입자수수료 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("D7", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.전체_수익_CS관리수수료 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("D8", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.전체_수익_업무취급수수료 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("D9", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.전체_수익_사업자모델매입에따른추가수익 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("D10", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.전체_수익_유통모델매입에따른추가수익_현금_Volume = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("D11", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.전체_수익_직영매장판매수익 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("D12", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.전체_수익_소계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("D13", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set전체_비용_대리점투자비용(NullToEmpty(_WorkSheet.get_Range("D14", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set전체_비용_인건비_급여_복리후생비(NullToEmpty(_WorkSheet.get_Range("D15", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set전체_비용_임차료(NullToEmpty(_WorkSheet.get_Range("D16", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set전체_비용_이자비용(NullToEmpty(_WorkSheet.get_Range("D17", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set전체_비용_부가세(NullToEmpty(_WorkSheet.get_Range("D18", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set전체_비용_법인세(NullToEmpty(_WorkSheet.get_Range("D19", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set전체_비용_기타판매관리비(NullToEmpty(_WorkSheet.get_Range("D20", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.전체_비용_소계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("D21", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.전체손익계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("D22", Type.Missing).Value2));

            CDataControl.g_FileResultBusinessTotal.set도매_수익_가입자관리수수료(NullToEmpty(_WorkSheet.get_Range("D28", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set도매_수익_CS관리수수료(NullToEmpty(_WorkSheet.get_Range("D29", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set도매_수익_사업자모델매입에따른추가수익(NullToEmpty(_WorkSheet.get_Range("D30", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set도매_수익_유통모델매입에따른추가수익_현금_Volume(NullToEmpty(_WorkSheet.get_Range("D31", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.도매_수익_소계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("D32", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set도매_비용_대리점투자비용(NullToEmpty(_WorkSheet.get_Range("D33", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set도매_비용_인건비_급여_복리후생비(NullToEmpty(_WorkSheet.get_Range("D34", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set도매_비용_임차료(NullToEmpty(_WorkSheet.get_Range("D35", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set도매_비용_이자비용(NullToEmpty(_WorkSheet.get_Range("D36", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set도매_비용_부가세(NullToEmpty(_WorkSheet.get_Range("D37", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set도매_비용_법인세(NullToEmpty(_WorkSheet.get_Range("D38", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set도매_비용_기타판매관리비(NullToEmpty(_WorkSheet.get_Range("D39", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.도매_비용_소계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("D40", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.도매손익계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("D41", Type.Missing).Value2));

            CDataControl.g_FileResultBusinessTotal.set소매_수익_업무취급수수료(NullToEmpty(_WorkSheet.get_Range("D46", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set소매_수익_직영매장판매수익(NullToEmpty(_WorkSheet.get_Range("D47", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.소매_수익_소계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("D48", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set소매_비용_인건비_급여_복리후생비(NullToEmpty(_WorkSheet.get_Range("D49", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set소매_비용_임차료(NullToEmpty(_WorkSheet.get_Range("D50", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set소매_비용_이자비용(NullToEmpty(_WorkSheet.get_Range("D51", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set소매_비용_부가세(NullToEmpty(_WorkSheet.get_Range("D52", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set소매_비용_법인세(NullToEmpty(_WorkSheet.get_Range("D53", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.set소매_비용_기타판매관리비(NullToEmpty(_WorkSheet.get_Range("D54", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.소매_비용_소계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("D55", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.소매손익계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("D56", Type.Missing).Value2));
            CDataControl.g_FileResultBusinessTotal.점별손익추정 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("D57", Type.Missing).Value2));



            CDataControl.g_FileResultBusiness.전체_수익_가입자수수료 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("E7", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.전체_수익_CS관리수수료 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("E8", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.전체_수익_업무취급수수료 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("E9", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.전체_수익_사업자모델매입에따른추가수익 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("E10", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.전체_수익_유통모델매입에따른추가수익_현금_Volume = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("E11", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.전체_수익_직영매장판매수익 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("E12", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.전체_수익_소계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("E13", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set전체_비용_대리점투자비용(NullToEmpty(_WorkSheet.get_Range("E14", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set전체_비용_인건비_급여_복리후생비(NullToEmpty(_WorkSheet.get_Range("E15", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set전체_비용_임차료(NullToEmpty(_WorkSheet.get_Range("E16", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set전체_비용_이자비용(NullToEmpty(_WorkSheet.get_Range("E17", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set전체_비용_부가세(NullToEmpty(_WorkSheet.get_Range("E18", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set전체_비용_법인세(NullToEmpty(_WorkSheet.get_Range("E19", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set전체_비용_기타판매관리비(NullToEmpty(_WorkSheet.get_Range("E20", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.전체_비용_소계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("E21", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.전체손익계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("E22", Type.Missing).Value2));

            CDataControl.g_FileResultBusiness.set도매_수익_가입자관리수수료(NullToEmpty(_WorkSheet.get_Range("E28", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set도매_수익_CS관리수수료(NullToEmpty(_WorkSheet.get_Range("E29", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set도매_수익_사업자모델매입에따른추가수익(NullToEmpty(_WorkSheet.get_Range("E30", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set도매_수익_유통모델매입에따른추가수익_현금_Volume(NullToEmpty(_WorkSheet.get_Range("E31", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.도매_수익_소계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("E32", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set도매_비용_대리점투자비용(NullToEmpty(_WorkSheet.get_Range("E33", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set도매_비용_인건비_급여_복리후생비(NullToEmpty(_WorkSheet.get_Range("E34", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set도매_비용_임차료(NullToEmpty(_WorkSheet.get_Range("E35", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set도매_비용_이자비용(NullToEmpty(_WorkSheet.get_Range("E36", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set도매_비용_부가세(NullToEmpty(_WorkSheet.get_Range("E37", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set도매_비용_법인세(NullToEmpty(_WorkSheet.get_Range("E38", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set도매_비용_기타판매관리비(NullToEmpty(_WorkSheet.get_Range("E39", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.도매_비용_소계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("E40", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.도매손익계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("E41", Type.Missing).Value2));

            CDataControl.g_FileResultBusiness.set소매_수익_업무취급수수료(NullToEmpty(_WorkSheet.get_Range("E46", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set소매_수익_직영매장판매수익(NullToEmpty(_WorkSheet.get_Range("E47", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.소매_수익_소계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("E48", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set소매_비용_인건비_급여_복리후생비(NullToEmpty(_WorkSheet.get_Range("E49", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set소매_비용_임차료(NullToEmpty(_WorkSheet.get_Range("E50", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set소매_비용_이자비용(NullToEmpty(_WorkSheet.get_Range("E51", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set소매_비용_부가세(NullToEmpty(_WorkSheet.get_Range("E52", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set소매_비용_법인세(NullToEmpty(_WorkSheet.get_Range("E53", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.set소매_비용_기타판매관리비(NullToEmpty(_WorkSheet.get_Range("E54", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.소매_비용_소계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("E55", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.소매손익계 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("E56", Type.Missing).Value2));
            CDataControl.g_FileResultBusiness.점별손익추정 = StringToIntVal(NullToEmpty(_WorkSheet.get_Range("E57", Type.Missing).Value2));

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="_WorkSheet"></param>
        public static void ReadExcelFileToDataResultStore(excel.Worksheet _WorkSheet)
        {
            CDataControl.g_FileResultStoreTotal.전체_수익_가입자수수료 = StringToIntVal(_WorkSheet.get_Range("I7", Type.Missing).Value2);
            CDataControl.g_FileResultStoreTotal.전체_수익_CS관리수수료 = StringToIntVal(_WorkSheet.get_Range("I8", Type.Missing).Value2);
            CDataControl.g_FileResultStoreTotal.전체_수익_업무취급수수료 = StringToIntVal(_WorkSheet.get_Range("I9", Type.Missing).Value2);
            CDataControl.g_FileResultStoreTotal.전체_수익_사업자모델매입에따른추가수익 = StringToIntVal(_WorkSheet.get_Range("I10", Type.Missing).Value2);
            CDataControl.g_FileResultStoreTotal.전체_수익_유통모델매입에따른추가수익_현금_Volume = StringToIntVal(_WorkSheet.get_Range("I11", Type.Missing).Value2);
            CDataControl.g_FileResultStoreTotal.전체_수익_직영매장판매수익 = StringToIntVal(_WorkSheet.get_Range("I12", Type.Missing).Value2);
            CDataControl.g_FileResultStoreTotal.전체_수익_소계 = StringToIntVal(_WorkSheet.get_Range("I13", Type.Missing).Value2);
            CDataControl.g_FileResultStoreTotal.set전체_비용_대리점투자비용(NullToEmpty(_WorkSheet.get_Range("I14", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set전체_비용_인건비_급여_복리후생비(NullToEmpty(_WorkSheet.get_Range("I15", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set전체_비용_임차료(NullToEmpty(_WorkSheet.get_Range("I16", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set전체_비용_이자비용(NullToEmpty(_WorkSheet.get_Range("I17", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set전체_비용_부가세(NullToEmpty(_WorkSheet.get_Range("I18", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set전체_비용_법인세(NullToEmpty(_WorkSheet.get_Range("I19", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set전체_비용_기타판매관리비(NullToEmpty(_WorkSheet.get_Range("I20", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.전체_비용_소계 = StringToIntVal(_WorkSheet.get_Range("I21", Type.Missing).Value2);
            CDataControl.g_FileResultStoreTotal.전체손익계 = StringToIntVal(_WorkSheet.get_Range("I22", Type.Missing).Value2);

            CDataControl.g_FileResultStoreTotal.set도매_수익_가입자관리수수료(NullToEmpty(_WorkSheet.get_Range("I28", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set도매_수익_CS관리수수료(NullToEmpty(_WorkSheet.get_Range("I29", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set도매_수익_사업자모델매입에따른추가수익(NullToEmpty(_WorkSheet.get_Range("I30", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set도매_수익_유통모델매입에따른추가수익_현금_Volume(NullToEmpty(_WorkSheet.get_Range("I31", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.도매_수익_소계 = StringToIntVal(_WorkSheet.get_Range("I32", Type.Missing).Value2);
            CDataControl.g_FileResultStoreTotal.set도매_비용_대리점투자비용(NullToEmpty(_WorkSheet.get_Range("I33", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set도매_비용_인건비_급여_복리후생비(NullToEmpty(_WorkSheet.get_Range("I34", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set도매_비용_임차료(NullToEmpty(_WorkSheet.get_Range("I35", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set도매_비용_이자비용(NullToEmpty(_WorkSheet.get_Range("I36", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set도매_비용_부가세(NullToEmpty(_WorkSheet.get_Range("I37", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set도매_비용_법인세(NullToEmpty(_WorkSheet.get_Range("I38", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set도매_비용_기타판매관리비(NullToEmpty(_WorkSheet.get_Range("I39", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.도매_비용_소계 = StringToIntVal(_WorkSheet.get_Range("I40", Type.Missing).Value2);
            CDataControl.g_FileResultStoreTotal.도매손익계 = StringToIntVal(_WorkSheet.get_Range("I41", Type.Missing).Value2);

            CDataControl.g_FileResultStoreTotal.set소매_수익_업무취급수수료(NullToEmpty(_WorkSheet.get_Range("I46", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set소매_수익_직영매장판매수익(NullToEmpty(_WorkSheet.get_Range("I47", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.소매_수익_소계 = StringToIntVal(_WorkSheet.get_Range("I48", Type.Missing).Value2);
            CDataControl.g_FileResultStoreTotal.set소매_비용_인건비_급여_복리후생비(NullToEmpty(_WorkSheet.get_Range("I49", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set소매_비용_임차료(NullToEmpty(_WorkSheet.get_Range("I50", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set소매_비용_이자비용(NullToEmpty(_WorkSheet.get_Range("I51", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set소매_비용_부가세(NullToEmpty(_WorkSheet.get_Range("I52", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set소매_비용_법인세(NullToEmpty(_WorkSheet.get_Range("I53", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.set소매_비용_기타판매관리비(NullToEmpty(_WorkSheet.get_Range("I54", Type.Missing).Value2));
            CDataControl.g_FileResultStoreTotal.소매_비용_소계 = StringToIntVal(_WorkSheet.get_Range("I55", Type.Missing).Value2);
            CDataControl.g_FileResultStoreTotal.소매손익계 = StringToIntVal(_WorkSheet.get_Range("I56", Type.Missing).Value2);
            CDataControl.g_FileResultStoreTotal.점별손익추정 = StringToIntVal(_WorkSheet.get_Range("I57", Type.Missing).Value2);



            CDataControl.g_FileResultStore.전체_수익_가입자수수료 = StringToIntVal(_WorkSheet.get_Range("J7", Type.Missing).Value2);
            CDataControl.g_FileResultStore.전체_수익_CS관리수수료 = StringToIntVal(_WorkSheet.get_Range("J8", Type.Missing).Value2);
            CDataControl.g_FileResultStore.전체_수익_업무취급수수료 = StringToIntVal(_WorkSheet.get_Range("J9", Type.Missing).Value2);
            CDataControl.g_FileResultStore.전체_수익_사업자모델매입에따른추가수익 = StringToIntVal(_WorkSheet.get_Range("J10", Type.Missing).Value2);
            CDataControl.g_FileResultStore.전체_수익_유통모델매입에따른추가수익_현금_Volume = StringToIntVal(_WorkSheet.get_Range("J11", Type.Missing).Value2);
            CDataControl.g_FileResultStore.전체_수익_직영매장판매수익 = StringToIntVal(_WorkSheet.get_Range("J12", Type.Missing).Value2);
            CDataControl.g_FileResultStore.전체_수익_소계 = StringToIntVal(_WorkSheet.get_Range("J13", Type.Missing).Value2);
            CDataControl.g_FileResultStore.set전체_비용_대리점투자비용(NullToEmpty(_WorkSheet.get_Range("J14", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set전체_비용_인건비_급여_복리후생비(NullToEmpty(_WorkSheet.get_Range("J15", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set전체_비용_임차료(NullToEmpty(_WorkSheet.get_Range("J16", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set전체_비용_이자비용(NullToEmpty(_WorkSheet.get_Range("J17", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set전체_비용_부가세(NullToEmpty(_WorkSheet.get_Range("J18", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set전체_비용_법인세(NullToEmpty(_WorkSheet.get_Range("J19", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set전체_비용_기타판매관리비(NullToEmpty(_WorkSheet.get_Range("J20", Type.Missing).Value2));
            CDataControl.g_FileResultStore.전체_비용_소계 = StringToIntVal(_WorkSheet.get_Range("J21", Type.Missing).Value2);
            CDataControl.g_FileResultStore.전체손익계 = StringToIntVal(_WorkSheet.get_Range("J22", Type.Missing).Value2);

            CDataControl.g_FileResultStore.set도매_수익_가입자관리수수료(NullToEmpty(_WorkSheet.get_Range("J28", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set도매_수익_CS관리수수료(NullToEmpty(_WorkSheet.get_Range("J29", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set도매_수익_사업자모델매입에따른추가수익(NullToEmpty(_WorkSheet.get_Range("J30", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set도매_수익_유통모델매입에따른추가수익_현금_Volume(NullToEmpty(_WorkSheet.get_Range("J31", Type.Missing).Value2));
            CDataControl.g_FileResultStore.도매_수익_소계 = StringToIntVal(_WorkSheet.get_Range("J32", Type.Missing).Value2);
            CDataControl.g_FileResultStore.set도매_비용_대리점투자비용(NullToEmpty(_WorkSheet.get_Range("J33", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set도매_비용_인건비_급여_복리후생비(NullToEmpty(_WorkSheet.get_Range("J34", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set도매_비용_임차료(NullToEmpty(_WorkSheet.get_Range("J35", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set도매_비용_이자비용(NullToEmpty(_WorkSheet.get_Range("J36", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set도매_비용_부가세(NullToEmpty(_WorkSheet.get_Range("J37", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set도매_비용_법인세(NullToEmpty(_WorkSheet.get_Range("J38", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set도매_비용_기타판매관리비(NullToEmpty(_WorkSheet.get_Range("J39", Type.Missing).Value2));
            CDataControl.g_FileResultStore.도매_비용_소계 = StringToIntVal(_WorkSheet.get_Range("J40", Type.Missing).Value2);
            CDataControl.g_FileResultStore.도매손익계 = StringToIntVal(_WorkSheet.get_Range("J41", Type.Missing).Value2);

            CDataControl.g_FileResultStore.set소매_수익_업무취급수수료(NullToEmpty(_WorkSheet.get_Range("J46", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set소매_수익_직영매장판매수익(NullToEmpty(_WorkSheet.get_Range("J47", Type.Missing).Value2));
            CDataControl.g_FileResultStore.소매_수익_소계 = StringToIntVal(_WorkSheet.get_Range("J48", Type.Missing).Value2);
            CDataControl.g_FileResultStore.set소매_비용_인건비_급여_복리후생비(NullToEmpty(_WorkSheet.get_Range("J49", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set소매_비용_임차료(NullToEmpty(_WorkSheet.get_Range("J50", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set소매_비용_이자비용(NullToEmpty(_WorkSheet.get_Range("J51", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set소매_비용_부가세(NullToEmpty(_WorkSheet.get_Range("J52", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set소매_비용_법인세(NullToEmpty(_WorkSheet.get_Range("J53", Type.Missing).Value2));
            CDataControl.g_FileResultStore.set소매_비용_기타판매관리비(NullToEmpty(_WorkSheet.get_Range("J54", Type.Missing).Value2));
            CDataControl.g_FileResultStore.소매_비용_소계 = StringToIntVal(_WorkSheet.get_Range("J55", Type.Missing).Value2);
            CDataControl.g_FileResultStore.소매손익계 = StringToIntVal(_WorkSheet.get_Range("J56", Type.Missing).Value2);
            CDataControl.g_FileResultStore.점별손익추정 = StringToIntVal(_WorkSheet.get_Range("J57", Type.Missing).Value2);

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="_WorkSheet"></param>
        public static void ReadExcelFileToDataResultFuture(excel.Worksheet _WorkSheet)
        {
            CDataControl.g_FileResultFutureTotal.전체_수익_가입자수수료 = StringToIntVal(_WorkSheet.get_Range("N7", Type.Missing).Value2);
            CDataControl.g_FileResultFutureTotal.전체_수익_CS관리수수료 = StringToIntVal(_WorkSheet.get_Range("N8", Type.Missing).Value2);
            CDataControl.g_FileResultFutureTotal.전체_수익_업무취급수수료 = StringToIntVal(_WorkSheet.get_Range("N9", Type.Missing).Value2);
            CDataControl.g_FileResultFutureTotal.전체_수익_사업자모델매입에따른추가수익 = StringToIntVal(_WorkSheet.get_Range("N10", Type.Missing).Value2);
            CDataControl.g_FileResultFutureTotal.전체_수익_유통모델매입에따른추가수익_현금_Volume = StringToIntVal(_WorkSheet.get_Range("N11", Type.Missing).Value2);
            CDataControl.g_FileResultFutureTotal.전체_수익_직영매장판매수익 = StringToIntVal(_WorkSheet.get_Range("N12", Type.Missing).Value2);
            CDataControl.g_FileResultFutureTotal.전체_수익_소계 = StringToIntVal(_WorkSheet.get_Range("N13", Type.Missing).Value2);
            CDataControl.g_FileResultFutureTotal.set전체_비용_대리점투자비용(_WorkSheet.get_Range("N14", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set전체_비용_인건비_급여_복리후생비(_WorkSheet.get_Range("N15", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set전체_비용_임차료(_WorkSheet.get_Range("N16", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set전체_비용_이자비용(_WorkSheet.get_Range("N17", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set전체_비용_부가세(_WorkSheet.get_Range("N18", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set전체_비용_법인세(_WorkSheet.get_Range("N19", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set전체_비용_기타판매관리비(_WorkSheet.get_Range("N20", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.전체_비용_소계 = StringToIntVal(_WorkSheet.get_Range("N21", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.전체손익계 = StringToIntVal(_WorkSheet.get_Range("N22", Type.Missing).Value2.ToString());

            CDataControl.g_FileResultFutureTotal.set도매_수익_가입자관리수수료(_WorkSheet.get_Range("N28", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_수익_CS관리수수료(_WorkSheet.get_Range("N29", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_수익_사업자모델매입에따른추가수익(_WorkSheet.get_Range("N30", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_수익_유통모델매입에따른추가수익_현금_Volume(_WorkSheet.get_Range("N31", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.도매_수익_소계 = StringToIntVal(_WorkSheet.get_Range("N32", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_비용_대리점투자비용(_WorkSheet.get_Range("N33", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_비용_인건비_급여_복리후생비(_WorkSheet.get_Range("N34", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_비용_임차료(_WorkSheet.get_Range("N35", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_비용_이자비용(_WorkSheet.get_Range("N36", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_비용_부가세(_WorkSheet.get_Range("N37", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_비용_법인세(_WorkSheet.get_Range("N38", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set도매_비용_기타판매관리비(_WorkSheet.get_Range("N39", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.도매_비용_소계 = StringToIntVal(_WorkSheet.get_Range("N40", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.도매손익계 = StringToIntVal(_WorkSheet.get_Range("N41", Type.Missing).Value2.ToString());

            CDataControl.g_FileResultFutureTotal.set소매_수익_업무취급수수료(_WorkSheet.get_Range("N46", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set소매_수익_직영매장판매수익(_WorkSheet.get_Range("N47", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.소매_수익_소계 = StringToIntVal(_WorkSheet.get_Range("N48", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set소매_비용_인건비_급여_복리후생비(_WorkSheet.get_Range("N49", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set소매_비용_임차료(_WorkSheet.get_Range("N50", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set소매_비용_이자비용(_WorkSheet.get_Range("N51", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set소매_비용_부가세(_WorkSheet.get_Range("N52", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set소매_비용_법인세(_WorkSheet.get_Range("N53", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.set소매_비용_기타판매관리비(_WorkSheet.get_Range("N54", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.소매_비용_소계 = StringToIntVal(_WorkSheet.get_Range("N55", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.소매손익계 = StringToIntVal(_WorkSheet.get_Range("N56", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFutureTotal.점별손익추정 = StringToIntVal(_WorkSheet.get_Range("N57", Type.Missing).Value2.ToString());



            CDataControl.g_FileResultFuture.전체_수익_가입자수수료 = StringToIntVal(_WorkSheet.get_Range("O7", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.전체_수익_CS관리수수료 = StringToIntVal(_WorkSheet.get_Range("O8", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.전체_수익_업무취급수수료 = StringToIntVal(_WorkSheet.get_Range("O9", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.전체_수익_사업자모델매입에따른추가수익 = StringToIntVal(_WorkSheet.get_Range("O10", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.전체_수익_유통모델매입에따른추가수익_현금_Volume = StringToIntVal(_WorkSheet.get_Range("O11", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.전체_수익_직영매장판매수익 = StringToIntVal(_WorkSheet.get_Range("O12", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.전체_수익_소계 = StringToIntVal(_WorkSheet.get_Range("O13", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set전체_비용_대리점투자비용(_WorkSheet.get_Range("O14", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set전체_비용_인건비_급여_복리후생비(_WorkSheet.get_Range("O15", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set전체_비용_임차료(_WorkSheet.get_Range("O16", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set전체_비용_이자비용(_WorkSheet.get_Range("O17", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set전체_비용_부가세(_WorkSheet.get_Range("O18", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set전체_비용_법인세(_WorkSheet.get_Range("O19", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set전체_비용_기타판매관리비(_WorkSheet.get_Range("O20", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.전체_비용_소계 = StringToIntVal(_WorkSheet.get_Range("O21", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.전체손익계 = StringToIntVal(_WorkSheet.get_Range("O22", Type.Missing).Value2.ToString());

            CDataControl.g_FileResultFuture.set도매_수익_가입자관리수수료(_WorkSheet.get_Range("O28", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set도매_수익_CS관리수수료(_WorkSheet.get_Range("O29", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set도매_수익_사업자모델매입에따른추가수익(_WorkSheet.get_Range("O30", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set도매_수익_유통모델매입에따른추가수익_현금_Volume(_WorkSheet.get_Range("O31", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.도매_수익_소계 = StringToIntVal(_WorkSheet.get_Range("O32", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set도매_비용_대리점투자비용(_WorkSheet.get_Range("O33", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set도매_비용_인건비_급여_복리후생비(_WorkSheet.get_Range("O34", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set도매_비용_임차료(_WorkSheet.get_Range("O35", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set도매_비용_이자비용(_WorkSheet.get_Range("O36", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set도매_비용_부가세(_WorkSheet.get_Range("O37", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set도매_비용_법인세(_WorkSheet.get_Range("O38", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set도매_비용_기타판매관리비(_WorkSheet.get_Range("O39", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.도매_비용_소계 = StringToIntVal(_WorkSheet.get_Range("O40", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.도매손익계 = StringToIntVal(_WorkSheet.get_Range("O41", Type.Missing).Value2.ToString());

            CDataControl.g_FileResultFuture.set소매_수익_업무취급수수료(_WorkSheet.get_Range("O46", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set소매_수익_직영매장판매수익(_WorkSheet.get_Range("O47", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.소매_수익_소계 = StringToIntVal(_WorkSheet.get_Range("O48", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set소매_비용_인건비_급여_복리후생비(_WorkSheet.get_Range("O49", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set소매_비용_임차료(_WorkSheet.get_Range("O50", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set소매_비용_이자비용(_WorkSheet.get_Range("O51", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set소매_비용_부가세(_WorkSheet.get_Range("O52", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set소매_비용_법인세(_WorkSheet.get_Range("O53", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.set소매_비용_기타판매관리비(_WorkSheet.get_Range("O54", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.소매_비용_소계 = StringToIntVal(_WorkSheet.get_Range("O55", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.소매손익계 = StringToIntVal(_WorkSheet.get_Range("O56", Type.Missing).Value2.ToString());
            CDataControl.g_FileResultFuture.점별손익추정 = StringToIntVal(_WorkSheet.get_Range("O57", Type.Missing).Value2.ToString());

        }








        /// <summary>
        /// 
        /// </summary>
        /// <param name="_WorkSheet"></param>
        public static void WriteExcelFileToDataResultBusiness(excel.Worksheet _WorkSheet)
        {
            _WorkSheet.get_Range("D7", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.전체_수익_가입자수수료;
            _WorkSheet.get_Range("D8", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.전체_수익_CS관리수수료;
            _WorkSheet.get_Range("D9", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.전체_수익_업무취급수수료;
            _WorkSheet.get_Range("D10", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.전체_수익_사업자모델매입에따른추가수익;
            _WorkSheet.get_Range("D11", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.전체_수익_유통모델매입에따른추가수익_현금_Volume;
            _WorkSheet.get_Range("D12", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.전체_수익_직영매장판매수익;
            _WorkSheet.get_Range("D13", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.전체_수익_소계;
            _WorkSheet.get_Range("D14", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get전체_비용_대리점투자비용();
            _WorkSheet.get_Range("D15", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get전체_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("D16", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get전체_비용_임차료();
            _WorkSheet.get_Range("D17", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get전체_비용_이자비용();
            _WorkSheet.get_Range("D18", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get전체_비용_부가세();
            _WorkSheet.get_Range("D19", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get전체_비용_법인세();
            _WorkSheet.get_Range("D20", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get전체_비용_기타판매관리비();
            _WorkSheet.get_Range("D21", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.전체_비용_소계;
            _WorkSheet.get_Range("D22", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.전체손익계;

            _WorkSheet.get_Range("D28", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get도매_수익_가입자관리수수료();
            _WorkSheet.get_Range("D29", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get도매_수익_CS관리수수료();
            _WorkSheet.get_Range("D30", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get도매_수익_사업자모델매입에따른추가수익();
            _WorkSheet.get_Range("D31", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
            _WorkSheet.get_Range("D32", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.도매_수익_소계;
            _WorkSheet.get_Range("D33", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get도매_비용_대리점투자비용();
            _WorkSheet.get_Range("D34", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get도매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("D35", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get도매_비용_임차료();
            _WorkSheet.get_Range("D36", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get도매_비용_이자비용();
            _WorkSheet.get_Range("D37", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get도매_비용_부가세();
            _WorkSheet.get_Range("D38", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get도매_비용_법인세();
            _WorkSheet.get_Range("D39", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get도매_비용_기타판매관리비();
            _WorkSheet.get_Range("D40", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.도매_비용_소계;
            _WorkSheet.get_Range("D41", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.도매손익계;

            _WorkSheet.get_Range("D46", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get소매_수익_업무취급수수료();
            _WorkSheet.get_Range("D47", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get소매_수익_직영매장판매수익();
            _WorkSheet.get_Range("D48", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.소매_수익_소계;
            _WorkSheet.get_Range("D49", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get소매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("D50", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get소매_비용_임차료();
            _WorkSheet.get_Range("D51", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get소매_비용_이자비용();
            _WorkSheet.get_Range("D52", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get소매_비용_부가세();
            _WorkSheet.get_Range("D53", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get소매_비용_법인세();
            _WorkSheet.get_Range("D54", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.get소매_비용_기타판매관리비();
            _WorkSheet.get_Range("D55", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.소매_비용_소계;
            _WorkSheet.get_Range("D56", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.소매손익계;
            _WorkSheet.get_Range("D57", Type.Missing).Value2 = CDataControl.g_FileResultBusinessTotal.점별손익추정;



            _WorkSheet.get_Range("E7", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.전체_수익_가입자수수료;
            _WorkSheet.get_Range("E8", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.전체_수익_CS관리수수료;
            _WorkSheet.get_Range("E9", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.전체_수익_업무취급수수료;
            _WorkSheet.get_Range("E10", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.전체_수익_사업자모델매입에따른추가수익;
            _WorkSheet.get_Range("E11", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.전체_수익_유통모델매입에따른추가수익_현금_Volume;
            _WorkSheet.get_Range("E12", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.전체_수익_직영매장판매수익;
            _WorkSheet.get_Range("E13", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.전체_수익_소계;
            _WorkSheet.get_Range("E14", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get전체_비용_대리점투자비용();
            _WorkSheet.get_Range("E15", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get전체_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("E16", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get전체_비용_임차료();
            _WorkSheet.get_Range("E17", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get전체_비용_이자비용();
            _WorkSheet.get_Range("E18", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get전체_비용_부가세();
            _WorkSheet.get_Range("E19", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get전체_비용_법인세();
            _WorkSheet.get_Range("E20", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get전체_비용_기타판매관리비();
            _WorkSheet.get_Range("E21", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.전체_비용_소계;
            _WorkSheet.get_Range("E22", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.전체손익계;

            _WorkSheet.get_Range("E28", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get도매_수익_가입자관리수수료();
            _WorkSheet.get_Range("E29", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get도매_수익_CS관리수수료();
            _WorkSheet.get_Range("E30", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get도매_수익_사업자모델매입에따른추가수익();
            _WorkSheet.get_Range("E31", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
            _WorkSheet.get_Range("E32", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.도매_수익_소계;
            _WorkSheet.get_Range("E33", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get도매_비용_대리점투자비용();
            _WorkSheet.get_Range("E34", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get도매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("E35", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get도매_비용_임차료();
            _WorkSheet.get_Range("E36", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get도매_비용_이자비용();
            _WorkSheet.get_Range("E37", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get도매_비용_부가세();
            _WorkSheet.get_Range("E38", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get도매_비용_법인세();
            _WorkSheet.get_Range("E39", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get도매_비용_기타판매관리비();
            _WorkSheet.get_Range("E40", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.도매_비용_소계;
            _WorkSheet.get_Range("E41", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.도매손익계;

            _WorkSheet.get_Range("E46", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get소매_수익_업무취급수수료();
            _WorkSheet.get_Range("E47", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get소매_수익_직영매장판매수익();
            _WorkSheet.get_Range("E48", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.소매_수익_소계;
            _WorkSheet.get_Range("E49", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get소매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("E50", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get소매_비용_임차료();
            _WorkSheet.get_Range("E51", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get소매_비용_이자비용();
            _WorkSheet.get_Range("E52", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get소매_비용_부가세();
            _WorkSheet.get_Range("E53", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get소매_비용_법인세();
            _WorkSheet.get_Range("E54", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.get소매_비용_기타판매관리비();
            _WorkSheet.get_Range("E55", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.소매_비용_소계;
            _WorkSheet.get_Range("E56", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.소매손익계;
            _WorkSheet.get_Range("E57", Type.Missing).Value2 = CDataControl.g_FileResultBusiness.점별손익추정;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="_WorkSheet"></param>
        public static void WriteExcelFileToDataResultStore(excel.Worksheet _WorkSheet)
        {
            _WorkSheet.get_Range("I7", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.전체_수익_가입자수수료;
            _WorkSheet.get_Range("I8", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.전체_수익_CS관리수수료;
            _WorkSheet.get_Range("I9", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.전체_수익_업무취급수수료;
            _WorkSheet.get_Range("I10", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.전체_수익_사업자모델매입에따른추가수익;
            _WorkSheet.get_Range("I11", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.전체_수익_유통모델매입에따른추가수익_현금_Volume;
            _WorkSheet.get_Range("I12", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.전체_수익_직영매장판매수익;
            _WorkSheet.get_Range("I13", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.전체_수익_소계;
            _WorkSheet.get_Range("I14", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get전체_비용_대리점투자비용();
            _WorkSheet.get_Range("I15", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get전체_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("I16", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get전체_비용_임차료();
            _WorkSheet.get_Range("I17", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get전체_비용_이자비용();
            _WorkSheet.get_Range("I18", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get전체_비용_부가세();
            _WorkSheet.get_Range("I19", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get전체_비용_법인세();
            _WorkSheet.get_Range("I20", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get전체_비용_기타판매관리비();
            _WorkSheet.get_Range("I21", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.전체_비용_소계;
            _WorkSheet.get_Range("I22", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.전체손익계;

            _WorkSheet.get_Range("I28", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get도매_수익_가입자관리수수료();
            _WorkSheet.get_Range("I29", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get도매_수익_CS관리수수료();
            _WorkSheet.get_Range("I30", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get도매_수익_사업자모델매입에따른추가수익();
            _WorkSheet.get_Range("I31", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
            _WorkSheet.get_Range("I32", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.도매_수익_소계;
            _WorkSheet.get_Range("I33", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get도매_비용_대리점투자비용();
            _WorkSheet.get_Range("I34", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get도매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("I35", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get도매_비용_임차료();
            _WorkSheet.get_Range("I36", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get도매_비용_이자비용();
            _WorkSheet.get_Range("I37", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get도매_비용_부가세();
            _WorkSheet.get_Range("I38", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get도매_비용_법인세();
            _WorkSheet.get_Range("I39", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get도매_비용_기타판매관리비();
            _WorkSheet.get_Range("I40", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.도매_비용_소계;
            _WorkSheet.get_Range("I41", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.도매손익계;

            _WorkSheet.get_Range("I46", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get소매_수익_업무취급수수료();
            _WorkSheet.get_Range("I47", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get소매_수익_직영매장판매수익();
            _WorkSheet.get_Range("I48", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.소매_수익_소계;
            _WorkSheet.get_Range("I49", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get소매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("I50", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get소매_비용_임차료();
            _WorkSheet.get_Range("I51", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get소매_비용_이자비용();
            _WorkSheet.get_Range("I52", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get소매_비용_부가세();
            _WorkSheet.get_Range("I53", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get소매_비용_법인세();
            _WorkSheet.get_Range("I54", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.get소매_비용_기타판매관리비();
            _WorkSheet.get_Range("I55", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.소매_비용_소계;
            _WorkSheet.get_Range("I56", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.소매손익계;
            _WorkSheet.get_Range("I57", Type.Missing).Value2 = CDataControl.g_ResultStoreTotal.점별손익추정;



            _WorkSheet.get_Range("J7", Type.Missing).Value2 = CDataControl.g_ResultStore.전체_수익_가입자수수료;
            _WorkSheet.get_Range("J8", Type.Missing).Value2 = CDataControl.g_ResultStore.전체_수익_CS관리수수료;
            _WorkSheet.get_Range("J9", Type.Missing).Value2 = CDataControl.g_ResultStore.전체_수익_업무취급수수료;
            _WorkSheet.get_Range("J10", Type.Missing).Value2 = CDataControl.g_ResultStore.전체_수익_사업자모델매입에따른추가수익;
            _WorkSheet.get_Range("J11", Type.Missing).Value2 = CDataControl.g_ResultStore.전체_수익_유통모델매입에따른추가수익_현금_Volume;
            _WorkSheet.get_Range("J12", Type.Missing).Value2 = CDataControl.g_ResultStore.전체_수익_직영매장판매수익;
            _WorkSheet.get_Range("J13", Type.Missing).Value2 = CDataControl.g_ResultStore.전체_수익_소계;
            _WorkSheet.get_Range("J14", Type.Missing).Value2 = CDataControl.g_ResultStore.get전체_비용_대리점투자비용();
            _WorkSheet.get_Range("J15", Type.Missing).Value2 = CDataControl.g_ResultStore.get전체_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("J16", Type.Missing).Value2 = CDataControl.g_ResultStore.get전체_비용_임차료();
            _WorkSheet.get_Range("J17", Type.Missing).Value2 = CDataControl.g_ResultStore.get전체_비용_이자비용();
            _WorkSheet.get_Range("J18", Type.Missing).Value2 = CDataControl.g_ResultStore.get전체_비용_부가세();
            _WorkSheet.get_Range("J19", Type.Missing).Value2 = CDataControl.g_ResultStore.get전체_비용_법인세();
            _WorkSheet.get_Range("J20", Type.Missing).Value2 = CDataControl.g_ResultStore.get전체_비용_기타판매관리비();
            _WorkSheet.get_Range("J21", Type.Missing).Value2 = CDataControl.g_ResultStore.전체_비용_소계;
            _WorkSheet.get_Range("J22", Type.Missing).Value2 = CDataControl.g_ResultStore.전체손익계;

            _WorkSheet.get_Range("J28", Type.Missing).Value2 = CDataControl.g_ResultStore.get도매_수익_가입자관리수수료();
            _WorkSheet.get_Range("J29", Type.Missing).Value2 = CDataControl.g_ResultStore.get도매_수익_CS관리수수료();
            _WorkSheet.get_Range("J30", Type.Missing).Value2 = CDataControl.g_ResultStore.get도매_수익_사업자모델매입에따른추가수익();
            _WorkSheet.get_Range("J31", Type.Missing).Value2 = CDataControl.g_ResultStore.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
            _WorkSheet.get_Range("J32", Type.Missing).Value2 = CDataControl.g_ResultStore.도매_수익_소계;
            _WorkSheet.get_Range("J33", Type.Missing).Value2 = CDataControl.g_ResultStore.get도매_비용_대리점투자비용();
            _WorkSheet.get_Range("J34", Type.Missing).Value2 = CDataControl.g_ResultStore.get도매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("J35", Type.Missing).Value2 = CDataControl.g_ResultStore.get도매_비용_임차료();
            _WorkSheet.get_Range("J36", Type.Missing).Value2 = CDataControl.g_ResultStore.get도매_비용_이자비용();
            _WorkSheet.get_Range("J37", Type.Missing).Value2 = CDataControl.g_ResultStore.get도매_비용_부가세();
            _WorkSheet.get_Range("J38", Type.Missing).Value2 = CDataControl.g_ResultStore.get도매_비용_법인세();
            _WorkSheet.get_Range("J39", Type.Missing).Value2 = CDataControl.g_ResultStore.get도매_비용_기타판매관리비();
            _WorkSheet.get_Range("J40", Type.Missing).Value2 = CDataControl.g_ResultStore.도매_비용_소계;
            _WorkSheet.get_Range("J41", Type.Missing).Value2 = CDataControl.g_ResultStore.도매손익계;

            _WorkSheet.get_Range("J46", Type.Missing).Value2 = CDataControl.g_ResultStore.get소매_수익_업무취급수수료();
            _WorkSheet.get_Range("J47", Type.Missing).Value2 = CDataControl.g_ResultStore.get소매_수익_직영매장판매수익();
            _WorkSheet.get_Range("J48", Type.Missing).Value2 = CDataControl.g_ResultStore.소매_수익_소계;
            _WorkSheet.get_Range("J49", Type.Missing).Value2 = CDataControl.g_ResultStore.get소매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("J50", Type.Missing).Value2 = CDataControl.g_ResultStore.get소매_비용_임차료();
            _WorkSheet.get_Range("J51", Type.Missing).Value2 = CDataControl.g_ResultStore.get소매_비용_이자비용();
            _WorkSheet.get_Range("J52", Type.Missing).Value2 = CDataControl.g_ResultStore.get소매_비용_부가세();
            _WorkSheet.get_Range("J53", Type.Missing).Value2 = CDataControl.g_ResultStore.get소매_비용_법인세();
            _WorkSheet.get_Range("J54", Type.Missing).Value2 = CDataControl.g_ResultStore.get소매_비용_기타판매관리비();
            _WorkSheet.get_Range("J55", Type.Missing).Value2 = CDataControl.g_ResultStore.소매_비용_소계;
            _WorkSheet.get_Range("J56", Type.Missing).Value2 = CDataControl.g_ResultStore.소매손익계;
            _WorkSheet.get_Range("J57", Type.Missing).Value2 = CDataControl.g_ResultStore.점별손익추정;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="_WorkSheet"></param>
        public static void WriteExcelFileToDataResultFuture(excel.Worksheet _WorkSheet)
        {
            _WorkSheet.get_Range("N7", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.전체_수익_가입자수수료;
            _WorkSheet.get_Range("N8", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.전체_수익_CS관리수수료;
            _WorkSheet.get_Range("N9", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.전체_수익_업무취급수수료;
            _WorkSheet.get_Range("N10", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.전체_수익_사업자모델매입에따른추가수익;
            _WorkSheet.get_Range("N11", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.전체_수익_유통모델매입에따른추가수익_현금_Volume;
            _WorkSheet.get_Range("N12", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.전체_수익_직영매장판매수익;
            _WorkSheet.get_Range("N13", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.전체_수익_소계;
            _WorkSheet.get_Range("N14", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get전체_비용_대리점투자비용();
            _WorkSheet.get_Range("N15", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get전체_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("N16", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get전체_비용_임차료();
            _WorkSheet.get_Range("N17", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get전체_비용_이자비용();
            _WorkSheet.get_Range("N18", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get전체_비용_부가세();
            _WorkSheet.get_Range("N19", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get전체_비용_법인세();
            _WorkSheet.get_Range("N20", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get전체_비용_기타판매관리비();
            _WorkSheet.get_Range("N21", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.전체_비용_소계;
            _WorkSheet.get_Range("N22", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.전체손익계;

            _WorkSheet.get_Range("N28", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get도매_수익_가입자관리수수료();
            _WorkSheet.get_Range("N29", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get도매_수익_CS관리수수료();
            _WorkSheet.get_Range("N30", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get도매_수익_사업자모델매입에따른추가수익();
            _WorkSheet.get_Range("N31", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
            _WorkSheet.get_Range("N32", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.도매_수익_소계;
            _WorkSheet.get_Range("N33", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get도매_비용_대리점투자비용();
            _WorkSheet.get_Range("N34", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get도매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("N35", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get도매_비용_임차료();
            _WorkSheet.get_Range("N36", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get도매_비용_이자비용();
            _WorkSheet.get_Range("N37", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get도매_비용_부가세();
            _WorkSheet.get_Range("N38", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get도매_비용_법인세();
            _WorkSheet.get_Range("N39", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get도매_비용_기타판매관리비();
            _WorkSheet.get_Range("N40", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.도매_비용_소계;
            _WorkSheet.get_Range("N41", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.도매손익계;

            _WorkSheet.get_Range("N46", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get소매_수익_업무취급수수료();
            _WorkSheet.get_Range("N47", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get소매_수익_직영매장판매수익();
            _WorkSheet.get_Range("N48", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.소매_수익_소계;
            _WorkSheet.get_Range("N49", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get소매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("N50", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get소매_비용_임차료();
            _WorkSheet.get_Range("N51", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get소매_비용_이자비용();
            _WorkSheet.get_Range("N52", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get소매_비용_부가세();
            _WorkSheet.get_Range("N53", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get소매_비용_법인세();
            _WorkSheet.get_Range("N54", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.get소매_비용_기타판매관리비();
            _WorkSheet.get_Range("N55", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.소매_비용_소계;
            _WorkSheet.get_Range("N56", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.소매손익계;
            _WorkSheet.get_Range("N57", Type.Missing).Value2 = CDataControl.g_ResultFutureTotal.점별손익추정;



            _WorkSheet.get_Range("O7", Type.Missing).Value2 = CDataControl.g_ResultFuture.전체_수익_가입자수수료;
            _WorkSheet.get_Range("O8", Type.Missing).Value2 = CDataControl.g_ResultFuture.전체_수익_CS관리수수료;
            _WorkSheet.get_Range("O9", Type.Missing).Value2 = CDataControl.g_ResultFuture.전체_수익_업무취급수수료;
            _WorkSheet.get_Range("O10", Type.Missing).Value2 = CDataControl.g_ResultFuture.전체_수익_사업자모델매입에따른추가수익;
            _WorkSheet.get_Range("O11", Type.Missing).Value2 = CDataControl.g_ResultFuture.전체_수익_유통모델매입에따른추가수익_현금_Volume;
            _WorkSheet.get_Range("O12", Type.Missing).Value2 = CDataControl.g_ResultFuture.전체_수익_직영매장판매수익;
            _WorkSheet.get_Range("O13", Type.Missing).Value2 = CDataControl.g_ResultFuture.전체_수익_소계;
            _WorkSheet.get_Range("O14", Type.Missing).Value2 = CDataControl.g_ResultFuture.get전체_비용_대리점투자비용();
            _WorkSheet.get_Range("O15", Type.Missing).Value2 = CDataControl.g_ResultFuture.get전체_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("O16", Type.Missing).Value2 = CDataControl.g_ResultFuture.get전체_비용_임차료();
            _WorkSheet.get_Range("O17", Type.Missing).Value2 = CDataControl.g_ResultFuture.get전체_비용_이자비용();
            _WorkSheet.get_Range("O18", Type.Missing).Value2 = CDataControl.g_ResultFuture.get전체_비용_부가세();
            _WorkSheet.get_Range("O19", Type.Missing).Value2 = CDataControl.g_ResultFuture.get전체_비용_법인세();
            _WorkSheet.get_Range("O20", Type.Missing).Value2 = CDataControl.g_ResultFuture.get전체_비용_기타판매관리비();
            _WorkSheet.get_Range("O21", Type.Missing).Value2 = CDataControl.g_ResultFuture.전체_비용_소계;
            _WorkSheet.get_Range("O22", Type.Missing).Value2 = CDataControl.g_ResultFuture.전체손익계;

            _WorkSheet.get_Range("O28", Type.Missing).Value2 = CDataControl.g_ResultFuture.get도매_수익_가입자관리수수료();
            _WorkSheet.get_Range("O29", Type.Missing).Value2 = CDataControl.g_ResultFuture.get도매_수익_CS관리수수료();
            _WorkSheet.get_Range("O30", Type.Missing).Value2 = CDataControl.g_ResultFuture.get도매_수익_사업자모델매입에따른추가수익();
            _WorkSheet.get_Range("O31", Type.Missing).Value2 = CDataControl.g_ResultFuture.get도매_수익_유통모델매입에따른추가수익_현금_Volume();
            _WorkSheet.get_Range("O32", Type.Missing).Value2 = CDataControl.g_ResultFuture.도매_수익_소계;
            _WorkSheet.get_Range("O33", Type.Missing).Value2 = CDataControl.g_ResultFuture.get도매_비용_대리점투자비용();
            _WorkSheet.get_Range("O34", Type.Missing).Value2 = CDataControl.g_ResultFuture.get도매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("O35", Type.Missing).Value2 = CDataControl.g_ResultFuture.get도매_비용_임차료();
            _WorkSheet.get_Range("O36", Type.Missing).Value2 = CDataControl.g_ResultFuture.get도매_비용_이자비용();
            _WorkSheet.get_Range("O37", Type.Missing).Value2 = CDataControl.g_ResultFuture.get도매_비용_부가세();
            _WorkSheet.get_Range("O38", Type.Missing).Value2 = CDataControl.g_ResultFuture.get도매_비용_법인세();
            _WorkSheet.get_Range("O39", Type.Missing).Value2 = CDataControl.g_ResultFuture.get도매_비용_기타판매관리비();
            _WorkSheet.get_Range("O40", Type.Missing).Value2 = CDataControl.g_ResultFuture.도매_비용_소계;
            _WorkSheet.get_Range("O41", Type.Missing).Value2 = CDataControl.g_ResultFuture.도매손익계;

            _WorkSheet.get_Range("O46", Type.Missing).Value2 = CDataControl.g_ResultFuture.get소매_수익_업무취급수수료();
            _WorkSheet.get_Range("O47", Type.Missing).Value2 = CDataControl.g_ResultFuture.get소매_수익_직영매장판매수익();
            _WorkSheet.get_Range("O48", Type.Missing).Value2 = CDataControl.g_ResultFuture.소매_수익_소계;
            _WorkSheet.get_Range("O49", Type.Missing).Value2 = CDataControl.g_ResultFuture.get소매_비용_인건비_급여_복리후생비();
            _WorkSheet.get_Range("O50", Type.Missing).Value2 = CDataControl.g_ResultFuture.get소매_비용_임차료();
            _WorkSheet.get_Range("O51", Type.Missing).Value2 = CDataControl.g_ResultFuture.get소매_비용_이자비용();
            _WorkSheet.get_Range("O52", Type.Missing).Value2 = CDataControl.g_ResultFuture.get소매_비용_부가세();
            _WorkSheet.get_Range("O53", Type.Missing).Value2 = CDataControl.g_ResultFuture.get소매_비용_법인세();
            _WorkSheet.get_Range("O54", Type.Missing).Value2 = CDataControl.g_ResultFuture.get소매_비용_기타판매관리비();
            _WorkSheet.get_Range("O55", Type.Missing).Value2 = CDataControl.g_ResultFuture.소매_비용_소계;
            _WorkSheet.get_Range("O56", Type.Missing).Value2 = CDataControl.g_ResultFuture.소매손익계;
            _WorkSheet.get_Range("O57", Type.Missing).Value2 = CDataControl.g_ResultFuture.점별손익추정;


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
    }
}
