using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace KIWI
{
    public partial class CDataControl
    {
        public static CBasicInput g_BasicInput = new CBasicInput();     //기본입력
        public static CBusinessData g_DetailInput = new CBusinessData();  //상세입력
        
        //설정된 업계평균은 레지스트리 또는 파일에 항상 가지고 있어야 함.
        //파일 실행시 레지스트리 또는 파일에서 읽어 변수에 세팅
        public static CBusinessData g_BusinessAvg = new CBusinessData();  //업계평균, 관리자가 배포한 데이타 및 현재 클라이언트 계산에 적용하는 값

        public static CResultData g_ResultBusinessTotal = new CResultData();    //업계 총계
        public static CResultData g_ResultBusiness = new CResultData();         //업계 단위금액
        public static CResultData g_ResultStoreTotal = new CResultData();       //당대리점 총계
        public static CResultData g_ResultStore = new CResultData();            //당대리점 단위금액
        public static CResultData g_ResultFutureTotal = new CResultData();      //미래수익 총계
        public static CResultData g_ResultFuture = new CResultData();           //미래수익 단위금액

        public static CBasicInput g_SimBasicInput = new CBasicInput();     //시뮬레이션 기본입력
        public static CBusinessData g_SimDetailInput = new CBusinessData();  //시뮬레이션 상세입력

        public static CResultData g_SimResultBusinessTotal = new CResultData();    //시뮬레이션 업계 총계
        public static CResultData g_SimResultBusiness = new CResultData();         //시뮬레이션 업계 단위금액
        public static CResultData g_SimResultStoreTotal = new CResultData();       //시뮬레이션 당대리점 총계
        public static CResultData g_SimResultStore = new CResultData();            //시뮬레이션 당대리점 단위금액
        public static CResultData g_SimResultFutureTotal = new CResultData();      //시뮬레이션 미래수익 총계
        public static CResultData g_SimResultFuture = new CResultData();           //시뮬레이션 미래수익 단위금액

        //파일에서 읽은 데이터 저장 용
        //기본입력(g_BasicInput)과 상세입력(g_DetailInput)에 입력하여 현재세팅된 업계평균을 적용하여 결과를 계산할 수 있다
        //계산 없이 결과를 보기위해 g_FileResultBusinessTotal등의 아래값을 출력하면 됨
        public static CBasicInput g_FileBasicInput = new CBasicInput();     //파일에서 읽은 기본입력
        public static CBusinessData g_FileDetailInput = new CBusinessData();  //파일에서 읽은 상세입력
        //public static CBusinessData g_FileBusinessAvg;  //파일에서 읽은 업계평균
        public static CResultData g_FileResultBusinessTotal = new CResultData();    //업계 총계
        public static CResultData g_FileResultBusiness = new CResultData();         //업계 단위금액
        public static CResultData g_FileResultStoreTotal = new CResultData();       //당대리점 총계
        public static CResultData g_FileResultStore = new CResultData();            //당대리점 단위금액
        public static CResultData g_FileResultFutureTotal = new CResultData();      //미래수익 총계
        public static CResultData g_FileResultFuture = new CResultData();           //미래수익 단위금액


        //보고서용 데이터
        public static CReportData g_ReportData = new CReportData();         // 리포트에 추가적으로 들어갈 자료(이름, 코멘트 등)

        internal static String getSplitedLGEFileFromData(String splitter)
        {
            Object[] arrWarp = new Object[]{
                CDataControl.g_ReportData.getArrData(),
                CDataControl.g_BasicInput.getArrData(),
                CDataControl.g_DetailInput.getArrData(),
                CDataControl.g_ResultBusinessTotal.getArrayOutput전체(),
                CDataControl.g_ResultBusiness.getArrayOutput전체(),
                CDataControl.g_ResultStoreTotal.getArrayOutput전체(),
                CDataControl.g_ResultStore.getArrayOutput전체(),
                CDataControl.g_ResultFutureTotal.getArrayOutput전체(),
                CDataControl.g_ResultFuture.getArrayOutput전체()
            };
            return getSplitedLGEFileFromArray(arrWarp, splitter);
        }

        internal static String getSplitedLGEFileFromSimulData(String splitter)
        {
            Object[] arrWarp = new Object[]{
                CDataControl.g_ReportData.getArrData(),
                CDataControl.g_SimBasicInput.getArrData(),
                CDataControl.g_SimDetailInput.getArrData(),
                CDataControl.g_SimResultBusinessTotal.getArrayOutput전체(),
                CDataControl.g_SimResultBusiness.getArrayOutput전체(),
                CDataControl.g_SimResultStoreTotal.getArrayOutput전체(),
                CDataControl.g_SimResultStore.getArrayOutput전체(),
                CDataControl.g_SimResultFutureTotal.getArrayOutput전체(),
                CDataControl.g_SimResultFuture.getArrayOutput전체()
            };
            return getSplitedLGEFileFromArray(arrWarp, splitter);
        }

        private static String getSplitedLGEFileFromArray(Object[] arrWarp, String splitter)
        {
            string returnLge = "";
            for (int i = 0; i < arrWarp.Length; i++)
            {
                if (arrWarp[i].GetType() == Type.GetType("System.String[]"))
                {
                    foreach (String str in (arrWarp[i] as String[]))
                    {
                        returnLge += str.Replace(splitter, splitter == "|" ? "l" : "") + splitter; // 파이프를 구분자로 쓰기위해 엘(L)소문자로 고침
                    }
                }
                else if (arrWarp[i].GetType() == Type.GetType("System.Int64[]"))
                {
                    foreach (Int64 val in (arrWarp[i] as Int64[]))
                    {
                        returnLge += val.ToString() + splitter;
                    }
                }
            }
            return returnLge;
        }

        internal static void setDataFromLGEFile(String lge, String spliter)
        {
            String[] splittedLge = lge.Split(spliter.ToCharArray());

            int startIndex = 0;
            int length = 6;
            String[] param = splittedLge.Take(length).ToArray<String>();
            CDataControl.g_ReportData.setArrData(param);
            startIndex += length;
            length = 14;
            param = splittedLge.Skip(startIndex).Take(length).ToArray<String>();
            CDataControl.g_FileBasicInput.setArrData(param);
            startIndex += length;
            length = 31;
            param = splittedLge.Skip(startIndex).Take(length).ToArray<String>();
            CDataControl.g_FileDetailInput.setArrData(param);
            startIndex += length;
            length = 42;
            param = splittedLge.Skip(startIndex).Take(length).ToArray<String>();
            CDataControl.g_FileResultBusinessTotal.setArrayOutput전체(param);
            startIndex += length;
            param = splittedLge.Skip(startIndex).Take(length).ToArray<String>();
            CDataControl.g_FileResultBusiness.setArrayOutput전체(param);
            startIndex += length;
            param = splittedLge.Skip(startIndex).Take(length).ToArray<String>();
            CDataControl.g_FileResultStoreTotal.setArrayOutput전체(param);
            startIndex += length;
            param = splittedLge.Skip(startIndex).Take(length).ToArray<String>();
            CDataControl.g_FileResultStore.setArrayOutput전체(param);
            startIndex += length;
            param = splittedLge.Skip(startIndex).Take(length).ToArray<String>();
            CDataControl.g_FileResultFutureTotal.setArrayOutput전체(param);
            startIndex += length;
            param = splittedLge.Skip(startIndex).Take(length).ToArray<String>();
            CDataControl.g_FileResultFuture.setArrayOutput전체(param);

            CDataControl.g_FileBasicInput.set지역(CDataControl.g_ReportData.get지역());
            CDataControl.g_FileBasicInput.set대리점(CDataControl.g_ReportData.get대리점());
            CDataControl.g_FileBasicInput.set마케터(CDataControl.g_ReportData.get판매자());
        }
        
        internal static String getAdminDataBySerialization(String splitter)
        {
            string returnLge = "";
            foreach (Int64 val in g_BusinessAvg.getArrData())
            {
                returnLge += val.ToString() + splitter;
            }
            return returnLge;
        }

        internal static void setAdminDataFromLGEFile(String lge, String spliter)
        {
            String[] splittedLge = lge.Split(spliter.ToCharArray());
            CDataControl.g_BusinessAvg.setArrData_DetailInput(splittedLge);
        }
    }
}
