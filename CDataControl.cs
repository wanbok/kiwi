using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
    }
}
