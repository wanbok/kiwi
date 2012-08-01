using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace KIWI
{
    public class CBusinessData
    {
        //IO
        //개별 항목 Int64, String 입출력
        //전체 항목 Int64배열 입출력 길이31

        //입력중 상세입력에 변수로 사용 가능
        //업계평균에 사용 가능

        private Int64 도매_수익_월평균관리수수료;
        public void set도매_수익_월평균관리수수료(Int64 value)
        {
            도매_수익_월평균관리수수료 = value;
        }
        public Int64 get도매_수익_월평균관리수수료()
        {
            return 도매_수익_월평균관리수수료;
        }
        public void set도매_수익_월평균관리수수료(String value)
        {
            도매_수익_월평균관리수수료 = getFormatInt64(value);
        }
        public String getstr도매_수익_월평균관리수수료()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_수익_월평균관리수수료);
        }


        private Int64 도매_수익_CS관리수수료;
        public void set도매_수익_CS관리수수료(Int64 value)
        {
            도매_수익_CS관리수수료 = value;
        }
        public Int64 get도매_수익_CS관리수수료()
        {
            return 도매_수익_CS관리수수료;
        }
        public void set도매_수익_CS관리수수료(String value)
        {
            도매_수익_CS관리수수료 = getFormatInt64(value);
        }
        public String getstr도매_수익_CS관리수수료()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_수익_CS관리수수료);
        }
        public Int64 get도매_수익_CS관리수수료_분기()
        {
            return 도매_수익_CS관리수수료*CommonUtil.QUARTER;
        }
        public String getstr도매_수익_CS관리수수료_분기()
        {//   ','가 적용된 값 리턴
            return getFormatString(get도매_수익_CS관리수수료_분기());
        }


        private Int64 도매_수익_사업자모델매입관련추가수익;
        public void set도매_수익_사업자모델매입관련추가수익(Int64 value)
        {
            도매_수익_사업자모델매입관련추가수익 = value;
        }
        public Int64 get도매_수익_사업자모델매입관련추가수익()
        {
            return 도매_수익_사업자모델매입관련추가수익;
        }
        public void set도매_수익_사업자모델매입관련추가수익(String value)
        {
            도매_수익_사업자모델매입관련추가수익 = getFormatInt64(value);
        }
        public String getstr도매_수익_사업자모델매입관련추가수익()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_수익_사업자모델매입관련추가수익);
        }

        private Int64 도매_수익_유통모델매입관련추가수익_현금DC;
        public void set도매_수익_유통모델매입관련추가수익_현금DC(Int64 value)
        {
            도매_수익_유통모델매입관련추가수익_현금DC = value;
        }
        public Int64 get도매_수익_유통모델매입관련추가수익_현금DC()
        {
            return 도매_수익_유통모델매입관련추가수익_현금DC;
        }
        public void set도매_수익_유통모델매입관련추가수익_현금DC(String value)
        {
            도매_수익_유통모델매입관련추가수익_현금DC = getFormatInt64(value);
        }
        public String getstr도매_수익_유통모델매입관련추가수익_현금DC()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_수익_유통모델매입관련추가수익_현금DC);
        }

        private Int64 도매_수익_유통모델매입관련추가수익_VolumeDC;
        public void set도매_수익_유통모델매입관련추가수익_VolumeDC(Int64 value)
        {
            도매_수익_유통모델매입관련추가수익_VolumeDC = value;
        }
        public Int64 get도매_수익_유통모델매입관련추가수익_VolumeDC()
        {
            return 도매_수익_유통모델매입관련추가수익_VolumeDC;
        }
        public void set도매_수익_유통모델매입관련추가수익_VolumeDC(String value)
        {
            도매_수익_유통모델매입관련추가수익_VolumeDC = getFormatInt64(value);
        }
        public String getstr도매_수익_유통모델매입관련추가수익_VolumeDC()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_수익_유통모델매입관련추가수익_VolumeDC);
        }


        private Int64 도매_비용_대리점투자금액_신규;
        public void set도매_비용_대리점투자금액_신규(Int64 value)
        {
            도매_비용_대리점투자금액_신규 = value;
        }
        public Int64 get도매_비용_대리점투자금액_신규()
        {
            return 도매_비용_대리점투자금액_신규;
        }
        public void set도매_비용_대리점투자금액_신규(String value)
        {
            도매_비용_대리점투자금액_신규 = getFormatInt64(value);
        }
        public String getstr도매_비용_대리점투자금액_신규()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_대리점투자금액_신규);
        }

        private Int64 도매_비용_대리점투자금액_기변;
        public void set도매_비용_대리점투자금액_기변(Int64 value)
        {
            도매_비용_대리점투자금액_기변 = value;
        }
        public Int64 get도매_비용_대리점투자금액_기변()
        {
            return 도매_비용_대리점투자금액_기변;
        }
        public void set도매_비용_대리점투자금액_기변(String value)
        {
            도매_비용_대리점투자금액_기변 = getFormatInt64(value);
        }
        public String getstr도매_비용_대리점투자금액_기변()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_대리점투자금액_기변);
        }

        private Int64 도매_비용_직원급여_간부급;
        public void set도매_비용_직원급여_간부급(Int64 value)
        {
            도매_비용_직원급여_간부급 = value;
        }
        public Int64 get도매_비용_직원급여_간부급()
        {
            return 도매_비용_직원급여_간부급;
        }
        public void set도매_비용_직원급여_간부급(String value)
        {
            도매_비용_직원급여_간부급 = getFormatInt64(value);
        }
        public String getstr도매_비용_직원급여_간부급()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_직원급여_간부급);
        }
        public Int64 get도매_비용_직원급여_간부급_총액(Int64 도매_간부수)
        {
            return CommonUtil.StringToIntVal(도매_비용_직원급여_간부급.ToString()) * 도매_간부수;
        }
        public String getstr도매_비용_직원급여_간부급_총액(Int64 도매_간부수)
        {//   ','가 적용된 값 리턴
            return getFormatString(get도매_비용_직원급여_간부급_총액(도매_간부수));
        }

        private Int64 도매_비용_직원급여_평사원;
        public void set도매_비용_직원급여_평사원(Int64 value)
        {
            도매_비용_직원급여_평사원 = value;
        }
        public Int64 get도매_비용_직원급여_평사원()
        {
            return 도매_비용_직원급여_평사원;
        }
        public void set도매_비용_직원급여_평사원(String value)
        {
            도매_비용_직원급여_평사원 = getFormatInt64(value);
        }
        public String getstr도매_비용_직원급여_평사원()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_직원급여_평사원);
        }
        public Int64 get도매_비용_직원급여_평사원_총액(Int64 도매_평사원)
        {
            return CommonUtil.StringToIntVal(도매_비용_직원급여_평사원.ToString()) * 도매_평사원;
        }
        public String getstr도매_비용_직원급여_평사원_총액(Int64 도매_평사원)
        {//   ','가 적용된 값 리턴
            return getFormatString(get도매_비용_직원급여_평사원_총액(도매_평사원));
        }


        private Int64 도매_비용_지급임차료;
        public void set도매_비용_지급임차료(Int64 value)
        {
            도매_비용_지급임차료 = value;
        }
        public Int64 get도매_비용_지급임차료()
        {
            return 도매_비용_지급임차료;
        }
        public void set도매_비용_지급임차료(String value)
        {
            도매_비용_지급임차료 = getFormatInt64(value);
        }
        public String getstr도매_비용_지급임차료()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_지급임차료);
        }

        private Int64 도매_비용_운반비;
        public void set도매_비용_운반비(Int64 value)
        {
            도매_비용_운반비 = value;
        }
        public Int64 get도매_비용_운반비()
        {
            return 도매_비용_운반비;
        }
        public void set도매_비용_운반비(String value)
        {
            도매_비용_운반비 = getFormatInt64(value);
        }
        public String getstr도매_비용_운반비()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_운반비);
        }

        private Int64 도매_비용_차량유지비;
        public void set도매_비용_차량유지비(Int64 value)
        {
            도매_비용_차량유지비 = value;
        }
        public Int64 get도매_비용_차량유지비()
        {
            return 도매_비용_차량유지비;
        }
        public void set도매_비용_차량유지비(String value)
        {
            도매_비용_차량유지비 = getFormatInt64(value);
        }
        public String getstr도매_비용_차량유지비()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_차량유지비);
        }

        private Int64 도매_비용_지급수수료;
        public void set도매_비용_지급수수료(Int64 value)
        {
            도매_비용_지급수수료 = value;
        }
        public Int64 get도매_비용_지급수수료()
        {
            return 도매_비용_지급수수료;
        }
        public void set도매_비용_지급수수료(String value)
        {
            도매_비용_지급수수료 = getFormatInt64(value);
        }
        public String getstr도매_비용_지급수수료()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_지급수수료);
        }

        private Int64 도매_비용_판매촉진비;
        public void set도매_비용_판매촉진비(Int64 value)
        {
            도매_비용_판매촉진비 = value;
        }
        public Int64 get도매_비용_판매촉진비()
        {
            return 도매_비용_판매촉진비;
        }
        public void set도매_비용_판매촉진비(String value)
        {
            도매_비용_판매촉진비 = getFormatInt64(value);
        }
        public String getstr도매_비용_판매촉진비()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_판매촉진비);
        }

        private Int64 도매_비용_건물관리비;
        public void set도매_비용_건물관리비(Int64 value)
        {
            도매_비용_건물관리비 = value;
        }
        public Int64 get도매_비용_건물관리비()
        {
            return 도매_비용_건물관리비;
        }
        public void set도매_비용_건물관리비(String value)
        {
            도매_비용_건물관리비 = getFormatInt64(value);
        }
        public String getstr도매_비용_건물관리비()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_건물관리비);
        }


        private Int64 소매_수익_월평균업무취급수수료;
        public void set소매_수익_월평균업무취급수수료(Int64 value)
        {
            소매_수익_월평균업무취급수수료 = value;
        }
        public Int64 get소매_수익_월평균업무취급수수료()
        {
            return 소매_수익_월평균업무취급수수료;
        }
        public void set소매_수익_월평균업무취급수수료(String value)
        {
            소매_수익_월평균업무취급수수료 = getFormatInt64(value);
        }
        public String getstr소매_수익_월평균업무취급수수료()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_수익_월평균업무취급수수료);
        }

        private Int64 소매_수익_직영매장판매수익;
        public void set소매_수익_직영매장판매수익(Int64 value)
        {
            소매_수익_직영매장판매수익 = value;
        }
        public Int64 get소매_수익_직영매장판매수익()
        {
            return 소매_수익_직영매장판매수익;
        }
        public void set소매_수익_직영매장판매수익(String value)
        {
            소매_수익_직영매장판매수익 = getFormatInt64(value);
        }
        public String getstr소매_수익_직영매장판매수익()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_수익_직영매장판매수익);
        }


        private Int64 소매_비용_직원급여_간부급;
        public void set소매_비용_직원급여_간부급(Int64 value)
        {
            소매_비용_직원급여_간부급 = value;
        }
        public Int64 get소매_비용_직원급여_간부급()
        {
            return 소매_비용_직원급여_간부급;
        }
        public void set소매_비용_직원급여_간부급(String value)
        {
            소매_비용_직원급여_간부급 = getFormatInt64(value);
        }
        public String getstr소매_비용_직원급여_간부급()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_직원급여_간부급);
        }
        public Int64 get소매_비용_직원급여_간부급_총액(Int64 소매_간부급)
        {
            return CommonUtil.StringToIntVal(소매_비용_직원급여_간부급.ToString()) * 소매_간부급;
        }
        public String getstr소매_비용_직원급여_간부급_총액(Int64 소매_간부급)
        {//   ','가 적용된 값 리턴
            return getFormatString(get소매_비용_직원급여_간부급_총액(소매_간부급));
        }

        private Int64 소매_비용_직원급여_평사원;
        public void set소매_비용_직원급여_평사원(Int64 value)
        {
            소매_비용_직원급여_평사원 = value;
        }
        public Int64 get소매_비용_직원급여_평사원()
        {
            return 소매_비용_직원급여_평사원;
        }
        public void set소매_비용_직원급여_평사원(String value)
        {
            소매_비용_직원급여_평사원 = getFormatInt64(value);
        }
        public String getstr소매_비용_직원급여_평사원()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_직원급여_평사원);
        }
        public Int64 get소매_비용_직원급여_평사원_총액(Int64 소매_평사원)
        {
            return CommonUtil.StringToIntVal(소매_비용_직원급여_평사원.ToString()) * 소매_평사원;
        }
        public String getstr소매_비용_직원급여_평사원_총액(Int64 소매_평사원)
        {//   ','가 적용된 값 리턴
            return getFormatString(get소매_비용_직원급여_평사원_총액(소매_평사원));
        }


        private Int64 소매_비용_지급임차료;
        public void set소매_비용_지급임차료(Int64 value)
        {
            소매_비용_지급임차료 = value;
        }
        public Int64 get소매_비용_지급임차료()
        {
            return 소매_비용_지급임차료;
        }
        public void set소매_비용_지급임차료(String value)
        {
            소매_비용_지급임차료 = getFormatInt64(value);
        }
        public String getstr소매_비용_지급임차료()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_지급임차료);
        }

        private Int64 소매_비용_지급수수료;
        public void set소매_비용_지급수수료(Int64 value)
        {
            소매_비용_지급수수료 = value;
        }
        public Int64 get소매_비용_지급수수료()
        {
            return 소매_비용_지급수수료;
        }
        public void set소매_비용_지급수수료(String value)
        {
            소매_비용_지급수수료 = getFormatInt64(value);
        }
        public String getstr소매_비용_지급수수료()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_지급수수료);
        }

        private Int64 소매_비용_판매촉진비;
        public void set소매_비용_판매촉진비(Int64 value)
        {
            소매_비용_판매촉진비 = value;
        }
        public Int64 get소매_비용_판매촉진비()
        {
            return 소매_비용_판매촉진비;
        }
        public void set소매_비용_판매촉진비(String value)
        {
            소매_비용_판매촉진비 = getFormatInt64(value);
        }
        public String getstr소매_비용_판매촉진비()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_판매촉진비);
        }

        private Int64 소매_비용_건물관리비;
        public void set소매_비용_건물관리비(Int64 value)
        {
            소매_비용_건물관리비 = value;
        }
        public Int64 get소매_비용_건물관리비()
        {
            return 소매_비용_건물관리비;
        }
        public void set소매_비용_건물관리비(String value)
        {
            소매_비용_건물관리비 = getFormatInt64(value);
        }
        public String getstr소매_비용_건물관리비()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_건물관리비);
        }


        private Int64 도소매_비용_복리후생비;
        public void set도소매_비용_복리후생비(Int64 value)
        {
            도소매_비용_복리후생비 = value;
        }
        public Int64 get도소매_비용_복리후생비()
        {
            return 도소매_비용_복리후생비;
        }
        public void set도소매_비용_복리후생비(String value)
        {
            도소매_비용_복리후생비 = getFormatInt64(value);
        }
        public String getstr도소매_비용_복리후생비()
        {//   ','가 적용된 값 리턴
            return getFormatString(도소매_비용_복리후생비);
        }

        private Int64 도소매_비용_통신비;
        public void set도소매_비용_통신비(Int64 value)
        {
            도소매_비용_통신비 = value;
        }
        public Int64 get도소매_비용_통신비()
        {
            return 도소매_비용_통신비;
        }
        public void set도소매_비용_통신비(String value)
        {
            도소매_비용_통신비 = getFormatInt64(value);
        }
        public String getstr도소매_비용_통신비()
        {//   ','가 적용된 값 리턴
            return getFormatString(도소매_비용_통신비);
        }

        private Int64 도소매_비용_공과금;
        public void set도소매_비용_공과금(Int64 value)
        {
            도소매_비용_공과금 = value;
        }
        public Int64 get도소매_비용_공과금()
        {
            return 도소매_비용_공과금;
        }
        public void set도소매_비용_공과금(String value)
        {
            도소매_비용_공과금 = getFormatInt64(value);
        }
        public String getstr도소매_비용_공과금()
        {//   ','가 적용된 값 리턴
            return getFormatString(도소매_비용_공과금);
        }

        private Int64 도소매_비용_소모품비;
        public void set도소매_비용_소모품비(Int64 value)
        {
            도소매_비용_소모품비 = value;
        }
        public Int64 get도소매_비용_소모품비()
        {
            return 도소매_비용_소모품비;
        }
        public void set도소매_비용_소모품비(String value)
        {
            도소매_비용_소모품비 = getFormatInt64(value);
        }
        public String getstr도소매_비용_소모품비()
        {//   ','가 적용된 값 리턴
            return getFormatString(도소매_비용_소모품비);
        }

        private Int64 도소매_비용_이자비용;
        public void set도소매_비용_이자비용(Int64 value)
        {
            도소매_비용_이자비용 = value;
        }
        public Int64 get도소매_비용_이자비용()
        {
            return 도소매_비용_이자비용;
        }
        public void set도소매_비용_이자비용(String value)
        {
            도소매_비용_이자비용 = getFormatInt64(value);
        }
        public String getstr도소매_비용_이자비용()
        {//   ','가 적용된 값 리턴
            return getFormatString(도소매_비용_이자비용);
        }

        private Int64 도소매_비용_부가세;
        public void set도소매_비용_부가세(Int64 value)
        {
            도소매_비용_부가세 = value;
        }
        public Int64 get도소매_비용_부가세()
        {
            return 도소매_비용_부가세;
        }
        public void set도소매_비용_부가세(String value)
        {
            도소매_비용_부가세 = getFormatInt64(value);
        }
        public String getstr도소매_비용_부가세()
        {//   ','가 적용된 값 리턴
            return getFormatString(도소매_비용_부가세);
        }

        private Int64 도소매_비용_법인세;
        public void set도소매_비용_법인세(Int64 value)
        {
            도소매_비용_법인세 = value;
        }
        public Int64 get도소매_비용_법인세()
        {
            return 도소매_비용_법인세;
        }
        public void set도소매_비용_법인세(String value)
        {
            도소매_비용_법인세 = getFormatInt64(value);
        }
        public String getstr도소매_비용_법인세()
        {//   ','가 적용된 값 리턴
            return getFormatString(도소매_비용_법인세);
        }

        private Int64 도소매_비용_기타;
        public void set도소매_비용_기타(Int64 value)
        {
            도소매_비용_기타 = value;
        }
        public Int64 get도소매_비용_기타()
        {
            return 도소매_비용_기타;
        }
        public void set도소매_비용_기타(String value)
        {
            도소매_비용_기타 = getFormatInt64(value);
        }
        public String getstr도소매_비용_기타()
        {//   ','가 적용된 값 리턴
            return getFormatString(도소매_비용_기타);
        }


        //배열 IO 
        public void setArrData(Int64[] arrvalue)
        {
            도매_수익_월평균관리수수료 = arrvalue[0];
            도매_수익_CS관리수수료 = arrvalue[1];
            도매_수익_사업자모델매입관련추가수익 = arrvalue[2];
            도매_수익_유통모델매입관련추가수익_현금DC = arrvalue[3];
            도매_수익_유통모델매입관련추가수익_VolumeDC = arrvalue[4];

            도매_비용_대리점투자금액_신규 = arrvalue[5];
            도매_비용_대리점투자금액_기변 = arrvalue[6];
            도매_비용_직원급여_간부급 = arrvalue[7];
            도매_비용_직원급여_평사원 = arrvalue[8];
            도매_비용_지급임차료 = arrvalue[9];
            도매_비용_운반비 = arrvalue[10];
            도매_비용_차량유지비 = arrvalue[11];
            도매_비용_지급수수료 = arrvalue[12];
            도매_비용_판매촉진비 = arrvalue[13];
            도매_비용_건물관리비 = arrvalue[14];

            소매_수익_월평균업무취급수수료 = arrvalue[15];
            소매_수익_직영매장판매수익 = arrvalue[16];

            소매_비용_직원급여_간부급 = arrvalue[17];
            소매_비용_직원급여_평사원 = arrvalue[18];
            소매_비용_지급임차료 = arrvalue[19];
            소매_비용_지급수수료 = arrvalue[20];
            소매_비용_판매촉진비 = arrvalue[21];
            소매_비용_건물관리비 = arrvalue[22];

            도소매_비용_복리후생비 = arrvalue[23];
            도소매_비용_통신비 = arrvalue[24];
            도소매_비용_공과금 = arrvalue[25];
            도소매_비용_소모품비 = arrvalue[26];
            도소매_비용_이자비용 = arrvalue[27];
            도소매_비용_부가세 = arrvalue[28];
            도소매_비용_법인세 = arrvalue[29];
            도소매_비용_기타 = arrvalue[30];
        }

        public Int64[] getArrData()
        {
            Int64[] arrvalue = new Int64[31];

            arrvalue[0] = 도매_수익_월평균관리수수료;
            arrvalue[1] = 도매_수익_CS관리수수료;
            arrvalue[2] = 도매_수익_사업자모델매입관련추가수익;
            arrvalue[3] = 도매_수익_유통모델매입관련추가수익_현금DC;
            arrvalue[4] = 도매_수익_유통모델매입관련추가수익_VolumeDC;

            arrvalue[5] = 도매_비용_대리점투자금액_신규;
            arrvalue[6] = 도매_비용_대리점투자금액_기변;
            arrvalue[7] = 도매_비용_직원급여_간부급;
            arrvalue[8] = 도매_비용_직원급여_평사원;
            arrvalue[9] = 도매_비용_지급임차료;
            arrvalue[10] = 도매_비용_운반비;
            arrvalue[11] = 도매_비용_차량유지비;
            arrvalue[12] = 도매_비용_지급수수료;
            arrvalue[13] = 도매_비용_판매촉진비;
            arrvalue[14] = 도매_비용_건물관리비;

            arrvalue[15] = 소매_수익_월평균업무취급수수료;
            arrvalue[16] = 소매_수익_직영매장판매수익;
            arrvalue[17] = 소매_비용_직원급여_간부급;
            arrvalue[18] = 소매_비용_직원급여_평사원;
            arrvalue[19] = 소매_비용_지급임차료;
            arrvalue[20] = 소매_비용_지급수수료;
            arrvalue[21] = 소매_비용_판매촉진비;
            arrvalue[22] = 소매_비용_건물관리비;

            arrvalue[23] = 도소매_비용_복리후생비;
            arrvalue[24] = 도소매_비용_통신비;
            arrvalue[25] = 도소매_비용_공과금;
            arrvalue[26] = 도소매_비용_소모품비;
            arrvalue[27] = 도소매_비용_이자비용;
            arrvalue[28] = 도소매_비용_부가세;
            arrvalue[29] = 도소매_비용_법인세;
            arrvalue[30] = 도소매_비용_기타;

            return arrvalue;
        }

        public string[] getArrData_BusinessAvg()
        {
            string[] arrvalue = new string[31];

            arrvalue[0] = 도매_수익_월평균관리수수료.ToString();
            arrvalue[1] = 도매_수익_CS관리수수료.ToString();
            arrvalue[2] = 도매_수익_사업자모델매입관련추가수익.ToString();
            arrvalue[3] = 도매_수익_유통모델매입관련추가수익_현금DC.ToString();
            arrvalue[4] = 도매_수익_유통모델매입관련추가수익_VolumeDC.ToString();

            arrvalue[5] = 도매_비용_대리점투자금액_신규.ToString();
            arrvalue[6] = 도매_비용_대리점투자금액_기변.ToString();
            arrvalue[7] = 도매_비용_직원급여_간부급.ToString();
            arrvalue[8] = 도매_비용_직원급여_평사원.ToString();
            arrvalue[9] = 도매_비용_지급임차료.ToString();
            arrvalue[10] = 도매_비용_운반비.ToString();
            arrvalue[11] = 도매_비용_차량유지비.ToString();
            arrvalue[12] = 도매_비용_지급수수료.ToString();
            arrvalue[13] = 도매_비용_판매촉진비.ToString();
            arrvalue[14] = 도매_비용_건물관리비.ToString();

            arrvalue[15] = 소매_수익_월평균업무취급수수료.ToString();
            arrvalue[16] = 소매_수익_직영매장판매수익.ToString();
            arrvalue[17] = 소매_비용_직원급여_간부급.ToString();
            arrvalue[18] = 소매_비용_직원급여_평사원.ToString();
            arrvalue[19] = 소매_비용_지급임차료.ToString();
            arrvalue[20] = 소매_비용_지급수수료.ToString();
            arrvalue[21] = 소매_비용_판매촉진비.ToString();
            arrvalue[22] = 소매_비용_건물관리비.ToString();

            arrvalue[23] = 도소매_비용_복리후생비.ToString();
            arrvalue[24] = 도소매_비용_통신비.ToString();
            arrvalue[25] = 도소매_비용_공과금.ToString();
            arrvalue[26] = 도소매_비용_소모품비.ToString();
            arrvalue[27] = 도소매_비용_이자비용.ToString();
            arrvalue[28] = 도소매_비용_부가세.ToString();
            arrvalue[29] = 도소매_비용_법인세.ToString();
            arrvalue[30] = 도소매_비용_기타.ToString();

            return arrvalue;
        }
        public void setArrData_DetailInput(string[] arrvalue)
        {
            도매_수익_월평균관리수수료 = CommonUtil.StringToIntVal( arrvalue[0]);
            도매_수익_CS관리수수료 = CommonUtil.StringToIntVal( arrvalue[1]);
            도매_수익_사업자모델매입관련추가수익 = CommonUtil.StringToIntVal( arrvalue[2]);
            도매_수익_유통모델매입관련추가수익_현금DC = CommonUtil.StringToIntVal( arrvalue[3]);
            도매_수익_유통모델매입관련추가수익_VolumeDC = CommonUtil.StringToIntVal( arrvalue[4]);

            도매_비용_대리점투자금액_신규 = CommonUtil.StringToIntVal( arrvalue[5]);
            도매_비용_대리점투자금액_기변 = CommonUtil.StringToIntVal( arrvalue[6]);
            도매_비용_직원급여_간부급 = CommonUtil.StringToIntVal( arrvalue[7]);
            도매_비용_직원급여_평사원 = CommonUtil.StringToIntVal( arrvalue[8]);
            도매_비용_지급임차료 = CommonUtil.StringToIntVal( arrvalue[9]);
            도매_비용_운반비 = CommonUtil.StringToIntVal( arrvalue[10]);
            도매_비용_차량유지비 = CommonUtil.StringToIntVal( arrvalue[11]);
            도매_비용_지급수수료 = CommonUtil.StringToIntVal( arrvalue[12]);
            도매_비용_판매촉진비 = CommonUtil.StringToIntVal( arrvalue[13]);
            도매_비용_건물관리비 = CommonUtil.StringToIntVal( arrvalue[14]);

            소매_수익_월평균업무취급수수료 = CommonUtil.StringToIntVal( arrvalue[15]);
            소매_수익_직영매장판매수익 = CommonUtil.StringToIntVal( arrvalue[16]);

            소매_비용_직원급여_간부급 = CommonUtil.StringToIntVal( arrvalue[17]);
            소매_비용_직원급여_평사원 = CommonUtil.StringToIntVal( arrvalue[18]);
            소매_비용_지급임차료 = CommonUtil.StringToIntVal( arrvalue[19]);
            소매_비용_지급수수료 = CommonUtil.StringToIntVal( arrvalue[20]);
            소매_비용_판매촉진비 = CommonUtil.StringToIntVal( arrvalue[21]);
            소매_비용_건물관리비 = CommonUtil.StringToIntVal( arrvalue[22]);

            도소매_비용_복리후생비 = CommonUtil.StringToIntVal( arrvalue[23]);
            도소매_비용_통신비 = CommonUtil.StringToIntVal( arrvalue[24]);
            도소매_비용_공과금 = CommonUtil.StringToIntVal( arrvalue[25]);
            도소매_비용_소모품비 = CommonUtil.StringToIntVal( arrvalue[26]);
            도소매_비용_이자비용 = CommonUtil.StringToIntVal( arrvalue[27]);
            도소매_비용_부가세 = CommonUtil.StringToIntVal( arrvalue[28]);
            도소매_비용_법인세 = CommonUtil.StringToIntVal( arrvalue[29]);
            도소매_비용_기타 = CommonUtil.StringToIntVal( arrvalue[30]);
        }


        public Int64[] getArrData_DetailInput(Int64 도매_간부수, Int64 도매_평사원수, Int64 소매_간부수, Int64 수매_평사원수)
        {
            Int64[] arrvalue = new Int64[36];

            int i = 0;
            arrvalue[i] = 도매_수익_월평균관리수수료;
            arrvalue[i++] = 도매_수익_CS관리수수료;
            arrvalue[i++] = get도매_수익_CS관리수수료_분기();
            arrvalue[i++] = 도매_수익_사업자모델매입관련추가수익;
            arrvalue[i++] = 도매_수익_유통모델매입관련추가수익_현금DC;
            arrvalue[i++] = 도매_수익_유통모델매입관련추가수익_VolumeDC;

            arrvalue[i++] = 도매_비용_대리점투자금액_신규;
            arrvalue[i++] = 도매_비용_대리점투자금액_기변;
            arrvalue[i++] = get도매_비용_직원급여_간부급_총액(도매_간부수);
            arrvalue[i++] = get도매_비용_직원급여_평사원_총액(도매_평사원수);
            arrvalue[i++] = 도매_비용_직원급여_간부급;
            arrvalue[i++] = 도매_비용_직원급여_평사원;

            arrvalue[i++] = 도매_비용_지급임차료;
            arrvalue[i++] = 도매_비용_운반비;
            arrvalue[i++] = 도매_비용_차량유지비;
            arrvalue[i++] = 도매_비용_지급수수료;
            arrvalue[i++] = 도매_비용_판매촉진비;
            arrvalue[i++] = 도매_비용_건물관리비;

            arrvalue[i++] = 소매_수익_월평균업무취급수수료;
            arrvalue[i++] = 소매_수익_직영매장판매수익;
            arrvalue[i++] = get소매_비용_직원급여_간부급_총액(소매_간부수);
            arrvalue[i++] = get소매_비용_직원급여_평사원_총액(수매_평사원수);
            arrvalue[i++] = 소매_비용_직원급여_간부급;
            arrvalue[i++] = 소매_비용_직원급여_평사원;

            arrvalue[i++] = 소매_비용_지급임차료;
            arrvalue[i++] = 소매_비용_지급수수료;
            arrvalue[i++] = 소매_비용_판매촉진비;
            arrvalue[i++] = 소매_비용_건물관리비;

            arrvalue[i++] = 도소매_비용_복리후생비;
            arrvalue[i++] = 도소매_비용_통신비;
            arrvalue[i++] = 도소매_비용_공과금;
            arrvalue[i++] = 도소매_비용_소모품비;
            arrvalue[i++] = 도소매_비용_이자비용;
            arrvalue[i++] = 도소매_비용_부가세;
            arrvalue[i++] = 도소매_비용_법인세;
            arrvalue[i++] = 도소매_비용_기타;

            return arrvalue;
        }

        private String getFormatString(Int64 value)
        {
            CultureInfo cur = new CultureInfo(CultureInfo.InvariantCulture.LCID);
            cur.NumberFormat.NumberDecimalDigits = 0;
            return value.ToString("N", cur);
        }

        private Int64 getFormatInt64(String value)
        {             
            String temp = "0";
            temp = value.Replace(",", "");
            return CommonUtil.StringToIntVal(temp);
        }



    }
}
