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

        private Double 도매_수익_월평균관리수수료;
        public void set도매_수익_월평균관리수수료(Double value)
        {
            도매_수익_월평균관리수수료 = value;
        }
        public Double get도매_수익_월평균관리수수료()
        {
            return 도매_수익_월평균관리수수료;
        }
        public void set도매_수익_월평균관리수수료(String value)
        {
            도매_수익_월평균관리수수료 = getFormatDouble(value);
        }
        public String getstr도매_수익_월평균관리수수료()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_수익_월평균관리수수료);
        }


        private Double 도매_수익_CS관리수수료;
        public void set도매_수익_CS관리수수료(Double value)
        {
            도매_수익_CS관리수수료 = value;
        }
        public Double get도매_수익_CS관리수수료()
        {
            return 도매_수익_CS관리수수료;
        }
        public void set도매_수익_CS관리수수료(String value)
        {
            도매_수익_CS관리수수료 = getFormatDouble(value);
        }
        public String getstr도매_수익_CS관리수수료()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_수익_CS관리수수료);
        }
        public Double get도매_수익_CS관리수수료_분기()
        {
            return 도매_수익_CS관리수수료*CommonUtil.QUARTER;
        }
        public String getstr도매_수익_CS관리수수료_분기()
        {//   ','가 적용된 값 리턴
            return getFormatString(get도매_수익_CS관리수수료_분기());
        }


        private Double 도매_수익_사업자모델매입관련추가수익;
        public void set도매_수익_사업자모델매입관련추가수익(Double value)
        {
            도매_수익_사업자모델매입관련추가수익 = value;
        }
        public Double get도매_수익_사업자모델매입관련추가수익()
        {
            return 도매_수익_사업자모델매입관련추가수익;
        }
        public void set도매_수익_사업자모델매입관련추가수익(String value)
        {
            도매_수익_사업자모델매입관련추가수익 = getFormatDouble(value);
        }
        public String getstr도매_수익_사업자모델매입관련추가수익()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_수익_사업자모델매입관련추가수익);
        }

        private Double 도매_수익_유통모델매입관련추가수익_현금DC;
        public void set도매_수익_유통모델매입관련추가수익_현금DC(Double value)
        {
            도매_수익_유통모델매입관련추가수익_현금DC = value;
        }
        public Double get도매_수익_유통모델매입관련추가수익_현금DC()
        {
            return 도매_수익_유통모델매입관련추가수익_현금DC;
        }
        public void set도매_수익_유통모델매입관련추가수익_현금DC(String value)
        {
            도매_수익_유통모델매입관련추가수익_현금DC = getFormatDouble(value);
        }
        public String getstr도매_수익_유통모델매입관련추가수익_현금DC()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_수익_유통모델매입관련추가수익_현금DC);
        }

        private Double 도매_수익_유통모델매입관련추가수익_VolumeDC;
        public void set도매_수익_유통모델매입관련추가수익_VolumeDC(Double value)
        {
            도매_수익_유통모델매입관련추가수익_VolumeDC = value;
        }
        public Double get도매_수익_유통모델매입관련추가수익_VolumeDC()
        {
            return 도매_수익_유통모델매입관련추가수익_VolumeDC;
        }
        public void set도매_수익_유통모델매입관련추가수익_VolumeDC(String value)
        {
            도매_수익_유통모델매입관련추가수익_VolumeDC = getFormatDouble(value);
        }
        public String getstr도매_수익_유통모델매입관련추가수익_VolumeDC()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_수익_유통모델매입관련추가수익_VolumeDC);
        }


        private Double 도매_비용_대리점투자금액_신규;
        public void set도매_비용_대리점투자금액_신규(Double value)
        {
            도매_비용_대리점투자금액_신규 = value;
        }
        public Double get도매_비용_대리점투자금액_신규()
        {
            return 도매_비용_대리점투자금액_신규;
        }
        public void set도매_비용_대리점투자금액_신규(String value)
        {
            도매_비용_대리점투자금액_신규 = getFormatDouble(value);
        }
        public String getstr도매_비용_대리점투자금액_신규()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_대리점투자금액_신규);
        }

        private Double 도매_비용_대리점투자금액_기변;
        public void set도매_비용_대리점투자금액_기변(Double value)
        {
            도매_비용_대리점투자금액_기변 = value;
        }
        public Double get도매_비용_대리점투자금액_기변()
        {
            return 도매_비용_대리점투자금액_기변;
        }
        public void set도매_비용_대리점투자금액_기변(String value)
        {
            도매_비용_대리점투자금액_기변 = getFormatDouble(value);
        }
        public String getstr도매_비용_대리점투자금액_기변()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_대리점투자금액_기변);
        }

        private Double 도매_비용_직원급여_간부급;
        public void set도매_비용_직원급여_간부급(Double value)
        {
            도매_비용_직원급여_간부급 = value;
        }
        public Double get도매_비용_직원급여_간부급()
        {
            return 도매_비용_직원급여_간부급;
        }
        public void set도매_비용_직원급여_간부급(String value)
        {
            도매_비용_직원급여_간부급 = getFormatDouble(value);
        }
        public String getstr도매_비용_직원급여_간부급()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_직원급여_간부급);
        }
        public Double get도매_비용_직원급여_간부급_총액(Double 도매_간부수)
        {
            return CommonUtil.StringToDoubleVal(도매_비용_직원급여_간부급.ToString()) * 도매_간부수;
        }
        public String getstr도매_비용_직원급여_간부급_총액(Double 도매_간부수)
        {//   ','가 적용된 값 리턴
            return getFormatString(get도매_비용_직원급여_간부급_총액(도매_간부수));
        }

        private Double 도매_비용_직원급여_평사원;
        public void set도매_비용_직원급여_평사원(Double value)
        {
            도매_비용_직원급여_평사원 = value;
        }
        public Double get도매_비용_직원급여_평사원()
        {
            return 도매_비용_직원급여_평사원;
        }
        public void set도매_비용_직원급여_평사원(String value)
        {
            도매_비용_직원급여_평사원 = getFormatDouble(value);
        }
        public String getstr도매_비용_직원급여_평사원()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_직원급여_평사원);
        }
        public Double get도매_비용_직원급여_평사원_총액(Double 도매_평사원)
        {
            return CommonUtil.StringToDoubleVal(도매_비용_직원급여_평사원.ToString()) * 도매_평사원;
        }
        public String getstr도매_비용_직원급여_평사원_총액(Double 도매_평사원)
        {//   ','가 적용된 값 리턴
            return getFormatString(get도매_비용_직원급여_평사원_총액(도매_평사원));
        }


        private Double 도매_비용_지급임차료;
        public void set도매_비용_지급임차료(Double value)
        {
            도매_비용_지급임차료 = value;
        }
        public Double get도매_비용_지급임차료()
        {
            return 도매_비용_지급임차료;
        }
        public void set도매_비용_지급임차료(String value)
        {
            도매_비용_지급임차료 = getFormatDouble(value);
        }
        public String getstr도매_비용_지급임차료()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_지급임차료);
        }

        private Double 도매_비용_운반비;
        public void set도매_비용_운반비(Double value)
        {
            도매_비용_운반비 = value;
        }
        public Double get도매_비용_운반비()
        {
            return 도매_비용_운반비;
        }
        public void set도매_비용_운반비(String value)
        {
            도매_비용_운반비 = getFormatDouble(value);
        }
        public String getstr도매_비용_운반비()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_운반비);
        }

        private Double 도매_비용_차량유지비;
        public void set도매_비용_차량유지비(Double value)
        {
            도매_비용_차량유지비 = value;
        }
        public Double get도매_비용_차량유지비()
        {
            return 도매_비용_차량유지비;
        }
        public void set도매_비용_차량유지비(String value)
        {
            도매_비용_차량유지비 = getFormatDouble(value);
        }
        public String getstr도매_비용_차량유지비()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_차량유지비);
        }

        private Double 도매_비용_지급수수료;
        public void set도매_비용_지급수수료(Double value)
        {
            도매_비용_지급수수료 = value;
        }
        public Double get도매_비용_지급수수료()
        {
            return 도매_비용_지급수수료;
        }
        public void set도매_비용_지급수수료(String value)
        {
            도매_비용_지급수수료 = getFormatDouble(value);
        }
        public String getstr도매_비용_지급수수료()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_지급수수료);
        }

        private Double 도매_비용_판매촉진비;
        public void set도매_비용_판매촉진비(Double value)
        {
            도매_비용_판매촉진비 = value;
        }
        public Double get도매_비용_판매촉진비()
        {
            return 도매_비용_판매촉진비;
        }
        public void set도매_비용_판매촉진비(String value)
        {
            도매_비용_판매촉진비 = getFormatDouble(value);
        }
        public String getstr도매_비용_판매촉진비()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_판매촉진비);
        }

        private Double 도매_비용_건물관리비;
        public void set도매_비용_건물관리비(Double value)
        {
            도매_비용_건물관리비 = value;
        }
        public Double get도매_비용_건물관리비()
        {
            return 도매_비용_건물관리비;
        }
        public void set도매_비용_건물관리비(String value)
        {
            도매_비용_건물관리비 = getFormatDouble(value);
        }
        public String getstr도매_비용_건물관리비()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_건물관리비);
        }


        private Double 소매_수익_월평균업무취급수수료;
        public void set소매_수익_월평균업무취급수수료(Double value)
        {
            소매_수익_월평균업무취급수수료 = value;
        }
        public Double get소매_수익_월평균업무취급수수료()
        {
            return 소매_수익_월평균업무취급수수료;
        }
        public void set소매_수익_월평균업무취급수수료(String value)
        {
            소매_수익_월평균업무취급수수료 = getFormatDouble(value);
        }
        public String getstr소매_수익_월평균업무취급수수료()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_수익_월평균업무취급수수료);
        }

        private Double 소매_수익_직영매장판매수익;
        public void set소매_수익_직영매장판매수익(Double value)
        {
            소매_수익_직영매장판매수익 = value;
        }
        public Double get소매_수익_직영매장판매수익()
        {
            return 소매_수익_직영매장판매수익;
        }
        public void set소매_수익_직영매장판매수익(String value)
        {
            소매_수익_직영매장판매수익 = getFormatDouble(value);
        }
        public String getstr소매_수익_직영매장판매수익()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_수익_직영매장판매수익);
        }


        private Double 소매_비용_직원급여_간부급;
        public void set소매_비용_직원급여_간부급(Double value)
        {
            소매_비용_직원급여_간부급 = value;
        }
        public Double get소매_비용_직원급여_간부급()
        {
            return 소매_비용_직원급여_간부급;
        }
        public void set소매_비용_직원급여_간부급(String value)
        {
            소매_비용_직원급여_간부급 = getFormatDouble(value);
        }
        public String getstr소매_비용_직원급여_간부급()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_직원급여_간부급);
        }
        public Double get소매_비용_직원급여_간부급_총액(Double 소매_간부급)
        {
            return CommonUtil.StringToDoubleVal(소매_비용_직원급여_간부급.ToString()) * 소매_간부급;
        }
        public String getstr소매_비용_직원급여_간부급_총액(Double 소매_간부급)
        {//   ','가 적용된 값 리턴
            return getFormatString(get소매_비용_직원급여_간부급_총액(소매_간부급));
        }

        private Double 소매_비용_직원급여_평사원;
        public void set소매_비용_직원급여_평사원(Double value)
        {
            소매_비용_직원급여_평사원 = value;
        }
        public Double get소매_비용_직원급여_평사원()
        {
            return 소매_비용_직원급여_평사원;
        }
        public void set소매_비용_직원급여_평사원(String value)
        {
            소매_비용_직원급여_평사원 = getFormatDouble(value);
        }
        public String getstr소매_비용_직원급여_평사원()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_직원급여_평사원);
        }
        public Double get소매_비용_직원급여_평사원_총액(Double 소매_평사원)
        {
            return CommonUtil.StringToDoubleVal(소매_비용_직원급여_평사원.ToString()) * 소매_평사원;
        }
        public String getstr소매_비용_직원급여_평사원_총액(Double 소매_평사원)
        {//   ','가 적용된 값 리턴
            return getFormatString(get소매_비용_직원급여_평사원_총액(소매_평사원));
        }


        private Double 소매_비용_지급임차료;
        public void set소매_비용_지급임차료(Double value)
        {
            소매_비용_지급임차료 = value;
        }
        public Double get소매_비용_지급임차료()
        {
            return 소매_비용_지급임차료;
        }
        public void set소매_비용_지급임차료(String value)
        {
            소매_비용_지급임차료 = getFormatDouble(value);
        }
        public String getstr소매_비용_지급임차료()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_지급임차료);
        }

        private Double 소매_비용_지급수수료;
        public void set소매_비용_지급수수료(Double value)
        {
            소매_비용_지급수수료 = value;
        }
        public Double get소매_비용_지급수수료()
        {
            return 소매_비용_지급수수료;
        }
        public void set소매_비용_지급수수료(String value)
        {
            소매_비용_지급수수료 = getFormatDouble(value);
        }
        public String getstr소매_비용_지급수수료()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_지급수수료);
        }

        private Double 소매_비용_판매촉진비;
        public void set소매_비용_판매촉진비(Double value)
        {
            소매_비용_판매촉진비 = value;
        }
        public Double get소매_비용_판매촉진비()
        {
            return 소매_비용_판매촉진비;
        }
        public void set소매_비용_판매촉진비(String value)
        {
            소매_비용_판매촉진비 = getFormatDouble(value);
        }
        public String getstr소매_비용_판매촉진비()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_판매촉진비);
        }

        private Double 소매_비용_건물관리비;
        public void set소매_비용_건물관리비(Double value)
        {
            소매_비용_건물관리비 = value;
        }
        public Double get소매_비용_건물관리비()
        {
            return 소매_비용_건물관리비;
        }
        public void set소매_비용_건물관리비(String value)
        {
            소매_비용_건물관리비 = getFormatDouble(value);
        }
        public String getstr소매_비용_건물관리비()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_건물관리비);
        }


        private Double 도소매_비용_복리후생비;
        public void set도소매_비용_복리후생비(Double value)
        {
            도소매_비용_복리후생비 = value;
        }
        public Double get도소매_비용_복리후생비()
        {
            return 도소매_비용_복리후생비;
        }
        public void set도소매_비용_복리후생비(String value)
        {
            도소매_비용_복리후생비 = getFormatDouble(value);
        }
        public String getstr도소매_비용_복리후생비()
        {//   ','가 적용된 값 리턴
            return getFormatString(도소매_비용_복리후생비);
        }

        private Double 도소매_비용_통신비;
        public void set도소매_비용_통신비(Double value)
        {
            도소매_비용_통신비 = value;
        }
        public Double get도소매_비용_통신비()
        {
            return 도소매_비용_통신비;
        }
        public void set도소매_비용_통신비(String value)
        {
            도소매_비용_통신비 = getFormatDouble(value);
        }
        public String getstr도소매_비용_통신비()
        {//   ','가 적용된 값 리턴
            return getFormatString(도소매_비용_통신비);
        }

        private Double 도소매_비용_공과금;
        public void set도소매_비용_공과금(Double value)
        {
            도소매_비용_공과금 = value;
        }
        public Double get도소매_비용_공과금()
        {
            return 도소매_비용_공과금;
        }
        public void set도소매_비용_공과금(String value)
        {
            도소매_비용_공과금 = getFormatDouble(value);
        }
        public String getstr도소매_비용_공과금()
        {//   ','가 적용된 값 리턴
            return getFormatString(도소매_비용_공과금);
        }

        private Double 도소매_비용_소모품비;
        public void set도소매_비용_소모품비(Double value)
        {
            도소매_비용_소모품비 = value;
        }
        public Double get도소매_비용_소모품비()
        {
            return 도소매_비용_소모품비;
        }
        public void set도소매_비용_소모품비(String value)
        {
            도소매_비용_소모품비 = getFormatDouble(value);
        }
        public String getstr도소매_비용_소모품비()
        {//   ','가 적용된 값 리턴
            return getFormatString(도소매_비용_소모품비);
        }

        private Double 도소매_비용_이자비용;
        public void set도소매_비용_이자비용(Double value)
        {
            도소매_비용_이자비용 = value;
        }
        public Double get도소매_비용_이자비용()
        {
            return 도소매_비용_이자비용;
        }
        public void set도소매_비용_이자비용(String value)
        {
            도소매_비용_이자비용 = getFormatDouble(value);
        }
        public String getstr도소매_비용_이자비용()
        {//   ','가 적용된 값 리턴
            return getFormatString(도소매_비용_이자비용);
        }

        private Double 도소매_비용_부가세;
        public void set도소매_비용_부가세(Double value)
        {
            도소매_비용_부가세 = value;
        }
        public Double get도소매_비용_부가세()
        {
            return 도소매_비용_부가세;
        }
        public void set도소매_비용_부가세(String value)
        {
            도소매_비용_부가세 = getFormatDouble(value);
        }
        public String getstr도소매_비용_부가세()
        {//   ','가 적용된 값 리턴
            return getFormatString(도소매_비용_부가세);
        }

        private Double 도소매_비용_법인세;
        public void set도소매_비용_법인세(Double value)
        {
            도소매_비용_법인세 = value;
        }
        public Double get도소매_비용_법인세()
        {
            return 도소매_비용_법인세;
        }
        public void set도소매_비용_법인세(String value)
        {
            도소매_비용_법인세 = getFormatDouble(value);
        }
        public String getstr도소매_비용_법인세()
        {//   ','가 적용된 값 리턴
            return getFormatString(도소매_비용_법인세);
        }

        private Double 도소매_비용_기타;
        public void set도소매_비용_기타(Double value)
        {
            도소매_비용_기타 = value;
        }
        public Double get도소매_비용_기타()
        {
            return 도소매_비용_기타;
        }
        public void set도소매_비용_기타(String value)
        {
            도소매_비용_기타 = getFormatDouble(value);
        }
        public String getstr도소매_비용_기타()
        {//   ','가 적용된 값 리턴
            return getFormatString(도소매_비용_기타);
        }


        //배열 IO 
        public void setArrData(String[] arrvalue)
        {
            Double[] arrDouble = new Double[arrvalue.Length];
            for (int i = 0; i < arrvalue.Length; i++)
            {
                arrDouble[i] = Convert.ToDouble(arrvalue[i]);
            }
            setArrData(arrDouble);
        }
        public void setArrData(Double[] arrvalue)
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

        public Double[] getArrData()
        {
            Double[] arrvalue = new Double[31];

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
            도매_수익_월평균관리수수료 = CommonUtil.StringToDoubleVal( arrvalue[0]);
            도매_수익_CS관리수수료 = CommonUtil.StringToDoubleVal( arrvalue[1]);
            도매_수익_사업자모델매입관련추가수익 = CommonUtil.StringToDoubleVal( arrvalue[2]);
            도매_수익_유통모델매입관련추가수익_현금DC = CommonUtil.StringToDoubleVal( arrvalue[3]);
            도매_수익_유통모델매입관련추가수익_VolumeDC = CommonUtil.StringToDoubleVal( arrvalue[4]);

            도매_비용_대리점투자금액_신규 = CommonUtil.StringToDoubleVal( arrvalue[5]);
            도매_비용_대리점투자금액_기변 = CommonUtil.StringToDoubleVal( arrvalue[6]);
            도매_비용_직원급여_간부급 = CommonUtil.StringToDoubleVal( arrvalue[7]);
            도매_비용_직원급여_평사원 = CommonUtil.StringToDoubleVal( arrvalue[8]);
            도매_비용_지급임차료 = CommonUtil.StringToDoubleVal( arrvalue[9]);
            도매_비용_운반비 = CommonUtil.StringToDoubleVal( arrvalue[10]);
            도매_비용_차량유지비 = CommonUtil.StringToDoubleVal( arrvalue[11]);
            도매_비용_지급수수료 = CommonUtil.StringToDoubleVal( arrvalue[12]);
            도매_비용_판매촉진비 = CommonUtil.StringToDoubleVal( arrvalue[13]);
            도매_비용_건물관리비 = CommonUtil.StringToDoubleVal( arrvalue[14]);

            소매_수익_월평균업무취급수수료 = CommonUtil.StringToDoubleVal( arrvalue[15]);
            소매_수익_직영매장판매수익 = CommonUtil.StringToDoubleVal( arrvalue[16]);

            소매_비용_직원급여_간부급 = CommonUtil.StringToDoubleVal( arrvalue[17]);
            소매_비용_직원급여_평사원 = CommonUtil.StringToDoubleVal( arrvalue[18]);
            소매_비용_지급임차료 = CommonUtil.StringToDoubleVal( arrvalue[19]);
            소매_비용_지급수수료 = CommonUtil.StringToDoubleVal( arrvalue[20]);
            소매_비용_판매촉진비 = CommonUtil.StringToDoubleVal( arrvalue[21]);
            소매_비용_건물관리비 = CommonUtil.StringToDoubleVal( arrvalue[22]);

            도소매_비용_복리후생비 = CommonUtil.StringToDoubleVal( arrvalue[23]);
            도소매_비용_통신비 = CommonUtil.StringToDoubleVal( arrvalue[24]);
            도소매_비용_공과금 = CommonUtil.StringToDoubleVal( arrvalue[25]);
            도소매_비용_소모품비 = CommonUtil.StringToDoubleVal( arrvalue[26]);
            도소매_비용_이자비용 = CommonUtil.StringToDoubleVal( arrvalue[27]);
            도소매_비용_부가세 = CommonUtil.StringToDoubleVal( arrvalue[28]);
            도소매_비용_법인세 = CommonUtil.StringToDoubleVal( arrvalue[29]);
            도소매_비용_기타 = CommonUtil.StringToDoubleVal( arrvalue[30]);
        }


        public Double[] getArrData_DetailInput(Double 도매_간부수, Double 도매_평사원수, Double 소매_간부수, Double 소매_평사원수)
        {
            Double[] arrvalue = new Double[36];

            int i = 0;
            arrvalue[i++] = 도매_수익_월평균관리수수료;
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
            arrvalue[i++] = get소매_비용_직원급여_평사원_총액(소매_평사원수);
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

        // 관리자 데이터 세팅
        private Double _ASP_유통_LG;
        public Double ASP_유통_LG
        {
            get { return _ASP_유통_LG; }
            set { _ASP_유통_LG = value; }
        }
        private Double _ASP_유통_SS;
        public Double ASP_유통_SS
        {
            get { return _ASP_유통_SS; }
            set { _ASP_유통_SS = value; }
        }
        private Double _ASP_유통_소계;
        public Double ASP_유통_소계
        {
            get { return _ASP_유통_소계; }
            set { _ASP_유통_소계 = value; }
        }
        private Double _ASP_사업자_LG;
        public Double ASP_사업자_LG
        {
            get { return _ASP_사업자_LG; }
            set { _ASP_사업자_LG = value; }
        }
        private Double _ASP_사업자_SS;
        public Double ASP_사업자_SS
        {
            get { return _ASP_사업자_SS; }
            set { _ASP_사업자_SS = value; }
        }
        private Double _ASP_사업자_소계;
        public Double ASP_사업자_소계
        {
            get { return _ASP_사업자_소계; }
            set { _ASP_사업자_소계 = value; }
        }
        private Double _ASP_총계;
        public Double ASP_총계
        {
            get { return _ASP_총계; }
            set { _ASP_총계 = value; }
        }
        private Double _Rebate;
        public Double Rebate
        {
            get { return _Rebate; }
            set { _Rebate = value; }
        }

        public void setArrData_관리자데이터(string[] arrvalue)
        {
            int k = 0;
            도매_수익_월평균관리수수료 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도매_수익_CS관리수수료 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도매_비용_대리점투자금액_신규 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도매_비용_대리점투자금액_기변 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도매_비용_직원급여_간부급 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도매_비용_직원급여_평사원 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도매_비용_지급임차료 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도매_비용_운반비 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도매_비용_차량유지비 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도매_비용_지급수수료 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도매_비용_판매촉진비 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도매_비용_건물관리비 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            소매_수익_월평균업무취급수수료 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            소매_수익_직영매장판매수익 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            소매_비용_직원급여_간부급 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            소매_비용_직원급여_평사원 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            소매_비용_지급임차료 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            소매_비용_지급수수료 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            소매_비용_판매촉진비 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            소매_비용_건물관리비 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도소매_비용_복리후생비 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도소매_비용_통신비 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도소매_비용_공과금 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도소매_비용_소모품비 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도소매_비용_이자비용 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            도소매_비용_기타 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            ASP_유통_LG = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            ASP_유통_SS = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            ASP_유통_소계 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            ASP_사업자_LG = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            ASP_사업자_SS = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            ASP_사업자_소계 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            ASP_총계 = CommonUtil.StringToDoubleVal(arrvalue[k++]);
            Rebate = CommonUtil.StringToDoubleVal(arrvalue[k++]);
        }


        private String getFormatString(Double value)
        {
            CultureInfo cur = new CultureInfo(CultureInfo.InvariantCulture.LCID);
            cur.NumberFormat.NumberDecimalDigits = 0;
            return value.ToString("N", cur);
        }

        private Double getFormatDouble(String value)
        {             
            String temp = "0";
            temp = value.Replace(",", "");
            return CommonUtil.StringToDoubleVal(temp);
        }



    }
}
