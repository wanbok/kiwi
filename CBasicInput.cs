using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
namespace KIWI
{
    public class CBasicInput
    {
        //도매_누적가입자수
        private Double 도매_누적가입자수;
        public void set도매_누적가입자수(Double value)
        {
            도매_누적가입자수 = value;
        }
        public Double get도매_누적가입자수()
        {
            return 도매_누적가입자수;
        }
        public void set도매_누적가입자수(String value)
        {
            도매_누적가입자수 = getFormatDouble(value);
        }
        public String getstr도매_누적가입자수()
        {
            return getFormatString(도매_누적가입자수);
        }
      

        //도매_월평균판매대수_신규
        private Double 도매_월평균판매대수_신규;
        public void set도매_월평균판매대수_신규(Double value)
        {
            도매_월평균판매대수_신규 = value;
        }
        public Double get도매_월평균판매대수_신규()
        {
            return 도매_월평균판매대수_신규;
        }
        public void set도매_월평균판매대수_신규(String value)
        {
            도매_월평균판매대수_신규 = getFormatDouble(value);
        }
        public String getstr도매_월평균판매대수_신규()
        {
            return getFormatString(도매_월평균판매대수_신규);
        }

        //도매_월평균판매대수_기변
        private Double 도매_월평균판매대수_기변;
        public void set도매_월평균판매대수_기변(Double value)
        {
            도매_월평균판매대수_기변 = value;
        }
        public Double get도매_월평균판매대수_기변()
        {
            return 도매_월평균판매대수_기변;
        }
        public void set도매_월평균판매대수_기변(String value)
        {
            도매_월평균판매대수_기변 = getFormatDouble(value);
        }
        public String getstr도매_월평균판매대수_기변()
        {
            return getFormatString(도매_월평균판매대수_기변);
        }

        //도매_월평균유통모델출고대수_LG
        private Double 도매_월평균유통모델출고대수_LG;
        public void set도매_월평균유통모델출고대수_LG(Double value)
        {
            도매_월평균유통모델출고대수_LG = value;
        }
        public Double get도매_월평균유통모델출고대수_LG()
        {
            return 도매_월평균유통모델출고대수_LG;
        }
        public void set도매_월평균유통모델출고대수_LG(String value)
        {
            도매_월평균유통모델출고대수_LG = getFormatDouble(value);
        }
        public String getstr도매_월평균유통모델출고대수_LG()
        {
            return getFormatString(도매_월평균유통모델출고대수_LG);
        }

        //도매_월평균유통모델출고대수_SS
        private Double 도매_월평균유통모델출고대수_SS;
        public void set도매_월평균유통모델출고대수_SS(Double value)
        {
            도매_월평균유통모델출고대수_SS = value;
        }
        public Double get도매_월평균유통모델출고대수_SS()
        {
            return 도매_월평균유통모델출고대수_SS;
        }
        public void set도매_월평균유통모델출고대수_SS(String value)
        {
            도매_월평균유통모델출고대수_SS = getFormatDouble(value);
        }
        public String getstr도매_월평균유통모델출고대수_SS()
        {
            return getFormatString(도매_월평균유통모델출고대수_SS);
        }

        //도매_거래선수_개통사무실
        private Double 도매_거래선수_개통사무실;
        public void set도매_거래선수_개통사무실(Double value)
        {
            도매_거래선수_개통사무실 = value;
        }
        public Double get도매_거래선수_개통사무실()
        {
            return 도매_거래선수_개통사무실;
        }
        public void set도매_거래선수_개통사무실(String value)
        {
            도매_거래선수_개통사무실 = getFormatDouble(value);
        }
        public String getstr도매_거래선수_개통사무실()
        {
            return getFormatString(도매_거래선수_개통사무실);
        }

        //도매_거래선수_판매점
        private Double 도매_거래선수_판매점;
        public void set도매_거래선수_판매점(Double value)
        {
            도매_거래선수_판매점 = value;
        }
        public Double get도매_거래선수_판매점()
        {
            return 도매_거래선수_판매점;
        }
        public void set도매_거래선수_판매점(String value)
        {
            도매_거래선수_판매점 = getFormatDouble(value);
        }
        public String getstr도매_거래선수_판매점()
        {
            return getFormatString(도매_거래선수_판매점);
        }

        //도매_직원수_간부급
        private Double 도매_직원수_간부급;
        public void set도매_직원수_간부급(Double value)
        {
            도매_직원수_간부급 = value;
        }
        public Double get도매_직원수_간부급()
        {
            return 도매_직원수_간부급;
        }
        public void set도매_직원수_간부급(String value)
        {
            도매_직원수_간부급 = getFormatDouble(value);
        }
        public String getstr도매_직원수_간부급()
        {
            return getFormatString(도매_직원수_간부급);
        }

        //도매_직원수_평사원
        private Double 도매_직원수_평사원;
        public void set도매_직원수_평사원(Double value)
        {
            도매_직원수_평사원 = value;
        }
        public Double get도매_직원수_평사원()
        {
            return 도매_직원수_평사원;
        }
        public void set도매_직원수_평사원(String value)
        {
            도매_직원수_평사원 = getFormatDouble(value);
        }
        public String getstr도매_직원수_평사원()
        {
            return getFormatString(도매_직원수_평사원);
        }



        //소매_월평균판매대수_신규
        private Double 소매_월평균판매대수_신규;
        public void set소매_월평균판매대수_신규(Double value)
        {
            소매_월평균판매대수_신규 = value;
        }
        public Double get소매_월평균판매대수_신규()
        {
            return 소매_월평균판매대수_신규;
        }
        public void set소매_월평균판매대수_신규(String value)
        {
            소매_월평균판매대수_신규 = getFormatDouble(value);
        }
        public String getstr소매_월평균판매대수_신규()
        {
            return getFormatString(소매_월평균판매대수_신규);
        }

        //소매_월평균판매대수_기변
        private Double 소매_월평균판매대수_기변;
        public void set소매_월평균판매대수_기변(Double value)
        {
            소매_월평균판매대수_기변 = value;
        }
        public Double get소매_월평균판매대수_기변()
        {
            return 소매_월평균판매대수_기변;
        }
        public void set소매_월평균판매대수_기변(String value)
        {
            소매_월평균판매대수_기변 = getFormatDouble(value);
        }
        public String getstr소매_월평균판매대수_기변()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_월평균판매대수_기변);
        }

        //소매_거래선수_직영점
        private Double 소매_거래선수_직영점;
        public void set소매_거래선수_직영점(Double value)
        {
            소매_거래선수_직영점 = value;
        }
        public Double get소매_거래선수_직영점()
        {
            return 소매_거래선수_직영점;
        }
        public void set소매_거래선수_직영점(String value)
        {
            소매_거래선수_직영점 = getFormatDouble(value);
        }
        public String getstr소매_거래선수_직영점()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_거래선수_직영점);
        }

        //소매_직원수_간부급
        private Double 소매_직원수_간부급;
        public void set소매_직원수_간부급(Double value)
        {
            소매_직원수_간부급 = value;
        }
        public Double get소매_직원수_간부급()
        {
            return 소매_직원수_간부급;
        }
        public void set소매_직원수_간부급(String value)
        {
            소매_직원수_간부급 = getFormatDouble(value);
        }
        public String getstr소매_직원수_간부급()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_직원수_간부급);
        }

        //소매_직원수_평사원
        private Double 소매_직원수_평사원;
        public void set소매_직원수_평사원(Double value)
        {
            소매_직원수_평사원 = value;
        }
        public Double get소매_직원수_평사원()
        {
            return 소매_직원수_평사원;
        }
        public void set소매_직원수_평사원(String value)
        {
            소매_직원수_평사원 = getFormatDouble(value);
        }
        public String getstr소매_직원수_평사원()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_직원수_평사원);
        }






        //도매_누적가입자수
        //도매_월평균판매대수_신규
        //도매_월평균판매대수_기변
        //도매_월평균유통모델출고대수_LG
        //도매_월평균유통모델출고대수_SS
        //도매_거래선수_개통사무실
        //도매_거래선수_판매점
        //도매_직원수_간부급
        //도매_직원수_평사원

        //소매_월평균판매대수_신규
        //소매_월평균판매대수_기변
        //소매_거래선수_직영점
        //소매_직원수_간부급
        //소매_직원수_평사원


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
            도매_누적가입자수 = arrvalue[0];
            도매_월평균판매대수_신규 = arrvalue[1];
            도매_월평균판매대수_기변 = arrvalue[2];
            도매_월평균유통모델출고대수_LG = arrvalue[3];
            도매_월평균유통모델출고대수_SS = arrvalue[4];
            도매_거래선수_개통사무실 = arrvalue[5];
            도매_거래선수_판매점 = arrvalue[6];
            도매_직원수_간부급 = arrvalue[7];
            도매_직원수_평사원 = arrvalue[8];

            소매_월평균판매대수_신규 = arrvalue[9];
            소매_월평균판매대수_기변 = arrvalue[10];
            소매_거래선수_직영점 = arrvalue[11];
            소매_직원수_간부급 = arrvalue[12];
            소매_직원수_평사원 = arrvalue[13];
        }

        public Double[] getArrData()
        {
            Double[] arrvalue = new Double[14];

            arrvalue[0] = 도매_누적가입자수;
            arrvalue[1] = 도매_월평균판매대수_신규;
            arrvalue[2] = 도매_월평균판매대수_기변;
            arrvalue[3] = 도매_월평균유통모델출고대수_LG;
            arrvalue[4] = 도매_월평균유통모델출고대수_SS;
            arrvalue[5] = 도매_거래선수_개통사무실;
            arrvalue[6] = 도매_거래선수_판매점;
            arrvalue[7] = 도매_직원수_간부급;
            arrvalue[8] = 도매_직원수_평사원;

            arrvalue[9] = 소매_월평균판매대수_신규;
            arrvalue[10] = 소매_월평균판매대수_기변;
            arrvalue[11] = 소매_거래선수_직영점;
            arrvalue[12] = 소매_직원수_간부급;
            arrvalue[13] = 소매_직원수_평사원;


            return arrvalue;
        }

        public void setArrData_BasicInput(string[] arrvalue)
        {
            도매_누적가입자수 = CommonUtil.StringToDoubleVal(arrvalue[0]);
            도매_월평균판매대수_신규 = CommonUtil.StringToDoubleVal(arrvalue[1]);
            도매_월평균판매대수_기변 = CommonUtil.StringToDoubleVal(arrvalue[2]);
            도매_월평균유통모델출고대수_LG = CommonUtil.StringToDoubleVal(arrvalue[3]);
            도매_월평균유통모델출고대수_SS = CommonUtil.StringToDoubleVal(arrvalue[4]);
            도매_거래선수_개통사무실 = CommonUtil.StringToDoubleVal(arrvalue[5]);
            도매_거래선수_판매점 = CommonUtil.StringToDoubleVal(arrvalue[6]);
            도매_직원수_간부급 = CommonUtil.StringToDoubleVal(arrvalue[7]);
            도매_직원수_평사원 = CommonUtil.StringToDoubleVal(arrvalue[8]);

            소매_월평균판매대수_신규 = CommonUtil.StringToDoubleVal(arrvalue[9]);
            소매_월평균판매대수_기변 = CommonUtil.StringToDoubleVal(arrvalue[10]);
            소매_거래선수_직영점 = CommonUtil.StringToDoubleVal(arrvalue[11]);
            소매_직원수_간부급 = CommonUtil.StringToDoubleVal(arrvalue[12]);
            소매_직원수_평사원 = CommonUtil.StringToDoubleVal(arrvalue[13]);
        }

        public Double[] getArrData_리포트용()
        {
            Double[] arrvalue = new Double[10];

            int i = 0;
            arrvalue[i++] = 도매_누적가입자수;
            arrvalue[i++] = 도매_월평균판매대수_신규;
            arrvalue[i++] = 소매_월평균판매대수_신규;
            arrvalue[i++] = 도매_월평균판매대수_기변;
            arrvalue[i++] = 소매_월평균판매대수_기변;
            arrvalue[i++] = 도매_월평균유통모델출고대수_LG + 도매_월평균유통모델출고대수_SS;
            arrvalue[i++] = 도매_거래선수_개통사무실 + 도매_거래선수_판매점;
            arrvalue[i++] = 소매_거래선수_직영점;
            arrvalue[i++] = 도매_직원수_간부급 + 도매_직원수_평사원;
            arrvalue[i++] = 소매_직원수_간부급 + 소매_직원수_평사원;

            return arrvalue;
        }

        public Double[] getArrData_BasicInput()
        {
            Double[] arrvalue = new Double[35];

            int i = 0;
            arrvalue[i++] = 도매_누적가입자수;
            arrvalue[i++] = 도매_월평균판매대수_신규;
            arrvalue[i++] = 도매_월평균판매대수_기변;
            arrvalue[i++] = get도매_월평균판매대수_소계();
            arrvalue[i++] = 도매_월평균유통모델출고대수_LG;
            arrvalue[i++] = 도매_월평균유통모델출고대수_SS;
            arrvalue[i++] = get도매_월평균유통모델출고대수_소계();
            arrvalue[i++] = 도매_거래선수_개통사무실;
            arrvalue[i++] = 도매_거래선수_판매점;
            arrvalue[i++] = get도매_거래선수_소계();
            arrvalue[i++] = 도매_직원수_간부급;
            arrvalue[i++] = 도매_직원수_평사원;
            arrvalue[i++] = get도매_직원수_소계();
            arrvalue[i++] = 소매_월평균판매대수_신규;
            arrvalue[i++] = 소매_월평균판매대수_기변;
            arrvalue[i++] = get소매_월평균판매대수_소계();
            arrvalue[i++] = 소매_거래선수_직영점;
            arrvalue[i++] = get소매_거래선수_소계();
            arrvalue[i++] = 소매_직원수_간부급;
            arrvalue[i++] = 소매_직원수_평사원;
            arrvalue[i++] = get소매_직원수_소계();
            arrvalue[i++] = get누적가입자수_합계();
            arrvalue[i++] = get월평균판매대수_신규_합계();
            arrvalue[i++] = get월평균판매대수_기변_합계();
            arrvalue[i++] = get월평균판매대수_소계_합계();
            arrvalue[i++] = get월평균유통모델출고대수_LG_합계();
            arrvalue[i++] = get월평균유통모델출고대수_SS_합계();
            arrvalue[i++] = get월평균유통모델출고대수_소계_합계();
            arrvalue[i++] = get거래선수_개통사무실_합계();
            arrvalue[i++] = get거래선수_직영점_합계();
            arrvalue[i++] = get거래선수_판매점_합계();
            arrvalue[i++] = get거래선수_소계_합계();
            arrvalue[i++] = get직원수_간부급_합계();
            arrvalue[i++] = get직원수_평사원_합계();
            arrvalue[i++] = get직원수_소계_합계();

            return arrvalue;
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


        //지역
        private String 지역;
        public void set지역(String value)
        {
            지역 = value;
        }
        public String get지역()
        {
            return 지역;
        }
        //대리점
        private String 대리점;
        public void set대리점(String value)
        {
            대리점 = value;
        }
        public String get대리점()
        {
            return 대리점;
        }
        //마케터
        private String 마케터;
        public void set마케터(String value)
        {
            마케터 = value;
        }
        public String get마케터()
        {
            return 마케터;
        }

        //소계값
        //도매_월평균판매대수_소계
        public Double get도매_월평균판매대수_소계()
        {
            return (도매_월평균판매대수_신규 + 도매_월평균판매대수_기변);
        }
        public String getstr도매_월평균판매대수_소계()
        {
            return getFormatString((도매_월평균판매대수_신규 + 도매_월평균판매대수_기변));
        }

        //도매_월평균유통모델출고대수_소계
        public Double get도매_월평균유통모델출고대수_소계()
        {
            return 도매_월평균유통모델출고대수_LG + 도매_월평균유통모델출고대수_SS;
        }
        public String getstr도매_월평균유통모델출고대수_소계()
        {
            return getFormatString(도매_월평균유통모델출고대수_LG + 도매_월평균유통모델출고대수_SS);
        }

        //도매_거래선수_소계
        public Double get도매_거래선수_소계()
        {
            return 도매_거래선수_개통사무실 + 도매_거래선수_판매점;
        }
        public String getstr도매_거래선수_소계()
        {
            return getFormatString(도매_거래선수_개통사무실 + 도매_거래선수_판매점);
        }

        //도매_직원수_소계
        public Double get도매_직원수_소계()
        {
            return 도매_직원수_간부급 + 도매_직원수_평사원;
        }
        public String getstr도매_직원수_소계()
        {
            return getFormatString(도매_직원수_간부급 + 도매_직원수_평사원);
        }


        //소매_월평균판매대수_소계
        public Double get소매_월평균판매대수_소계()
        {
            return 소매_월평균판매대수_신규 + 소매_월평균판매대수_기변;
        }
        public String getstr소매_월평균판매대수_소계()
        {
            return getFormatString(소매_월평균판매대수_신규 + 소매_월평균판매대수_기변);
        }

        //소매_거래선수_소계
        public Double get소매_거래선수_소계()
        {
            return 소매_거래선수_직영점;
        }
        public String getstr소매_거래선수_소계()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_거래선수_직영점);
        }

        //소매_직원수_소계
        public Double get소매_직원수_소계()
        {
            return 소매_직원수_간부급 + 소매_직원수_평사원;
        }
        public String getstr소매_직원수_소계()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_직원수_간부급 + 소매_직원수_평사원);
        }



        //합계값
        //누적가입자수_합계
        public Double get누적가입자수_합계()
        {
            return 도매_누적가입자수;
        }
        public String getstr누적가입자수_합계()
        {
            return getFormatString(도매_누적가입자수);
        }


        //월평균판매대수_신규_합계
        public Double get월평균판매대수_신규_합계()
        {
            return 도매_월평균판매대수_신규 + 소매_월평균판매대수_신규;
        }
        public String getstr월평균판매대수_신규_합계()
        {
            return getFormatString(도매_월평균판매대수_신규 + 소매_월평균판매대수_신규);
        }

        //월평균판매대수_기변_합계
        public Double get월평균판매대수_기변_합계()
        {
            return 도매_월평균판매대수_기변 + 소매_월평균판매대수_기변;
        }
        public String getstr월평균판매대수_기변_합계()
        {
            return getFormatString(도매_월평균판매대수_신규 + 소매_월평균판매대수_신규);
        }

        //월평균유통모델출고대수_LG
        public Double get월평균유통모델출고대수_LG_합계()
        {
            return 도매_월평균유통모델출고대수_LG;
        }
        public String getstr월평균유통모델출고대수_LG_합계()
        {
            return getFormatString(도매_월평균유통모델출고대수_LG);
        }

        //월평균유통모델출고대수_SS
        public Double get월평균유통모델출고대수_SS_합계()
        {
            return 도매_월평균유통모델출고대수_SS;
        }
        public String getstr월평균유통모델출고대수_SS_합계()
        {
            return getFormatString(도매_월평균유통모델출고대수_SS);
        }

        //거래선수_개통사무실
        public Double get거래선수_개통사무실_합계()
        {
            return 도매_거래선수_개통사무실;
        }
        public String getstr거래선수_개통사무실_합계()
        {
            return getFormatString(도매_거래선수_개통사무실);
        }

        //거래선수_판매점
        public Double get거래선수_판매점_합계()
        {
            return 도매_거래선수_판매점;
        }
        public String getstr거래선수_판매점_합계()
        {
            return getFormatString(도매_거래선수_판매점);
        }

        //직원수_간부급
        public Double get직원수_간부급_합계()
        {
            return 도매_직원수_간부급 + 소매_직원수_간부급;
        }
        public String getstr직원수_간부급_합계()
        {
            return getFormatString(도매_직원수_간부급 + 소매_직원수_간부급);
        }

        //직원수_평사원
        public Double get직원수_평사원_합계()
        {
            return 도매_직원수_평사원 + 소매_직원수_평사원;
        }
        public String getstr직원수_평사원_합계()
        {
            return getFormatString(도매_직원수_평사원 + 소매_직원수_평사원);
        }



        //거래선수_직영점
        public Double get거래선수_직영점_합계()
        {
            return 소매_거래선수_직영점;
        }
        public String getstr거래선수_직영점_합계()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_거래선수_직영점);
        }

        //소계의 합계

        //월평균판매대수_소계_합계
        public Double get월평균판매대수_소계_합계()
        {
            return get월평균판매대수_신규_합계() + get월평균판매대수_기변_합계();
        }
        public String getstr월평균판매대수_소계_합계()
        {
            return getFormatString(get월평균판매대수_신규_합계() + get월평균판매대수_기변_합계());
        }

        //월평균유통모델출고대수_소계_합계
        public Double get월평균유통모델출고대수_소계_합계()
        {
            return get월평균유통모델출고대수_LG_합계() + get월평균유통모델출고대수_SS_합계();
        }
        public String getstr월평균유통모델출고대수_소계_합계()
        {
            return getFormatString(get월평균유통모델출고대수_LG_합계() + get월평균유통모델출고대수_SS_합계());
        }

        //거래선수_소계_합계
        public Double get거래선수_소계_합계()
        {
            return get거래선수_개통사무실_합계() + get거래선수_직영점_합계() + get거래선수_판매점_합계();
        }
        public String getstr거래선수_소계_합계()
        {
            return getFormatString(get거래선수_개통사무실_합계() + get거래선수_직영점_합계() + get거래선수_판매점_합계());
        }

        //직원수_소계_합계
        public Double get직원수_소계_합계()
        {
            return get직원수_간부급_합계() + get직원수_평사원_합계();
        }
        public String getstr직원수_소계_합계()
        {
            return getFormatString(get직원수_간부급_합계() + get직원수_평사원_합계());
        }


    }
}
