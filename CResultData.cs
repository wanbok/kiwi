using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace KIWI
{
    public class CResultData
    {
        //결과화면에서 사용
        //시뮬레이션 결과 사용






        //수익은 전체,도매,소매 중복 없음
        //비용은 도매, 소매 일부 중복으로 개별 I/O 생성 ,중복안되는 항목도 개별 I/O 생성


        //******전체 수익
        //도매_수익_가입자관리수수료
        //도매_수익_CS관리수수료                    
        //소매_수익_업무취급수수료
        //도매_수익_사업자모델매입에따른추가수익
        //도매_수익_유통모델매입에따른추가수익_현금_Volume
        //소매_수익_직영매장판매수익       

        //******전체 비용
        //전체_비용_대리점투자비용
        //전체_비용_인건비_급여_복리후생비
        //전체_비용_임차료
        //전체_비용_이자비용
        //전체_비용_부가세
        //전체_비용_법인세
        //전체_비용_기타판매관리비

        //******도매 비용
        //도매_비용_대리점투자비용
        //도매_비용_인건비_급여_복리후생비
        //도매_비용_임차료
        //도매_비용_이자비용
        //도매_비용_부가세
        //도매_비용_법인세
        //도매_비용_기타판매관리비

        //******소매 비용
        //소매_비용_인건비_급여_복리후생비
        //소매_비용_임차료
        //소매_비용_이자비용
        //소매_비용_부가세
        //소매_비용_법인세
        //소매_비용_기타판매관리비

        


        //전체 수익
        //도매_수익_가입자관리수수료
        private Int64 도매_수익_가입자관리수수료;
        public void set도매_수익_가입자관리수수료(Int64 value)
        {
            도매_수익_가입자관리수수료 = value;
        }
        public Int64 get도매_수익_가입자관리수수료()
        {
            return 도매_수익_가입자관리수수료;
        }
        public void set도매_수익_가입자관리수수료(String value)
        {
            도매_수익_가입자관리수수료 = getFormatInt64(value);
        }
        public String getstr도매_수익_가입자관리수수료()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_수익_가입자관리수수료);
        }
        //도매_수익_CS관리수수료                    
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
        //소매_수익_업무취급수수료
        private Int64 소매_수익_업무취급수수료;
        public void set소매_수익_업무취급수수료(Int64 value)
        {
            소매_수익_업무취급수수료 = value;
        }
        public Int64 get소매_수익_업무취급수수료()
        {
            return 소매_수익_업무취급수수료;
        }
        public void set소매_수익_업무취급수수료(String value)
        {
            소매_수익_업무취급수수료 = getFormatInt64(value);
        }
        public String getstr소매_수익_업무취급수수료()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_수익_업무취급수수료);
        }
        //도매_수익_사업자모델매입에따른추가수익
        private Int64 도매_수익_사업자모델매입에따른추가수익;
        public void set도매_수익_사업자모델매입에따른추가수익(Int64 value)
        {
            도매_수익_사업자모델매입에따른추가수익 = value;
        }
        public Int64 get도매_수익_사업자모델매입에따른추가수익()
        {
            return 도매_수익_사업자모델매입에따른추가수익;
        }
        public void set도매_수익_사업자모델매입에따른추가수익(String value)
        {
            도매_수익_사업자모델매입에따른추가수익 = getFormatInt64(value);
        }
        public String getstr도매_수익_사업자모델매입에따른추가수익()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_수익_사업자모델매입에따른추가수익);
        }
        //도매_수익_유통모델매입에따른추가수익_현금_Volume
        private Int64 도매_수익_유통모델매입에따른추가수익_현금_Volume;
        public void set도매_수익_유통모델매입에따른추가수익_현금_Volume(Int64 value)
        {
            도매_수익_유통모델매입에따른추가수익_현금_Volume = value;
        }
        public Int64 get도매_수익_유통모델매입에따른추가수익_현금_Volume()
        {
            return 도매_수익_유통모델매입에따른추가수익_현금_Volume;
        }
        public void set도매_수익_유통모델매입에따른추가수익_현금_Volume(String value)
        {
            도매_수익_유통모델매입에따른추가수익_현금_Volume = getFormatInt64(value);
        }
        public String getstr도매_수익_유통모델매입에따른추가수익_현금_Volume()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_수익_유통모델매입에따른추가수익_현금_Volume);
        }
        //소매_수익_직영매장판매수익
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
       


        //******전체 비용
        //전체_비용_대리점투자비용
        private Int64 전체_비용_대리점투자비용;
        public void set전체_비용_대리점투자비용(Int64 value)
        {
            전체_비용_대리점투자비용 = value;
        }
        public Int64 get전체_비용_대리점투자비용()
        {
            return 전체_비용_대리점투자비용;
        }
        public void set전체_비용_대리점투자비용(String value)
        {
            전체_비용_대리점투자비용 = getFormatInt64(value);
        }
        public String getstr전체_비용_대리점투자비용()
        {//   ','가 적용된 값 리턴
            return getFormatString(전체_비용_대리점투자비용);
        }
        //전체_비용_인건비_급여_복리후생비
        private Int64 전체_비용_인건비_급여_복리후생비;
        public void set전체_비용_인건비_급여_복리후생비(Int64 value)
        {
            전체_비용_인건비_급여_복리후생비 = value;
        }
        public Int64 get전체_비용_인건비_급여_복리후생비()
        {
            return 전체_비용_인건비_급여_복리후생비;
        }
        public void set전체_비용_인건비_급여_복리후생비(String value)
        {
            전체_비용_인건비_급여_복리후생비 = getFormatInt64(value);
        }
        public String getstr전체_비용_인건비_급여_복리후생비()
        {//   ','가 적용된 값 리턴
            return getFormatString(전체_비용_인건비_급여_복리후생비);
        }
        //전체_비용_임차료
        private Int64 전체_비용_임차료;
        public void set전체_비용_임차료(Int64 value)
        {
            전체_비용_임차료 = value;
        }
        public Int64 get전체_비용_임차료()
        {
            return 전체_비용_임차료;
        }
        public void set전체_비용_임차료(String value)
        {
            전체_비용_임차료 = getFormatInt64(value);
        }
        public String getstr전체_비용_임차료()
        {//   ','가 적용된 값 리턴
            return getFormatString(전체_비용_임차료);
        }
        //전체_비용_이자비용
        private Int64 전체_비용_이자비용;
        public void set전체_비용_이자비용(Int64 value)
        {
            전체_비용_이자비용 = value;
        }
        public Int64 get전체_비용_이자비용()
        {
            return 전체_비용_이자비용;
        }
        public void set전체_비용_이자비용(String value)
        {
            전체_비용_이자비용 = getFormatInt64(value);
        }
        public String getstr전체_비용_이자비용()
        {//   ','가 적용된 값 리턴
            return getFormatString(전체_비용_이자비용);
        }
        //전체_비용_부가세
        private Int64 전체_비용_부가세;
        public void set전체_비용_부가세(Int64 value)
        {
            전체_비용_부가세 = value;
        }
        public Int64 get전체_비용_부가세()
        {
            return 전체_비용_부가세;
        }
        public void set전체_비용_부가세(String value)
        {
            전체_비용_부가세 = getFormatInt64(value);
        }
        public String getstr전체_비용_부가세()
        {//   ','가 적용된 값 리턴
            return getFormatString(전체_비용_부가세);
        }
        //전체_비용_법인세
        private Int64 전체_비용_법인세;
        public void set전체_비용_법인세(Int64 value)
        {
            전체_비용_법인세 = value;
        }
        public Int64 get전체_비용_법인세()
        {
            return 전체_비용_법인세;
        }
        public void set전체_비용_법인세(String value)
        {
            전체_비용_법인세 = getFormatInt64(value);
        }
        public String getstr전체_비용_법인세()
        {//   ','가 적용된 값 리턴
            return getFormatString(전체_비용_법인세);
        }
        //전체_비용_기타판매관리비
        private Int64 전체_비용_기타판매관리비;
        public void set전체_비용_기타판매관리비(Int64 value)
        {
            전체_비용_기타판매관리비 = value;
        }
        public Int64 get전체_비용_기타판매관리비()
        {
            return 전체_비용_기타판매관리비;
        }
        public void set전체_비용_기타판매관리비(String value)
        {
            전체_비용_기타판매관리비 = getFormatInt64(value);
        }
        public String getstr전체_비용_기타판매관리비()
        {//   ','가 적용된 값 리턴
            return getFormatString(전체_비용_기타판매관리비);
        }

        //******도매 비용
        //도매_비용_대리점투자비용
        private Int64 도매_비용_대리점투자비용;
        public void set도매_비용_대리점투자비용(Int64 value)
        {
            도매_비용_대리점투자비용 = value;
        }
        public Int64 get도매_비용_대리점투자비용()
        {
            return 도매_비용_대리점투자비용;
        }
        public void set도매_비용_대리점투자비용(String value)
        {
            도매_비용_대리점투자비용 = getFormatInt64(value);
        }
        public String getstr도매_비용_대리점투자비용()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_대리점투자비용);
        }
        //도매_비용_인건비_급여_복리후생비
        private Int64 도매_비용_인건비_급여_복리후생비;
        public void set도매_비용_인건비_급여_복리후생비(Int64 value)
        {
            도매_비용_인건비_급여_복리후생비 = value;
        }
        public Int64 get도매_비용_인건비_급여_복리후생비()
        {
            return 도매_비용_인건비_급여_복리후생비;
        }
        public void set도매_비용_인건비_급여_복리후생비(String value)
        {
            도매_비용_인건비_급여_복리후생비 = getFormatInt64(value);
        }
        public String getstr도매_비용_인건비_급여_복리후생비()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_인건비_급여_복리후생비);
        }
        //도매_비용_임차료
        private Int64 도매_비용_임차료;
        public void set도매_비용_임차료(Int64 value)
        {
            도매_비용_임차료 = value;
        }
        public Int64 get도매_비용_임차료()
        {
            return 도매_비용_임차료;
        }
        public void set도매_비용_임차료(String value)
        {
            도매_비용_임차료 = getFormatInt64(value);
        }
        public String getstr도매_비용_임차료()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_임차료);
        }
        //도매_비용_이자비용
        private Int64 도매_비용_이자비용;
        public void set도매_비용_이자비용(Int64 value)
        {
            도매_비용_이자비용 = value;
        }
        public Int64 get도매_비용_이자비용()
        {
            return 도매_비용_이자비용;
        }
        public void set도매_비용_이자비용(String value)
        {
            도매_비용_이자비용 = getFormatInt64(value);
        }
        public String getstr도매_비용_이자비용()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_이자비용);
        }
        //도매_비용_부가세
        private Int64 도매_비용_부가세;
        public void set도매_비용_부가세(Int64 value)
        {
            도매_비용_부가세 = value;
        }
        public Int64 get도매_비용_부가세()
        {
            return 도매_비용_부가세;
        }
        public void set도매_비용_부가세(String value)
        {
            도매_비용_부가세 = getFormatInt64(value);
        }
        public String getstr도매_비용_부가세()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_부가세);
        }
        //도매_비용_법인세
        private Int64 도매_비용_법인세;
        public void set도매_비용_법인세(Int64 value)
        {
            도매_비용_법인세 = value;
        }
        public Int64 get도매_비용_법인세()
        {
            return 도매_비용_법인세;
        }
        public void set도매_비용_법인세(String value)
        {
            도매_비용_법인세 = getFormatInt64(value);
        }
        public String getstr도매_비용_법인세()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_법인세);
        }
        //도매_비용_기타판매관리비
        private Int64 도매_비용_기타판매관리비;
        public void set도매_비용_기타판매관리비(Int64 value)
        {
            도매_비용_기타판매관리비 = value;
        }
        public Int64 get도매_비용_기타판매관리비()
        {
            return 도매_비용_기타판매관리비;
        }
        public void set도매_비용_기타판매관리비(String value)
        {
            도매_비용_기타판매관리비 = getFormatInt64(value);
        }
        public String getstr도매_비용_기타판매관리비()
        {//   ','가 적용된 값 리턴
            return getFormatString(도매_비용_기타판매관리비);
        }

        //******소매 비용
        //소매_비용_인건비_급여_복리후생비
        private Int64 소매_비용_인건비_급여_복리후생비;
        public void set소매_비용_인건비_급여_복리후생비(Int64 value)
        {
            소매_비용_인건비_급여_복리후생비 = value;
        }
        public Int64 get소매_비용_인건비_급여_복리후생비()
        {
            return 소매_비용_인건비_급여_복리후생비;
        }
        public void set소매_비용_인건비_급여_복리후생비(String value)
        {
            소매_비용_인건비_급여_복리후생비 = getFormatInt64(value);
        }
        public String getstr소매_비용_인건비_급여_복리후생비()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_인건비_급여_복리후생비);
        }
        //소매_비용_임차료
        private Int64 소매_비용_임차료;
        public void set소매_비용_임차료(Int64 value)
        {
            소매_비용_임차료 = value;
        }
        public Int64 get소매_비용_임차료()
        {
            return 소매_비용_임차료;
        }
        public void set소매_비용_임차료(String value)
        {
            소매_비용_임차료 = getFormatInt64(value);
        }
        public String getstr소매_비용_임차료()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_임차료);
        }
        //소매_비용_이자비용
        private Int64 소매_비용_이자비용;
        public void set소매_비용_이자비용(Int64 value)
        {
            소매_비용_이자비용 = value;
        }
        public Int64 get소매_비용_이자비용()
        {
            return 소매_비용_이자비용;
        }
        public void set소매_비용_이자비용(String value)
        {
            소매_비용_이자비용 = getFormatInt64(value);
        }
        public String getstr소매_비용_이자비용()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_이자비용);
        }
        //소매_비용_부가세
        private Int64 소매_비용_부가세;
        public void set소매_비용_부가세(Int64 value)
        {
            소매_비용_부가세 = value;
        }
        public Int64 get소매_비용_부가세()
        {
            return 소매_비용_부가세;
        }
        public void set소매_비용_부가세(String value)
        {
            소매_비용_부가세 = getFormatInt64(value);
        }
        public String getstr소매_비용_부가세()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_부가세);
        }
        //소매_비용_법인세
        private Int64 소매_비용_법인세;
        public void set소매_비용_법인세(Int64 value)
        {
            소매_비용_법인세 = value;
        }
        public Int64 get소매_비용_법인세()
        {
            return 소매_비용_법인세;
        }
        public void set소매_비용_법인세(String value)
        {
            소매_비용_법인세 = getFormatInt64(value);
        }
        public String getstr소매_비용_법인세()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_법인세);
        }
        //소매_비용_기타판매관리비
        private Int64 소매_비용_기타판매관리비;
        public void set소매_비용_기타판매관리비(Int64 value)
        {
            소매_비용_기타판매관리비 = value;
        }
        public Int64 get소매_비용_기타판매관리비()
        {
            return 소매_비용_기타판매관리비;
        }
        public void set소매_비용_기타판매관리비(String value)
        {
            소매_비용_기타판매관리비 = getFormatInt64(value);
        }
        public String getstr소매_비용_기타판매관리비()
        {//   ','가 적용된 값 리턴
            return getFormatString(소매_비용_기타판매관리비);
        }

        
        public void setArrData(Int64[] arrvalue)
        {
            //******전체 수익
            도매_수익_가입자관리수수료 = arrvalue[0];
            도매_수익_CS관리수수료 = arrvalue[1];
            소매_수익_업무취급수수료 = arrvalue[2];
            도매_수익_사업자모델매입에따른추가수익 = arrvalue[3];
            도매_수익_유통모델매입에따른추가수익_현금_Volume = arrvalue[4];
            소매_수익_직영매장판매수익= arrvalue[5];

            //******전체 비용
            전체_비용_대리점투자비용 = arrvalue[6];
            전체_비용_인건비_급여_복리후생비 = arrvalue[7];
            전체_비용_임차료 = arrvalue[8];
            전체_비용_이자비용 = arrvalue[9];
            전체_비용_부가세 = arrvalue[10];
            전체_비용_법인세 = arrvalue[11];
            전체_비용_기타판매관리비 = arrvalue[12];

            //******도매 비용
            도매_비용_대리점투자비용 = arrvalue[13];
            도매_비용_인건비_급여_복리후생비 = arrvalue[14];
            도매_비용_임차료 = arrvalue[15];
            도매_비용_이자비용 = arrvalue[16];
            도매_비용_부가세 = arrvalue[17];
            도매_비용_법인세 = arrvalue[18];
            도매_비용_기타판매관리비 = arrvalue[19];

            //******소매 비용
            소매_비용_인건비_급여_복리후생비 = arrvalue[20];
            소매_비용_임차료 = arrvalue[21];
            소매_비용_이자비용 = arrvalue[22];
            소매_비용_부가세 = arrvalue[23];
            소매_비용_법인세 = arrvalue[24];
            소매_비용_기타판매관리비 = arrvalue[25];
        }

        public Int64[] getArrData()
        {
            Int64[] arrvalue = new Int64[26];

            //******전체 수익
            arrvalue[0] = 도매_수익_가입자관리수수료;
            arrvalue[1] = 도매_수익_CS관리수수료;
            arrvalue[2] = 소매_수익_업무취급수수료;
            arrvalue[3] = 도매_수익_사업자모델매입에따른추가수익;
            arrvalue[4] = 도매_수익_유통모델매입에따른추가수익_현금_Volume;
            arrvalue[5] = 소매_수익_직영매장판매수익;

            //******전체 비용
            arrvalue[6] = 전체_비용_대리점투자비용;
            arrvalue[7] = 전체_비용_인건비_급여_복리후생비;
            arrvalue[8] = 전체_비용_임차료;
            arrvalue[9] = 전체_비용_이자비용;
            arrvalue[10] = 전체_비용_부가세;
            arrvalue[11] = 전체_비용_법인세;
            arrvalue[12] = 전체_비용_기타판매관리비;

            //******도매 비용
            arrvalue[13] = 도매_비용_대리점투자비용;
            arrvalue[14] = 도매_비용_인건비_급여_복리후생비;
            arrvalue[15] = 도매_비용_임차료;
            arrvalue[16] = 도매_비용_이자비용;
            arrvalue[17] = 도매_비용_부가세;
            arrvalue[18] = 도매_비용_법인세;
            arrvalue[19] = 도매_비용_기타판매관리비;

            //******소매 비용
            arrvalue[20] = 소매_비용_인건비_급여_복리후생비;
            arrvalue[21] = 소매_비용_임차료;
            arrvalue[22] = 소매_비용_이자비용;
            arrvalue[23] = 소매_비용_부가세;
            arrvalue[24] = 소매_비용_법인세;
            arrvalue[25] = 소매_비용_기타판매관리비;

            return arrvalue;
        }

        public Int64[] getArr전체_수익_비용_및_계산포함()
        {
            Int64[] arrvalue = new Int64[16];

            int i = 0;
            //******전체 수익
            arrvalue[i++] = 도매_수익_가입자관리수수료;
            arrvalue[i++] = 도매_수익_CS관리수수료;
            arrvalue[i++] = 소매_수익_업무취급수수료;
            arrvalue[i++] = 도매_수익_사업자모델매입에따른추가수익;
            arrvalue[i++] = 도매_수익_유통모델매입에따른추가수익_현금_Volume;
            arrvalue[i++] = 소매_수익_직영매장판매수익;

            // 소계
            Int64 전체_수익_소계 = 도매_수익_가입자관리수수료 + 도매_수익_CS관리수수료 +
                소매_수익_업무취급수수료 + 도매_수익_사업자모델매입에따른추가수익 +
                도매_수익_유통모델매입에따른추가수익_현금_Volume + 소매_수익_직영매장판매수익;
            arrvalue[i++] = 전체_수익_소계;

            //******전체 비용
            arrvalue[i++] = 전체_비용_대리점투자비용;
            arrvalue[i++] = 전체_비용_인건비_급여_복리후생비;
            arrvalue[i++] = 전체_비용_임차료;
            arrvalue[i++] = 전체_비용_이자비용;
            arrvalue[i++] = 전체_비용_부가세;
            arrvalue[i++] = 전체_비용_법인세;
            arrvalue[i++] = 전체_비용_기타판매관리비;

            // 소계
            Int64 전체_비용_소계 = 전체_비용_대리점투자비용 + 전체_비용_인건비_급여_복리후생비 +
                전체_비용_임차료 + 전체_비용_이자비용 + 전체_비용_부가세 +
                전체_비용_법인세 + 전체_비용_기타판매관리비;
            arrvalue[i++] = 전체_비용_소계;

            // 손익계
            arrvalue[i++] = 전체_수익_소계 - 전체_비용_소계;

            return arrvalue;
        }

        public Int64[] getArr도매_수익_비용_및_계산포함()
        {
            Int64[] arrvalue = new Int64[14];

            int i = 0;
            //******전체 수익
            arrvalue[i++] = 도매_수익_가입자관리수수료;
            arrvalue[i++] = 도매_수익_CS관리수수료;
            arrvalue[i++] = 도매_수익_사업자모델매입에따른추가수익;
            arrvalue[i++] = 도매_수익_유통모델매입에따른추가수익_현금_Volume;

            // 소계
            Int64 도매_수익_소계 = 도매_수익_가입자관리수수료 + 도매_수익_CS관리수수료 +
                도매_수익_사업자모델매입에따른추가수익 + 도매_수익_유통모델매입에따른추가수익_현금_Volume;
            arrvalue[i++] = 도매_수익_소계;

            //******도매 비용
            arrvalue[i++] = 도매_비용_대리점투자비용;
            arrvalue[i++] = 도매_비용_인건비_급여_복리후생비;
            arrvalue[i++] = 도매_비용_임차료;
            arrvalue[i++] = 도매_비용_이자비용;
            arrvalue[i++] = 도매_비용_부가세;
            arrvalue[i++] = 도매_비용_법인세;
            arrvalue[i++] = 도매_비용_기타판매관리비;

            // 소계
            Int64 도매_비용_소계 = 도매_비용_대리점투자비용 + 도매_비용_인건비_급여_복리후생비 +
                도매_비용_임차료 + 도매_비용_이자비용 + 도매_비용_부가세 + 도매_비용_법인세 +
                도매_비용_기타판매관리비;
            arrvalue[i++] = 도매_비용_소계;

            // 손익계
            arrvalue[i++] = 도매_수익_소계 - 도매_비용_소계;

            return arrvalue;
        }

        public Int64[] getArr소매_수익_비용_및_계산포함()
        {
            Int64[] arrvalue = new Int64[12];

            int i = 0;
            //******전체 수익
            arrvalue[i++] = 소매_수익_업무취급수수료;
            arrvalue[i++] = 소매_수익_직영매장판매수익;

            // 소계
            Int64 소매_수익_소계 = 소매_수익_업무취급수수료 + 소매_수익_직영매장판매수익;
            arrvalue[i++] = 소매_수익_소계;

            //******소매 비용
            arrvalue[i++] = 소매_비용_인건비_급여_복리후생비;
            arrvalue[i++] = 소매_비용_임차료;
            arrvalue[i++] = 소매_비용_이자비용;
            arrvalue[i++] = 소매_비용_부가세;
            arrvalue[i++] = 소매_비용_법인세;
            arrvalue[i++] = 소매_비용_기타판매관리비;

            // 소계
            Int64 소매_비용_소계 = 소매_비용_인건비_급여_복리후생비 + 소매_비용_임차료 +
                소매_비용_이자비용 + 소매_비용_부가세 + 소매_비용_법인세 + 소매_비용_기타판매관리비;
            arrvalue[i++] = 소매_비용_소계;

            // 손익계
            arrvalue[i++] = 소매_수익_소계 - 소매_비용_소계;

            // 점별손익추정
            if (CDataControl.g_BasicInput != null && CDataControl.g_BasicInput.get소매_거래선수_직영점() > 0)
            {
                arrvalue[i++] = (소매_수익_소계 - 소매_비용_소계) / CDataControl.g_BasicInput.get소매_거래선수_직영점();
            }
            else
            {
                arrvalue[i++] = 0;
            }

            return arrvalue;
        }

        public Int64[] getArrayOutput전체()
        {
            Int64[] arrvalue = new Int64[42];

            int i = 0;

            //******전체 수익
            arrvalue[i++] = 전체_수익_가입자수수료;
            arrvalue[i++] = 전체_수익_CS관리수수료;
            arrvalue[i++] = 전체_수익_업무취급수수료;
            arrvalue[i++] = 전체_수익_사업자모델매입에따른추가수익;
            arrvalue[i++] = 전체_수익_유통모델매입에따른추가수익_현금_Volume;
            arrvalue[i++] = 전체_수익_직영매장판매수익;
            arrvalue[i++] = 전체_수익_소계;

            //******전체 비용
            arrvalue[i++] = 전체_비용_대리점투자비용;
            arrvalue[i++] = 전체_비용_인건비_급여_복리후생비;
            arrvalue[i++] = 전체_비용_임차료;
            arrvalue[i++] = 전체_비용_이자비용;
            arrvalue[i++] = 전체_비용_부가세;
            arrvalue[i++] = 전체_비용_법인세;
            arrvalue[i++] = 전체_비용_기타판매관리비;
            arrvalue[i++] = 전체_비용_소계;
            arrvalue[i++] = 전체손익계;

            //******도매 수익
            arrvalue[i++] = 도매_수익_가입자관리수수료;
            arrvalue[i++] = 도매_수익_CS관리수수료;
            arrvalue[i++] = 도매_수익_사업자모델매입에따른추가수익;
            arrvalue[i++] = 도매_수익_유통모델매입에따른추가수익_현금_Volume;
            arrvalue[i++] = 도매_수익_소계;

            //******도매 비용
            arrvalue[i++] = 도매_비용_대리점투자비용;
            arrvalue[i++] = 도매_비용_인건비_급여_복리후생비;
            arrvalue[i++] = 도매_비용_임차료;
            arrvalue[i++] = 도매_비용_이자비용;
            arrvalue[i++] = 도매_비용_부가세;
            arrvalue[i++] = 도매_비용_법인세;
            arrvalue[i++] = 도매_비용_기타판매관리비;
            arrvalue[i++] = 도매_비용_소계;
            arrvalue[i++] = 도매손익계;

            //******소매 수익
            arrvalue[i++] = 소매_수익_월단위업무취급수수료;
            arrvalue[i++] = 소매_수익_직영매장판매수익;
            arrvalue[i++] = 소매_수익_소계;

            //******소매 비용
            arrvalue[i++] = 소매_비용_인건비_급여_복리후생비;
            arrvalue[i++] = 소매_비용_임차료;
            arrvalue[i++] = 소매_비용_이자비용;
            arrvalue[i++] = 소매_비용_부가세;
            arrvalue[i++] = 소매_비용_법인세;
            arrvalue[i++] = 소매_비용_기타판매관리비;
            arrvalue[i++] = 소매_비용_소계;
            arrvalue[i++] = 소매손익계;

            arrvalue[i++] = 점별손익추정;

            return arrvalue;
        }


        public void setArrayOutput전체(Int64[]  arrvalue)
        {
            int i = 0;

            //******전체 수익
            전체_수익_가입자수수료= arrvalue[i++];
            전체_수익_CS관리수수료= arrvalue[i++];
            전체_수익_업무취급수수료= arrvalue[i++];
            전체_수익_사업자모델매입에따른추가수익= arrvalue[i++];
            전체_수익_유통모델매입에따른추가수익_현금_Volume= arrvalue[i++];
            전체_수익_직영매장판매수익= arrvalue[i++];
            전체_수익_소계= arrvalue[i++];

            //******전체 비용
            전체_비용_대리점투자비용= arrvalue[i++];
            전체_비용_인건비_급여_복리후생비= arrvalue[i++];
            전체_비용_임차료= arrvalue[i++];
            전체_비용_이자비용= arrvalue[i++];
            전체_비용_부가세= arrvalue[i++];
            전체_비용_법인세= arrvalue[i++];
            전체_비용_기타판매관리비= arrvalue[i++];
            전체_비용_소계= arrvalue[i++];
            전체손익계= arrvalue[i++];

            //******도매 수익
            도매_수익_가입자관리수수료= arrvalue[i++];
            도매_수익_CS관리수수료= arrvalue[i++];
            도매_수익_사업자모델매입에따른추가수익= arrvalue[i++];
            도매_수익_유통모델매입에따른추가수익_현금_Volume= arrvalue[i++];
            도매_수익_소계= arrvalue[i++];

            //******도매 비용
            도매_비용_대리점투자비용= arrvalue[i++];
            도매_비용_인건비_급여_복리후생비= arrvalue[i++];
            도매_비용_임차료= arrvalue[i++];
            도매_비용_이자비용= arrvalue[i++];
            도매_비용_부가세= arrvalue[i++];
            도매_비용_법인세= arrvalue[i++];
            도매_비용_기타판매관리비= arrvalue[i++];
            도매_비용_소계= arrvalue[i++];
            도매손익계= arrvalue[i++];

            //******소매 수익
            소매_수익_월단위업무취급수수료= arrvalue[i++];
            소매_수익_직영매장판매수익= arrvalue[i++];
            소매_수익_소계= arrvalue[i++];

            //******소매 비용
            소매_비용_인건비_급여_복리후생비= arrvalue[i++];
            소매_비용_임차료= arrvalue[i++];
            소매_비용_이자비용= arrvalue[i++];
            소매_비용_부가세= arrvalue[i++];
            소매_비용_법인세= arrvalue[i++];
            소매_비용_기타판매관리비= arrvalue[i++];
            소매_비용_소계= arrvalue[i++];
            소매손익계= arrvalue[i++];

            점별손익추정= arrvalue[i++];
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



        private Int64 _전체_수익_가입자수수료;
        private Int64 _전체_수익_CS관리수수료;
        private Int64 _전체_수익_업무취급수수료;
        private Int64 _전체_수익_사업자모델매입에따른추가수익;
        private Int64 _전체_수익_유통모델매입에따른추가수익_현금_Volume;
        private Int64 _전체_수익_직영매장판매수익;

        
        private Int64 _전체_수익_소계;
        private Int64 _전체_비용_소계;
        private Int64 _전체손익계;
        private Int64 _도매_수익_소계;
        private Int64 _도매손익계;
        private Int64 _도매_비용_소계;
        private Int64 _소매_수익_소계;
        private Int64 _소매_비용_소계;
        private Int64 _소매손익계;
        private Int64 _점별손익추정;
        private Int64 _소매_수익_업무취급수수료;
        private Int64 _소매_수익_월단위업무취급수수료;



        public Int64 소매_수익_월단위업무취급수수료
        {
            get { return _소매_수익_월단위업무취급수수료; }
            set { _소매_수익_월단위업무취급수수료 = value; }
        }

        public Int64 전체_수익_가입자수수료
        {
            get { return _전체_수익_가입자수수료; }
            set { _전체_수익_가입자수수료 = value; }
        }

        public Int64 전체_수익_CS관리수수료
        {
            get { return _전체_수익_CS관리수수료; }
            set { _전체_수익_CS관리수수료 = value; }
        }

        public Int64 전체_수익_업무취급수수료
        {
            get { return _전체_수익_업무취급수수료; }
            set { _전체_수익_업무취급수수료 = value; }
        }

        public Int64 전체_수익_사업자모델매입에따른추가수익
        {
            get { return _전체_수익_사업자모델매입에따른추가수익; }
            set { _전체_수익_사업자모델매입에따른추가수익 = value; }
        }

        public Int64 전체_수익_유통모델매입에따른추가수익_현금_Volume
        {
            get { return _전체_수익_유통모델매입에따른추가수익_현금_Volume; }
            set { _전체_수익_유통모델매입에따른추가수익_현금_Volume = value; }
        }

        public Int64 전체_수익_직영매장판매수익
        {
            get { return _전체_수익_직영매장판매수익; }
            set { _전체_수익_직영매장판매수익 = value; }
        }

        public Int64 전체_수익_소계
        {
            get { return _전체_수익_소계; }
            set { _전체_수익_소계 = value; }
        }

        public Int64 전체_비용_소계
        {
            get { return _전체_비용_소계; }
            set { _전체_비용_소계 = value; }
        }

        public Int64 전체손익계
        {
            get { return _전체손익계; }
            set { _전체손익계 = value; }
        }

        public Int64 도매_수익_소계
        {
            get { return _도매_수익_소계; }
            set { _도매_수익_소계 = value; }
        }

        public Int64 도매_비용_소계
        {
            get { return _도매_비용_소계; }
            set { _도매_비용_소계 = value; }
        }

        public Int64 도매손익계
        {
            get { return _도매손익계; }
            set { _도매손익계 = value; }
        }

        public Int64 소매_수익_소계
        {
            get { return _소매_수익_소계; }
            set { _소매_수익_소계 = value; }
        }

        public Int64 소매_비용_소계
        {
            get { return _소매_비용_소계; }
            set { _소매_비용_소계 = value; }
        }

        public Int64 소매손익계
        {
            get { return _소매손익계; }
            set { _소매손익계 = value; }
        }

        public Int64 점별손익추정
        {
            get { return _점별손익추정; }
            set { _점별손익추정 = value; }
        }
    }
}
