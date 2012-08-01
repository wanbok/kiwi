using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KIWI
{
    class CAdminDataModel
    {
        private String 지역 = "";
        public void set지역(String value)
        {
            지역 = value;
        }
        public String get지역()
        {
            return 지역;
        }
        private String 대리점명 = "";
        public void set대리점명(String value)
        {
            대리점명 = value;
        }
        public String get대리점명()
        {
            return 대리점명;
        }
        private String 마케터 = "";
        public void set마케터(String value)
        {
            마케터 = value;
        }
        public String get마케터()
        {
            return 마케터;
        }
        private String 단위당손익 = "";
        public void set단위당손익(String value)
        {
            단위당손익 = value;
        }
        public String get단위당손익()
        {
            return 단위당손익;
        }
        private String 월capa = "";
        public void set월capa(String value)
        {
            월capa = value;
        }
        public String get월capa()
        {
            return 월capa;
        }
        private String 가입자수 = "";
        public void set가입자수(String value)
        {
            가입자수 = value;
        }
        public String get가입자수()
        {
            return 가입자수;
        }
        private String 직영점판매수익 = "";
        public void set직영점판매수익(String value)
        {
            직영점판매수익 = value;
        }
        public String get직영점판매수익()
        {
            return 직영점판매수익;
        }
        private String 선택여부 = "N";
        public void set선택여부(String value)
        {
            선택여부 = value;
        }
        public String get선택여부()
        {
            return 선택여부;
        }
        private String mExcelFileName = "";
        public void setmExcelFileName(String value)
        {
            mExcelFileName = value;
        }
        public String getmExcelFileName()
        {
            return mExcelFileName;
        }
        private CBasicInput mBI = new CBasicInput();
        public void setmBI(CBasicInput value)
        {
            mBI = value;
        }
        public CBasicInput getmBI()
        {
            return mBI;
        }
        private CBusinessData mDI = new CBusinessData();
        public void setmDI(CBusinessData value)
        {
            mDI = value;
        }
        public CBusinessData getmDI()
        {
            return mDI;
        }
        private CResultData mRD = new CResultData();
        public void setmRD(CResultData value)
        {
            mRD = value;
        }
        public CResultData getmRD()
        {
            return mRD;
        }
    }
}
