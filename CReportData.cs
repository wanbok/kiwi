using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KIWI
{
    public class CReportData
    {
        //지역
        private String 지역 = "";
        public void set지역(String value)
        {
            지역 = value;
        }
        public String get지역()
        {
            return 지역;
        }

        //대리점
        private String 대리점 = "";
        public void set대리점(String value)
        {
            대리점 = value;
        }
        public String get대리점()
        {
            return 대리점;
        }

        //판매자
        private String 판매자 = "";
        public void set판매자(String value)
        {
            판매자 = value;
        }
        public String get판매자()
        {
            return 판매자;
        }

        //배경_및_이슈
        private String 배경_및_이슈 = "";
        public void set배경_및_이슈(String value)
        {
            배경_및_이슈 = value;
        }
        public String get배경_및_이슈()
        {
            return 배경_및_이슈;
        }

        //분석내용_및_대리점_활동방향
        private String 분석내용_및_대리점_활동방향 = "";
        public void set분석내용_및_대리점_활동방향(String value)
        {
            분석내용_및_대리점_활동방향 = value;
        }
        public String get분석내용_및_대리점_활동방향()
        {
            return 분석내용_및_대리점_활동방향;
        }

        //LG_지원_활동
        private String LG_지원_활동 = "";
        public void setLG_지원_활동(String value)
        {
            LG_지원_활동 = value;
        }
        public String getLG_지원_활동()
        {
            return LG_지원_활동;
        }

        public void setArrData(String[] data)
        {
            int i = 0;
            지역 = data[i++];
            대리점 = data[i++];
            판매자 = data[i++];
            배경_및_이슈 = data[i++];
            분석내용_및_대리점_활동방향 = data[i++];
            LG_지원_활동 = data[i++];
        }

        public String[] getArrData()
        {
            String[] data = new String[6];
            int i = 0;
            data[i++] = 지역;
            data[i++] = 대리점;
            data[i++] = 판매자;
            data[i++] = 배경_및_이슈;
            data[i++] = 분석내용_및_대리점_활동방향;
            data[i++] = LG_지원_활동;

            return data;
        }
    }
}
