using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KIWI
{
    public class CReportData
    {
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
    }
}
