using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;

namespace KIWI
{
    class CAdminDataController
    {
        private String mXmlFileName = "files/resultdatalist.xml";//xml 파일명

        XmlDataDocument mXmlDoc = null;
        public CAdminDataController()
        {
            mXmlDoc = new XmlDataDocument();
            // If the directory doesn't exist, create it.
            if (!Directory.Exists("files"))
            {
                Directory.CreateDirectory("files");
            }
            FileInfo fi = new FileInfo(mXmlFileName);
            if (fi.Exists)
            {
                mXmlDoc.Load(mXmlFileName);
            }
        }

        public Int32 getDataLength()
        {
            if (mXmlDoc != null)
            {
                XmlNode nodeItems = mXmlDoc.FirstChild;
                if (nodeItems == null || nodeItems.ChildNodes == null) return 0;
                return nodeItems.ChildNodes.Count;
            }
            return 0;
        }

        public void setFileName(String FileName)
        {
            mXmlFileName = FileName;
        }

        private String getKey()
        {
            DateTime nowdatetime = DateTime.Now;
            String key = nowdatetime.ToString("yyyyMMDDhhmmss");
            return key;
        }

        //OpInspsr 정보를 파일로 저장한다
        public void AddSaveData(String 지역, String 대리점명, String 마케터, String 단위당손익, String 월capa, String 가입자수,
            String 직영점판매수익, String 선택여부, String mExcelFileName,
            CBasicInput mBI, CBusinessData mDI, CResultData mRD)
        {
            XmlNode nodeItems = null;

            //int nkey = 0;

            if (mXmlDoc.FirstChild == null)
            {
                //XmlDocument xmlDoc = new XmlDocument();

                nodeItems = mXmlDoc.CreateNode("element", "Items", "");
                //nkey = 0;

                //XmlAttribute xmlAtt = mXmlDoc.CreateAttribute("key");
                //xmlAtt.Value = "0";                
            }
            else
            {
                nodeItems = mXmlDoc.FirstChild;
            }

            XmlNode nodeItem = mXmlDoc.CreateNode("element", "Item", "");

            XmlNode node0 = mXmlDoc.CreateNode("element", "key", "");
            node0.InnerText = getKey();

            //리스트 기본 데이터
            XmlNode node1 = mXmlDoc.CreateNode("element", "지역", "");
            node1.InnerText = 지역;
            XmlNode node2 = mXmlDoc.CreateNode("element", "대리점명", "");
            node2.InnerText = 대리점명;
            XmlNode node3 = mXmlDoc.CreateNode("element", "마케터", "");
            node3.InnerText = 마케터;
            XmlNode node4 = mXmlDoc.CreateNode("element", "단위당손익", "");
            node4.InnerText = 단위당손익;
            XmlNode node5 = mXmlDoc.CreateNode("element", "월capa", "");
            node5.InnerText = 월capa;
            XmlNode node6 = mXmlDoc.CreateNode("element", "가입자수", "");
            node6.InnerText = 가입자수;
            XmlNode node7 = mXmlDoc.CreateNode("element", "직영점판매수익", "");
            node7.InnerText = 직영점판매수익;
            XmlNode node8 = mXmlDoc.CreateNode("element", "선택여부", "");
            node8.InnerText = 선택여부;
            XmlNode node9 = mXmlDoc.CreateNode("element", "엑셀파일명", "");
            node9.InnerText = mExcelFileName;


            //기본입력 데이터
            XmlNode node10 = mXmlDoc.CreateNode("element", "기본입력_도매_누적가입자수", "");
            node10.InnerText = mBI.getstr도매_누적가입자수();
            XmlNode node11 = mXmlDoc.CreateNode("element", "기본입력_도매_월평균판매대수_신규", "");
            node11.InnerText = mBI.getstr도매_월평균판매대수_신규();
            XmlNode node12 = mXmlDoc.CreateNode("element", "기본입력_도매_월평균판매대수_기변", "");
            node12.InnerText = mBI.getstr도매_월평균판매대수_기변();
            XmlNode node13 = mXmlDoc.CreateNode("element", "기본입력_도매_월평균유통모델출고대수_LG", "");
            node13.InnerText = mBI.getstr도매_월평균유통모델출고대수_LG();
            XmlNode node14 = mXmlDoc.CreateNode("element", "기본입력_도매_월평균유통모델출고대수_SS", "");
            node14.InnerText = mBI.getstr도매_월평균유통모델출고대수_SS();
            XmlNode node15 = mXmlDoc.CreateNode("element", "기본입력_도매_거래선수_개통사무실", "");
            node15.InnerText = mBI.getstr도매_거래선수_개통사무실();
            XmlNode node16 = mXmlDoc.CreateNode("element", "기본입력_도매_거래선수_판매점", "");
            node16.InnerText = mBI.getstr도매_거래선수_판매점();
            XmlNode node17 = mXmlDoc.CreateNode("element", "기본입력_도매_직원수_간부급", "");
            node17.InnerText = mBI.getstr도매_직원수_간부급();
            XmlNode node18 = mXmlDoc.CreateNode("element", "기본입력_도매_직원수_평사원", "");
            node18.InnerText = mBI.getstr도매_직원수_평사원();

            XmlNode node19 = mXmlDoc.CreateNode("element", "기본입력_소매_월평균판매대수_신규", "");
            node19.InnerText = mBI.getstr소매_월평균판매대수_신규();
            XmlNode node20 = mXmlDoc.CreateNode("element", "기본입력_소매_월평균판매대수_기변", "");
            node20.InnerText = mBI.getstr소매_월평균판매대수_기변();
            XmlNode node21 = mXmlDoc.CreateNode("element", "기본입력_소매_거래선수_직영점", "");
            node21.InnerText = mBI.getstr소매_거래선수_직영점();
            XmlNode node22 = mXmlDoc.CreateNode("element", "기본입력_소매_직원수_간부급", "");
            node22.InnerText = mBI.getstr소매_직원수_간부급();
            XmlNode node23 = mXmlDoc.CreateNode("element", "기본입력_소매_직원수_평사원", "");
            node23.InnerText = mBI.getstr소매_직원수_평사원();


            //상세입력 데이터
            XmlNode node24 = mXmlDoc.CreateNode("element", "상세입력_도매_수익_월평균관리수수료", "");
            node24.InnerText = mDI.getstr도매_수익_월평균관리수수료();
            XmlNode node25 = mXmlDoc.CreateNode("element", "상세입력_도매_수익_CS관리수수료", "");
            node25.InnerText = mDI.getstr도매_수익_CS관리수수료();
            XmlNode node26 = mXmlDoc.CreateNode("element", "상세입력_도매_수익_사업자모델매입관련추가수익", "");
            node26.InnerText = mDI.getstr도매_수익_사업자모델매입관련추가수익();
            XmlNode node27 = mXmlDoc.CreateNode("element", "상세입력_도매_수익_유통모델매입관련추가수익_현금DC", "");
            node27.InnerText = mDI.getstr도매_수익_유통모델매입관련추가수익_현금DC();
            XmlNode node28 = mXmlDoc.CreateNode("element", "상세입력_도매_수익_유통모델매입관련추가수익_VolumeDC", "");
            node28.InnerText = mDI.getstr도매_수익_유통모델매입관련추가수익_VolumeDC();

            XmlNode node29 = mXmlDoc.CreateNode("element", "상세입력_도매_비용_대리점투자금액_신규", "");
            node29.InnerText = mDI.getstr도매_비용_대리점투자금액_신규();
            XmlNode node30 = mXmlDoc.CreateNode("element", "상세입력_도매_비용_대리점투자금액_기변", "");
            node30.InnerText = mDI.getstr도매_비용_대리점투자금액_기변();
            XmlNode node31 = mXmlDoc.CreateNode("element", "상세입력_도매_비용_직원급여_간부급", "");
            node31.InnerText = mDI.getstr도매_비용_직원급여_간부급();
            XmlNode node32 = mXmlDoc.CreateNode("element", "상세입력_도매_비용_직원급여_평사원", "");
            node32.InnerText = mDI.getstr도매_비용_직원급여_평사원();
            XmlNode node33 = mXmlDoc.CreateNode("element", "상세입력_도매_비용_지급임차료", "");
            node33.InnerText = mDI.getstr도매_비용_지급임차료();
            XmlNode node34 = mXmlDoc.CreateNode("element", "상세입력_도매_비용_운반비", "");
            node34.InnerText = mDI.getstr도매_비용_운반비();
            XmlNode node35 = mXmlDoc.CreateNode("element", "상세입력_도매_비용_차량유지비", "");
            node35.InnerText = mDI.getstr도매_비용_차량유지비();
            XmlNode node36 = mXmlDoc.CreateNode("element", "상세입력_도매_비용_지급수수료", "");
            node36.InnerText = mDI.getstr도매_비용_지급수수료();
            XmlNode node37 = mXmlDoc.CreateNode("element", "상세입력_도매_비용_판매촉진비", "");
            node37.InnerText = mDI.getstr도매_비용_판매촉진비();
            XmlNode node38 = mXmlDoc.CreateNode("element", "상세입력_도매_비용_건물관리비", "");
            node38.InnerText = mDI.getstr도매_비용_건물관리비();

            XmlNode node39 = mXmlDoc.CreateNode("element", "상세입력_소매_수익_월평균업무취급수수료", "");
            node39.InnerText = mDI.getstr소매_수익_월평균업무취급수수료();
            XmlNode node40 = mXmlDoc.CreateNode("element", "상세입력_소매_수익_직영매장판매수익", "");
            node40.InnerText = mDI.getstr소매_수익_직영매장판매수익();

            XmlNode node41 = mXmlDoc.CreateNode("element", "상세입력_소매_비용_직원급여_간부급", "");
            node41.InnerText = mDI.getstr소매_비용_직원급여_간부급();
            XmlNode node42 = mXmlDoc.CreateNode("element", "상세입력_소매_비용_직원급여_평사원", "");
            node42.InnerText = mDI.getstr소매_비용_직원급여_평사원();
            XmlNode node43 = mXmlDoc.CreateNode("element", "상세입력_소매_비용_지급임차료", "");
            node43.InnerText = mDI.getstr소매_비용_지급임차료();
            XmlNode node44 = mXmlDoc.CreateNode("element", "상세입력_소매_비용_지급수수료", "");
            node44.InnerText = mDI.getstr소매_비용_지급수수료();
            XmlNode node45 = mXmlDoc.CreateNode("element", "상세입력_소매_비용_판매촉진비", "");
            node45.InnerText = mDI.getstr소매_비용_판매촉진비();
            XmlNode node46 = mXmlDoc.CreateNode("element", "상세입력_소매_비용_건물관리비", "");
            node46.InnerText = mDI.getstr소매_비용_건물관리비();

            XmlNode node47 = mXmlDoc.CreateNode("element", "상세입력_도소매_비용_복리후생비", "");
            node47.InnerText = mDI.getstr도소매_비용_복리후생비();
            XmlNode node48 = mXmlDoc.CreateNode("element", "상세입력_도소매_비용_통신비", "");
            node48.InnerText = mDI.getstr도소매_비용_통신비();
            XmlNode node49 = mXmlDoc.CreateNode("element", "상세입력_도소매_비용_공과금", "");
            node49.InnerText = mDI.getstr도소매_비용_공과금();
            XmlNode node50 = mXmlDoc.CreateNode("element", "상세입력_도소매_비용_소모품비", "");
            node50.InnerText = mDI.getstr도소매_비용_소모품비();
            XmlNode node51 = mXmlDoc.CreateNode("element", "상세입력_도소매_비용_이자비용", "");
            node51.InnerText = mDI.getstr도소매_비용_이자비용();
            XmlNode node52 = mXmlDoc.CreateNode("element", "상세입력_도소매_비용_부가세", "");
            node52.InnerText = mDI.getstr도소매_비용_부가세();
            XmlNode node53 = mXmlDoc.CreateNode("element", "상세입력_도소매_비용_법인세", "");
            node53.InnerText = mDI.getstr도소매_비용_법인세();
            XmlNode node54 = mXmlDoc.CreateNode("element", "상세입력_도소매_비용_기타", "");
            node54.InnerText = mDI.getstr도소매_비용_기타();


            //결과 단위당 데이터
            XmlNode node55 = mXmlDoc.CreateNode("element", "결과_도매_수익_가입자관리수수료", "");
            node55.InnerText = mRD.getstr도매_수익_가입자관리수수료();
            XmlNode node56 = mXmlDoc.CreateNode("element", "결과_도매_수익_CS관리수수료", "");
            node56.InnerText = mRD.getstr도매_수익_CS관리수수료();
            XmlNode node57 = mXmlDoc.CreateNode("element", "결과_소매_수익_업무취급수수료", "");
            node57.InnerText = mRD.getstr소매_수익_업무취급수수료();
            XmlNode node58 = mXmlDoc.CreateNode("element", "결과_도매_수익_사업자모델매입에따른추가수익", "");
            node58.InnerText = mRD.getstr도매_수익_사업자모델매입에따른추가수익();
            XmlNode node59 = mXmlDoc.CreateNode("element", "결과_도매_수익_유통모델매입에따른추가수익_현금_Volume", "");
            node59.InnerText = mRD.getstr도매_수익_유통모델매입에따른추가수익_현금_Volume();
            XmlNode node60 = mXmlDoc.CreateNode("element", "결과_소매_수익_직영매장판매수익", "");
            node60.InnerText = mRD.getstr소매_수익_직영매장판매수익();

            XmlNode node61 = mXmlDoc.CreateNode("element", "결과_전체_비용_대리점투자비용", "");
            node61.InnerText = mRD.getstr전체_비용_대리점투자비용();
            XmlNode node62 = mXmlDoc.CreateNode("element", "결과_전체_비용_인건비_급여_복리후생비", "");
            node62.InnerText = mRD.getstr전체_비용_인건비_급여_복리후생비();
            XmlNode node63 = mXmlDoc.CreateNode("element", "결과_전체_비용_임차료", "");
            node63.InnerText = mRD.getstr전체_비용_임차료();
            XmlNode node64 = mXmlDoc.CreateNode("element", "결과_전체_비용_이자비용", "");
            node64.InnerText = mRD.getstr전체_비용_이자비용();
            XmlNode node65 = mXmlDoc.CreateNode("element", "결과_전체_비용_부가세", "");
            node65.InnerText = mRD.getstr전체_비용_부가세();
            XmlNode node66 = mXmlDoc.CreateNode("element", "결과_전체_비용_법인세", "");
            node66.InnerText = mRD.getstr전체_비용_법인세();

            XmlNode node67 = mXmlDoc.CreateNode("element", "결과_도매_비용_대리점투자비용", "");
            node67.InnerText = mRD.getstr도매_비용_대리점투자비용();
            XmlNode node68 = mXmlDoc.CreateNode("element", "결과_도매_비용_인건비_급여_복리후생비", "");
            node68.InnerText = mRD.getstr도매_비용_인건비_급여_복리후생비();
            XmlNode node69 = mXmlDoc.CreateNode("element", "결과_도매_비용_임차료", "");
            node69.InnerText = mRD.getstr도매_비용_임차료();
            XmlNode node70 = mXmlDoc.CreateNode("element", "결과_도매_비용_이자비용", "");
            node70.InnerText = mRD.getstr도매_비용_이자비용();
            XmlNode node71 = mXmlDoc.CreateNode("element", "결과_도매_비용_부가세", "");
            node71.InnerText = mRD.getstr도매_비용_부가세();
            XmlNode node72 = mXmlDoc.CreateNode("element", "결과_도매_비용_법인세", "");
            node72.InnerText = mRD.getstr도매_비용_법인세();
            XmlNode node73 = mXmlDoc.CreateNode("element", "결과_도매_비용_기타판매관리비", "");
            node73.InnerText = mRD.getstr도매_비용_기타판매관리비();


            XmlNode node74 = mXmlDoc.CreateNode("element", "결과_소매_비용_인건비_급여_복리후생비", "");
            node74.InnerText = mRD.getstr소매_비용_인건비_급여_복리후생비();
            XmlNode node75 = mXmlDoc.CreateNode("element", "결과_소매_비용_임차료", "");
            node75.InnerText = mRD.getstr소매_비용_임차료();
            XmlNode node76 = mXmlDoc.CreateNode("element", "결과_소매_비용_이자비용", "");
            node76.InnerText = mRD.getstr소매_비용_이자비용();
            XmlNode node77 = mXmlDoc.CreateNode("element", "결과_소매_비용_부가세", "");
            node77.InnerText = mRD.getstr소매_비용_부가세();
            XmlNode node78 = mXmlDoc.CreateNode("element", "결과_소매_비용_법인세", "");
            node78.InnerText = mRD.getstr소매_비용_법인세();
            XmlNode node79 = mXmlDoc.CreateNode("element", "결과_소매_비용_기타판매관리비", "");
            node79.InnerText = mRD.getstr소매_비용_기타판매관리비();

            nodeItem.AppendChild(node0);
            nodeItem.AppendChild(node1);
            nodeItem.AppendChild(node2);
            nodeItem.AppendChild(node3);
            nodeItem.AppendChild(node4);
            nodeItem.AppendChild(node5);
            nodeItem.AppendChild(node6);
            nodeItem.AppendChild(node7);
            nodeItem.AppendChild(node8);
            nodeItem.AppendChild(node9);

            nodeItem.AppendChild(node10);
            nodeItem.AppendChild(node11);
            nodeItem.AppendChild(node12);
            nodeItem.AppendChild(node13);
            nodeItem.AppendChild(node14);
            nodeItem.AppendChild(node15);
            nodeItem.AppendChild(node16);
            nodeItem.AppendChild(node17);
            nodeItem.AppendChild(node18);
            nodeItem.AppendChild(node19);
            nodeItem.AppendChild(node20);
            nodeItem.AppendChild(node21);
            nodeItem.AppendChild(node22);
            nodeItem.AppendChild(node23);

            nodeItem.AppendChild(node24);
            nodeItem.AppendChild(node25);
            nodeItem.AppendChild(node26);
            nodeItem.AppendChild(node27);
            nodeItem.AppendChild(node28);
            nodeItem.AppendChild(node29);
            nodeItem.AppendChild(node30);

            nodeItem.AppendChild(node31);
            nodeItem.AppendChild(node32);
            nodeItem.AppendChild(node33);
            nodeItem.AppendChild(node34);
            nodeItem.AppendChild(node35);
            nodeItem.AppendChild(node36);
            nodeItem.AppendChild(node37);
            nodeItem.AppendChild(node38);
            nodeItem.AppendChild(node39);
            nodeItem.AppendChild(node40);

            nodeItem.AppendChild(node41);
            nodeItem.AppendChild(node42);
            nodeItem.AppendChild(node43);
            nodeItem.AppendChild(node44);
            nodeItem.AppendChild(node45);
            nodeItem.AppendChild(node46);
            nodeItem.AppendChild(node47);
            nodeItem.AppendChild(node48);
            nodeItem.AppendChild(node49);
            nodeItem.AppendChild(node50);

            nodeItem.AppendChild(node51);
            nodeItem.AppendChild(node52);
            nodeItem.AppendChild(node53);
            nodeItem.AppendChild(node54);
            nodeItem.AppendChild(node55);
            nodeItem.AppendChild(node56);
            nodeItem.AppendChild(node57);
            nodeItem.AppendChild(node58);
            nodeItem.AppendChild(node59);
            nodeItem.AppendChild(node60);

            nodeItem.AppendChild(node61);
            nodeItem.AppendChild(node62);
            nodeItem.AppendChild(node63);
            nodeItem.AppendChild(node64);
            nodeItem.AppendChild(node65);
            nodeItem.AppendChild(node66);
            nodeItem.AppendChild(node67);
            nodeItem.AppendChild(node68);
            nodeItem.AppendChild(node69);
            nodeItem.AppendChild(node70);

            nodeItem.AppendChild(node71);
            nodeItem.AppendChild(node72);
            nodeItem.AppendChild(node73);
            nodeItem.AppendChild(node74);
            nodeItem.AppendChild(node75);
            nodeItem.AppendChild(node76);
            nodeItem.AppendChild(node77);
            nodeItem.AppendChild(node78);
            nodeItem.AppendChild(node79);

            nodeItems.AppendChild(nodeItem);

            mXmlDoc.AppendChild(nodeItems);
            mXmlDoc.Save(mXmlFileName);
        }

        private int findIndexByKey(String key)
        {
            Boolean isFind = false;

            if (mXmlDoc != null)
            {
                int index = 0;

                XmlNode nodeItems = mXmlDoc.FirstChild;
                foreach (XmlNode nodeItem in nodeItems)//item
                {
                    index++;
                    foreach (XmlNode nodeItem2 in nodeItem)//item 내부
                    {
                        if ((nodeItem2.Name == "key") && (nodeItem2.InnerText == key))
                        {
                            isFind = true;
                            break;
                        }
                    }

                    if (isFind)
                    {

                        return index;
                    }
                }
            }

            return -1;
        }
        public void GetData(int index, out String key, out String 지역, out String 대리점명, out String 마케터, out String 단위당손익, out String 월capa, out String 가입자수,
            out String 직영점판매수익, out String 선택여부, out String mExcelFileName,
            out CBasicInput mBI, out CBusinessData mDI, out CResultData mRD)
        {
            key = "";
            지역 = "";
            대리점명 = "";
            마케터 = "";
            단위당손익 = "";
            월capa = "";
            가입자수 = "";
            직영점판매수익 = "";
            선택여부 = "N";
            mExcelFileName = "";
            mBI = new CBasicInput();
            mDI = new CBusinessData();
            mRD = new CResultData();

            if (mXmlDoc != null)
            {
                XmlNode nodeItems = mXmlDoc.FirstChild;
                XmlNode nodeItem = nodeItems.ChildNodes.Item(index);

                foreach (XmlNode nodeItem2 in nodeItem)
                {
                    switch (nodeItem2.Name)
                    {
                        case "key":
                            key = nodeItem2.InnerText;
                            break;

                        case "지역":
                            지역 = nodeItem2.InnerText;
                            break;
                        case "대리점명":
                            대리점명 = nodeItem2.InnerText;
                            break;
                        case "마케터":
                            마케터 = nodeItem2.InnerText;
                            break;
                        case "단위당손익":
                            단위당손익 = nodeItem2.InnerText;
                            break;
                        case "월capa":
                            월capa = nodeItem2.InnerText;
                            break;
                        case "가입자수":
                            가입자수 = nodeItem2.InnerText;
                            break;
                        case "직영점판매수익":
                            직영점판매수익 = nodeItem2.InnerText;
                            break;
                        case "선택여부":
                            선택여부 = nodeItem2.InnerText;
                            break;
                        case "mExcelFileName":
                            mExcelFileName = nodeItem2.InnerText;
                            break;


                        //기본입력 데이터
                        case "기본입력_도매_누적가입자수":
                            mBI.set도매_누적가입자수(nodeItem2.InnerText);
                            break;
                        case "기본입력_도매_월평균판매대수_신규":
                            mBI.set도매_월평균판매대수_신규(nodeItem2.InnerText);
                            break;
                        case "기본입력_도매_월평균판매대수_기변":
                            mBI.set도매_월평균판매대수_기변(nodeItem2.InnerText);
                            break;
                        case "기본입력_도매_월평균유통모델출고대수_LG":
                            mBI.set도매_월평균유통모델출고대수_LG(nodeItem2.InnerText);
                            break;
                        case "기본입력_도매_월평균유통모델출고대수_SS":
                            mBI.set도매_월평균유통모델출고대수_SS(nodeItem2.InnerText);
                            break;
                        case "기본입력_도매_거래선수_개통사무실":
                            mBI.set도매_거래선수_개통사무실(nodeItem2.InnerText);
                            break;
                        case "기본입력_도매_거래선수_판매점":
                            mBI.set도매_거래선수_판매점(nodeItem2.InnerText);
                            break;
                        case "기본입력_도매_직원수_간부급":
                            mBI.set도매_직원수_간부급(nodeItem2.InnerText);
                            break;
                        case "기본입력_도매_직원수_평사원":
                            mBI.set도매_직원수_평사원(nodeItem2.InnerText);
                            break;

                        case "기본입력_소매_월평균판매대수_신규":
                            mBI.set소매_월평균판매대수_신규(nodeItem2.InnerText);
                            break;
                        case "기본입력_소매_월평균판매대수_기변":
                            mBI.set소매_월평균판매대수_기변(nodeItem2.InnerText);
                            break;
                        case "기본입력_소매_거래선수_직영점":
                            mBI.set소매_거래선수_직영점(nodeItem2.InnerText);
                            break;
                        case "기본입력_소매_직원수_간부급":
                            mBI.set소매_직원수_간부급(nodeItem2.InnerText);
                            break;
                        case "기본입력_소매_직원수_평사원":
                            mBI.set소매_직원수_평사원(nodeItem2.InnerText);
                            break;

                        //상세입력 데이터

                        case "상세입력_도매_수익_월평균관리수수료":
                            mDI.set도매_수익_월평균관리수수료(nodeItem2.InnerText);
                            break;
                        case "상세입력_도매_수익_CS관리수수료":
                            mDI.set도매_수익_CS관리수수료(nodeItem2.InnerText);
                            break;
                        case "상세입력_도매_수익_사업자모델매입관련추가수익":
                            mDI.set도매_수익_사업자모델매입관련추가수익(nodeItem2.InnerText);
                            break;
                        case "상세입력_도매_수익_유통모델매입관련추가수익_현금DC":
                            mDI.set도매_수익_유통모델매입관련추가수익_현금DC(nodeItem2.InnerText);
                            break;
                        case "상세입력_도매_수익_유통모델매입관련추가수익_VolumeDC":
                            mDI.set도매_수익_유통모델매입관련추가수익_VolumeDC(nodeItem2.InnerText);
                            break;
                        case "상세입력_도매_비용_대리점투자금액_신규":
                            mDI.set도매_비용_대리점투자금액_신규(nodeItem2.InnerText);
                            break;
                        case "상세입력_도매_비용_대리점투자금액_기변":
                            mDI.set도매_비용_대리점투자금액_기변(nodeItem2.InnerText);
                            break;
                        case "상세입력_도매_비용_직원급여_간부급":
                            mDI.set도매_비용_직원급여_간부급(nodeItem2.InnerText);
                            break;
                        case "상세입력_도매_비용_직원급여_평사원":
                            mDI.set도매_비용_직원급여_평사원(nodeItem2.InnerText);
                            break;
                        case "상세입력_도매_비용_지급임차료":
                            mDI.set도매_비용_지급임차료(nodeItem2.InnerText);
                            break;
                        case "상세입력_도매_비용_운반비":
                            mDI.set도매_비용_운반비(nodeItem2.InnerText);
                            break;
                        case "상세입력_도매_비용_차량유지비":
                            mDI.set도매_비용_차량유지비(nodeItem2.InnerText);
                            break;
                        case "상세입력_도매_비용_지급수수료":
                            mDI.set도매_비용_지급수수료(nodeItem2.InnerText);
                            break;
                        case "상세입력_도매_비용_판매촉진비":
                            mDI.set도매_비용_판매촉진비(nodeItem2.InnerText);
                            break;
                        case "상세입력_도매_비용_건물관리비":
                            mDI.set도매_비용_건물관리비(nodeItem2.InnerText);
                            break;
                        case "상세입력_소매_수익_월평균업무취급수수료":
                            mDI.set소매_수익_월평균업무취급수수료(nodeItem2.InnerText);
                            break;
                        case "상세입력_소매_수익_직영매장판매수익":
                            mDI.set소매_수익_직영매장판매수익(nodeItem2.InnerText);
                            break;

                        case "상세입력_소매_비용_직원급여_간부급":
                            mDI.set소매_비용_직원급여_간부급(nodeItem2.InnerText);
                            break;
                        case "상세입력_소매_비용_직원급여_평사원":
                            mDI.set소매_비용_직원급여_평사원(nodeItem2.InnerText);
                            break;
                        case "상세입력_소매_비용_지급임차료":
                            mDI.set소매_비용_지급임차료(nodeItem2.InnerText);
                            break;
                        case "상세입력_소매_비용_지급수수료":
                            mDI.set소매_비용_지급수수료(nodeItem2.InnerText);
                            break;
                        case "상세입력_소매_비용_판매촉진비":
                            mDI.set소매_비용_판매촉진비(nodeItem2.InnerText);
                            break;
                        case "상세입력_소매_비용_건물관리비":
                            mDI.set소매_비용_건물관리비(nodeItem2.InnerText);
                            break;
                        case "상세입력_도소매_비용_복리후생비":
                            mDI.set도소매_비용_복리후생비(nodeItem2.InnerText);
                            break;
                        case "상세입력_도소매_비용_통신비":
                            mDI.set도소매_비용_통신비(nodeItem2.InnerText);
                            break;
                        case "상세입력_도소매_비용_공과금":
                            mDI.set도소매_비용_공과금(nodeItem2.InnerText);
                            break;
                        case "상세입력_도소매_비용_소모품비":
                            mDI.set도소매_비용_소모품비(nodeItem2.InnerText);
                            break;
                        case "상세입력_도소매_비용_이자비용":
                            mDI.set도소매_비용_이자비용(nodeItem2.InnerText);
                            break;
                        case "상세입력_도소매_비용_부가세":
                            mDI.set도소매_비용_부가세(nodeItem2.InnerText);
                            break;
                        case "상세입력_도소매_비용_법인세":
                            mDI.set도소매_비용_법인세(nodeItem2.InnerText);
                            break;
                        case "상세입력_도소매_비용_기타":
                            mDI.set도소매_비용_기타(nodeItem2.InnerText);
                            break;


                        //결과 단위당 데이터
                        case "결과_도매_수익_가입자관리수수료":
                            mRD.set도매_수익_가입자관리수수료(nodeItem2.InnerText);
                            break;
                        case "결과_도매_수익_CS관리수수료":
                            mRD.set도매_수익_CS관리수수료(nodeItem2.InnerText);
                            break;
                        case "결과_소매_수익_업무취급수수료":
                            mRD.set소매_수익_업무취급수수료(nodeItem2.InnerText);
                            break;
                        case "결과_도매_수익_사업자모델매입에따른추가수익":
                            mRD.set도매_수익_사업자모델매입에따른추가수익(nodeItem2.InnerText);
                            break;
                        case "결과_도매_수익_유통모델매입에따른추가수익_현금_Volume":
                            mRD.set도매_수익_유통모델매입에따른추가수익_현금_Volume(nodeItem2.InnerText);
                            break;
                        case "결과_소매_수익_직영매장판매수익":
                            mRD.set소매_수익_직영매장판매수익(nodeItem2.InnerText);
                            break;


                        case "결과_전체_비용_대리점투자비용":
                            mRD.set전체_비용_대리점투자비용(nodeItem2.InnerText);
                            break;
                        case "결과_전체_비용_인건비_급여_복리후생비":
                            mRD.set전체_비용_인건비_급여_복리후생비(nodeItem2.InnerText);
                            break;
                        case "결과_전체_비용_임차료":
                            mRD.set전체_비용_임차료(nodeItem2.InnerText);
                            break;
                        case "결과_전체_비용_이자비용":
                            mRD.set전체_비용_이자비용(nodeItem2.InnerText);
                            break;
                        case "결과_전체_비용_부가세":
                            mRD.set전체_비용_부가세(nodeItem2.InnerText);
                            break;
                        case "결과_전체_비용_법인세":
                            mRD.set전체_비용_법인세(nodeItem2.InnerText);
                            break;
                        case "결과_도매_비용_대리점투자비용":
                            mRD.set도매_비용_대리점투자비용(nodeItem2.InnerText);
                            break;
                        case "결과_도매_비용_인건비_급여_복리후생비":
                            mRD.set도매_비용_인건비_급여_복리후생비(nodeItem2.InnerText);
                            break;
                        case "결과_도매_비용_임차료":
                            mRD.set도매_비용_임차료(nodeItem2.InnerText);
                            break;
                        case "결과_도매_비용_이자비용":
                            mRD.set도매_비용_이자비용(nodeItem2.InnerText);
                            break;
                        case "결과_도매_비용_부가세":
                            mRD.set도매_비용_부가세(nodeItem2.InnerText);
                            break;
                        case "결과_도매_비용_법인세":
                            mRD.set도매_비용_법인세(nodeItem2.InnerText);
                            break;
                        case "결과_도매_비용_기타판매관리비":
                            mRD.set도매_비용_기타판매관리비(nodeItem2.InnerText);
                            break;

                        case "결과_소매_비용_인건비_급여_복리후생비":
                            mRD.set소매_비용_인건비_급여_복리후생비(nodeItem2.InnerText);
                            break;
                        case "결과_소매_비용_임차료":
                            mRD.set소매_비용_임차료(nodeItem2.InnerText);
                            break;
                        case "결과_소매_비용_이자비용":
                            mRD.set소매_비용_이자비용(nodeItem2.InnerText);
                            break;
                        case "결과_소매_비용_부가세":
                            mRD.set소매_비용_부가세(nodeItem2.InnerText);
                            break;
                        case "결과_소매_비용_법인세":
                            mRD.set소매_비용_법인세(nodeItem2.InnerText);
                            break;
                        case "결과_소매_비용_기타판매관리비":
                            mRD.set소매_비용_기타판매관리비(nodeItem2.InnerText);
                            break;

                    }
                }


            }
        }

        public Boolean toggle선택여부(String key)
        {
            Boolean isFind = false;

            if (mXmlDoc != null)
            {
                int index = 0;

                XmlNode nodeItems = mXmlDoc.FirstChild;
                foreach (XmlNode nodeItem in nodeItems)//item
                {
                    index++;
                    foreach (XmlNode nodeItem2 in nodeItem)//item 내부
                    {
                        if ((nodeItem2.Name == "key") && (nodeItem2.InnerText == key))
                        {
                            isFind = true;
                            break;
                        }
                    }

                    if (isFind)
                    {
                        foreach (XmlNode nodeItem2 in nodeItem)//item 내부
                        {
                            if (nodeItem2.Name == "선택여부")
                            {
                                if (nodeItem2.InnerText == "Y")
                                {
                                    nodeItem2.InnerText = "N";
                                }
                                else
                                {
                                    nodeItem2.InnerText = "Y";
                                }
                            }
                        }
                        mXmlDoc.Save(mXmlFileName);
                        return true;
                    }
                }
            }

            return false;
        }

        public Boolean deleteData(String key)
        {
            Boolean isFind = false;

            if (mXmlDoc != null)
            {
                int index = 0;

                XmlNode nodeItems = mXmlDoc.FirstChild;
                foreach (XmlNode nodeItem in nodeItems)//item
                {
                    index++;
                    foreach (XmlNode nodeItem2 in nodeItem)//item 내부
                    {
                        if ((nodeItem2.Name == "key") && (nodeItem2.InnerText == key))
                        {
                            isFind = true;
                            break;
                        }
                    }

                    if (isFind)
                    {
                        nodeItems.RemoveChild(nodeItem);
                        mXmlDoc.Save(mXmlFileName);
                        return true;
                    }
                }
            }

            return false;
        }


    }



}
