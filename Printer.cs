using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using excel = Microsoft.Office.Interop.Excel;

namespace KIWI
{
    public class Printer
    {
        private DataSet1 ds = null;

        public Printer() {

            //프린트용 데이터 저장
            setPrintOut();

            FormReport report = new FormReport(ds);
            report.ShowDialog();
        }

        // 프린트용 데이터 저장
        private void setPrintOut()
        {
            ds = new DataSet1();

            DataTable[] businessTables = { ds.SimplizedResultAverageTotal, ds.SimplizedResultAverageWholesale, ds.SimplizedResultAverageRetail };
            DataTable[] storeTables = { ds.SimplizedResultThisTotal, ds.SimplizedResultThisWholesale, ds.SimplizedResultThisRetail };
            DataTable[] futureTables = { ds.SimplizedResultFutureTotal, ds.SimplizedResultFutureWholesale, ds.SimplizedResultFutureRetail };
            DataTable diffrenceForAnalysis = ds.DifferenceForAnalysis;
            DataTable baseData = ds.BaseData;
            DataTable name = ds.Name;
            DataTable comments = ds.Comments;

            DataTable[] simBusinessTables = { ds.SimplizedAverageTotal, ds.SimplizedAverageWholesale, ds.SimplizedAverageRetail };
            DataTable[] simStoreTables = { ds.SimplizedThisTotal, ds.SimplizedThisWholesale, ds.SimplizedThisRetail };
            DataTable[] simFutureTables = { ds.SimplizedFutureTotal, ds.SimplizedFutureWholesale, ds.SimplizedFutureRetail };

            // 본 데이터
            setDataTableForAnalysis(businessTables, CDataControl.g_ResultBusinessTotal, CDataControl.g_ResultBusiness, CDataControl.g_BusinessAvg);
            setDataTableForAnalysis(storeTables, CDataControl.g_ResultStoreTotal, CDataControl.g_ResultStore, CDataControl.g_DetailInput);
            setDataTableForAnalysis(futureTables, CDataControl.g_ResultFutureTotal, CDataControl.g_ResultFuture, CDataControl.g_DetailInput);

            // 시뮬레이션 데이터
            setDataTableForAnalysis(simBusinessTables, CDataControl.g_SimResultBusinessTotal, CDataControl.g_SimResultBusiness, CDataControl.g_BusinessAvg);
            setDataTableForAnalysis(simStoreTables, CDataControl.g_SimResultStoreTotal, CDataControl.g_SimResultStore, CDataControl.g_SimDetailInput);
            setDataTableForAnalysis(simFutureTables, CDataControl.g_SimResultFutureTotal, CDataControl.g_SimResultFuture, CDataControl.g_SimDetailInput);

            DataRow r = diffrenceForAnalysis.NewRow();
            for (int i = 0; i < 16; i++)
            {
                String result;
                Int64 all = CDataControl.g_ResultBusiness.getArr전체_리포트용(CDataControl.g_DetailInput)[i];
                Int64 agency = CDataControl.g_ResultStore.getArr전체_리포트용(CDataControl.g_BusinessAvg)[i];
                if (all < agency)
                {
                    result = "+";
                }
                else if (all > agency)
                {
                    result = "-";
                }
                else
                {
                    result = "=";
                }
                r[i] = result;
            }
            diffrenceForAnalysis.Rows.Add(r);

            r = baseData.NewRow();
            for (int i = 0; i < CDataControl.g_BasicInput.getArrData_리포트용().Length; i++)
            {
                r[i] = CDataControl.g_BasicInput.getArrData_리포트용()[i];
            }
            baseData.Rows.Add(r);

            r = name.NewRow();
            r[0] = CDataControl.g_ReportData.get대리점();
            r[1] = CDataControl.g_ReportData.get마케터();
            name.Rows.Add(r);

            r = comments.NewRow();
            r[0] = CDataControl.g_ReportData.get분석내용_및_대리점_활동방향();
            r[1] = CDataControl.g_ReportData.getLG_지원_활동();
            r[2] = CDataControl.g_ReportData.get배경_및_이슈();
            comments.Rows.Add(r);
        }

        private void setDataTableForAnalysis(DataTable[] tables, CResultData total, CResultData agency, CBusinessData bd)
        {
            if (total == null || agency == null) return;
            for (int j = 0; j < tables.Length; j++)
            {
                DataTable t = tables[j];
                DataRow r = t.NewRow();
                Int64[] totalArr = null;
                Int64[] agencyArr = null;
                switch (j) {
                    case 0:
                        totalArr = total.getArr전체_리포트용(bd);
                        agencyArr = agency.getArr전체_리포트용(bd);
                        break;
                    case 1:
                        //totalArr = total.getArr전체_리포트용(bd);
                        //agencyArr = agency.getArr전체_리포트용(bd);
                        totalArr = total.getArr도매_수익_비용_및_계산포함();
                        agencyArr = agency.getArr도매_수익_비용_및_계산포함();
                        break;
                    case 2:
                        totalArr = total.getArr소매_수익_비용_및_계산포함();
                        agencyArr = agency.getArr소매_수익_비용_및_계산포함();
                        break;
                }
                for (int i = 0; i < totalArr.Length; i++)
                {
                    r[i * 2] = totalArr[i];
                    r[i * 2 + 1] = agencyArr[i];
                }
                t.Rows.Add(r);
            }
        }
    }
}
