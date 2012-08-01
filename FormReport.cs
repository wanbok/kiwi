using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KIWI
{
    public partial class FormReport : Form
    {
        private DataSet mDs;
        public FormReport()
        {
            InitializeComponent();
        }

        public FormReport(DataSet ds)
        {
            InitializeComponent();

            mDs = ds;
            CrystalReport1 report = (CrystalReport1)crystalReportViewer1.ReportSource;
            report.SetDataSource(ds);
            report.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape;
            crystalReportViewer1.Refresh();
        }
    }
}
