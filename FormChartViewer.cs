using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace KIWI
{
    public partial class FormChartViewer : Form
    {
        public FormChartViewer()
        {
            InitializeComponent();
        }

        public void MakeChart(Chart chart)
        {
            if (chart != null)
            {
                chart1 = chart;
            }
        }
    }
}
