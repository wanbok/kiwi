namespace KIWI
{
    partial class FormPopupChart
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint1 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(0D, 2557523D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint2 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(1D, 2274223D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint3 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(2D, 1719385D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint4 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(3D, 5262966D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint5 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(4D, 8537500D);
            System.Windows.Forms.DataVisualization.Charting.Title title1 = new System.Windows.Forms.DataVisualization.Charting.Title();
            this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
            this.SuspendLayout();
            // 
            // chart1
            // 
            this.chart1.BackImage = "Resources\\charbackground.jpg";
            this.chart1.BackImageWrapMode = System.Windows.Forms.DataVisualization.Charting.ChartImageWrapMode.Scaled;
            chartArea1.Area3DStyle.Enable3D = true;
            chartArea1.BackColor = System.Drawing.Color.Transparent;
            chartArea1.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea1);
            legend1.BackColor = System.Drawing.Color.Transparent;
            legend1.ForeColor = System.Drawing.Color.White;
            legend1.Name = "Legend1";
            this.chart1.Legends.Add(legend1);
            this.chart1.Location = new System.Drawing.Point(0, 33);
            this.chart1.Name = "chart1";
            series1.BackImageTransparentColor = System.Drawing.Color.White;
            series1.ChartArea = "ChartArea1";
            series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            series1.Color = System.Drawing.Color.White;
            series1.LabelBackColor = System.Drawing.Color.Transparent;
            series1.Legend = "Legend1";
            series1.Name = "Series1";
            dataPoint1.Label = "2,557,523";
            dataPoint1.LegendText = "누적가입자 수수료";
            dataPoint2.Label = "2,274,223";
            dataPoint2.LegendText = "월단위 업무취급 수수료";
            dataPoint3.Label = "1,719,385";
            dataPoint3.LegendText = "사업자모델 매입 추가수익";
            dataPoint4.Label = "5,262,966";
            dataPoint4.LabelFormat = "";
            dataPoint4.LegendText = "유통모델 매입 추가수익";
            dataPoint5.Label = "8,537,500";
            dataPoint5.LabelToolTip = "";
            dataPoint5.LegendText = "직영매장 판매수익";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);
            this.chart1.Series.Add(series1);
            this.chart1.Size = new System.Drawing.Size(1372, 787);
            this.chart1.TabIndex = 1;
            this.chart1.Text = "chart1";
            title1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Bold);
            title1.ForeColor = System.Drawing.Color.White;
            title1.Name = "Title1";
            title1.Text = "업계 평균 - 수익계정";
            this.chart1.Titles.Add(title1);
            // 
            // FormPopupChart
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1384, 832);
            this.Controls.Add(this.chart1);
            this.MaximizeBox = false;
            this.Name = "FormPopupChart";
            this.Text = "KIWI ";
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
    }
}