namespace KIWI
{
    partial class FormPopupChartBar
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
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint5 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(4D, 40537500D);
            System.Windows.Forms.DataVisualization.Charting.Series series2 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint6 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(0D, 3000000D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint7 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(1D, 1000000D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint8 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(2D, 2000000D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint9 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(3D, 21000000D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint10 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(4D, 50000000D);
            System.Windows.Forms.DataVisualization.Charting.Series series3 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint11 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(0D, 43200000D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint12 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(1D, 1000000D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint13 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(2D, 2000000D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint14 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(3D, 21000000D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint15 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(4D, 50000000D);
            System.Windows.Forms.DataVisualization.Charting.Title title1 = new System.Windows.Forms.DataVisualization.Charting.Title();
            this.chart7 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            ((System.ComponentModel.ISupportInitialize)(this.chart7)).BeginInit();
            this.SuspendLayout();
            // 
            // chart7
            // 
            this.chart7.BackColor = System.Drawing.Color.Transparent;
            this.chart7.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.chart7.BackImage = "Resources\\back.jpg";
            this.chart7.BackImageWrapMode = System.Windows.Forms.DataVisualization.Charting.ChartImageWrapMode.Scaled;
            chartArea1.Area3DStyle.Enable3D = true;
            chartArea1.Area3DStyle.IsClustered = true;
            chartArea1.AxisX.LabelStyle.ForeColor = System.Drawing.Color.White;
            chartArea1.AxisX.LineColor = System.Drawing.Color.White;
            chartArea1.AxisX.MajorGrid.LineColor = System.Drawing.Color.White;
            chartArea1.AxisX.TitleForeColor = System.Drawing.Color.White;
            chartArea1.AxisY.LabelStyle.ForeColor = System.Drawing.Color.White;
            chartArea1.AxisY.LineColor = System.Drawing.Color.White;
            chartArea1.AxisY.MajorGrid.LineColor = System.Drawing.Color.White;
            chartArea1.AxisY.TitleForeColor = System.Drawing.Color.White;
            chartArea1.BackColor = System.Drawing.Color.Transparent;
            chartArea1.BackImageWrapMode = System.Windows.Forms.DataVisualization.Charting.ChartImageWrapMode.Scaled;
            chartArea1.Name = "ChartArea1";
            this.chart7.ChartAreas.Add(chartArea1);
            legend1.BackColor = System.Drawing.Color.Transparent;
            legend1.ForeColor = System.Drawing.Color.White;
            legend1.Name = "Legend1";
            this.chart7.Legends.Add(legend1);
            this.chart7.Location = new System.Drawing.Point(12, 27);
            this.chart7.Name = "chart7";
            series1.ChartArea = "ChartArea1";
            series1.Color = System.Drawing.Color.Blue;
            series1.LabelForeColor = System.Drawing.Color.White;
            series1.Legend = "Legend1";
            series1.LegendText = "업계평균";
            series1.MarkerColor = System.Drawing.Color.White;
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);
            series2.ChartArea = "ChartArea1";
            series2.Color = System.Drawing.Color.Red;
            series2.LabelForeColor = System.Drawing.Color.White;
            series2.Legend = "Legend1";
            series2.LegendText = "당대리점";
            series2.Name = "Series2";
            dataPoint6.MarkerColor = System.Drawing.Color.White;
            series2.Points.Add(dataPoint6);
            series2.Points.Add(dataPoint7);
            series2.Points.Add(dataPoint8);
            series2.Points.Add(dataPoint9);
            series2.Points.Add(dataPoint10);
            series3.ChartArea = "ChartArea1";
            series3.Color = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            series3.LabelForeColor = System.Drawing.Color.White;
            series3.Legend = "Legend1";
            series3.LegendText = "미래수익";
            series3.Name = "Series3";
            dataPoint11.Label = "43,200,000";
            dataPoint12.Label = "1,000,000";
            dataPoint13.Label = "2,000,000";
            dataPoint14.Label = "21,000,000";
            dataPoint15.Label = "50,000,000";
            series3.Points.Add(dataPoint11);
            series3.Points.Add(dataPoint12);
            series3.Points.Add(dataPoint13);
            series3.Points.Add(dataPoint14);
            series3.Points.Add(dataPoint15);
            this.chart7.Series.Add(series1);
            this.chart7.Series.Add(series2);
            this.chart7.Series.Add(series3);
            this.chart7.Size = new System.Drawing.Size(1360, 793);
            this.chart7.TabIndex = 6;
            this.chart7.Text = "chart7";
            title1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Bold);
            title1.ForeColor = System.Drawing.Color.White;
            title1.Name = "Title1";
            title1.Text = "수익계정";
            this.chart7.Titles.Add(title1);
            // 
            // FormPopupChartBar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1384, 832);
            this.Controls.Add(this.chart7);
            this.MaximizeBox = false;
            this.Name = "FormPopupChartBar";
            this.Text = "KIWI";
            ((System.ComponentModel.ISupportInitialize)(this.chart7)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataVisualization.Charting.Chart chart7;
    }
}