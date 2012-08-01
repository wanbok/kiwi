namespace KIWI
{
    partial class FormPopupChartSample
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
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint1 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(0D, 4584700D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint2 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(1D, 23153895D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint3 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(2D, 26034783D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint4 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(3D, 3386394D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint5 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(4D, 5675947D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint6 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(5D, 10158024D);
            System.Windows.Forms.DataVisualization.Charting.Title title1 = new System.Windows.Forms.DataVisualization.Charting.Title();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea2 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend2 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series2 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint7 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(0D, 2557523D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint8 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(1D, 2274223D);
            System.Windows.Forms.DataVisualization.Charting.Title title2 = new System.Windows.Forms.DataVisualization.Charting.Title();
            this.chart2 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.chart2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
            this.SuspendLayout();
            // 
            // chart2
            // 
            this.chart2.BackImage = "Resources\\back.jpg";
            this.chart2.BackImageWrapMode = System.Windows.Forms.DataVisualization.Charting.ChartImageWrapMode.Scaled;
            chartArea1.Area3DStyle.Enable3D = true;
            chartArea1.Area3DStyle.IsClustered = true;
            chartArea1.AxisX.LabelStyle.ForeColor = System.Drawing.Color.White;
            chartArea1.AxisX.LineColor = System.Drawing.Color.White;
            chartArea1.AxisX.MajorGrid.LineColor = System.Drawing.Color.White;
            chartArea1.AxisY.LabelStyle.ForeColor = System.Drawing.Color.White;
            chartArea1.AxisY.LineColor = System.Drawing.Color.White;
            chartArea1.AxisY.MajorGrid.LineColor = System.Drawing.Color.White;
            chartArea1.BackColor = System.Drawing.Color.Transparent;
            chartArea1.BackImageWrapMode = System.Windows.Forms.DataVisualization.Charting.ChartImageWrapMode.Scaled;
            chartArea1.Name = "ChartArea1";
            this.chart2.ChartAreas.Add(chartArea1);
            legend1.BackColor = System.Drawing.Color.Transparent;
            legend1.ForeColor = System.Drawing.Color.White;
            legend1.Name = "Legend1";
            this.chart2.Legends.Add(legend1);
            this.chart2.Location = new System.Drawing.Point(692, 112);
            this.chart2.Name = "chart2";
            series1.ChartArea = "ChartArea1";
            series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Doughnut;
            series1.Legend = "Legend1";
            series1.LegendText = "aa\\bb\\\\dd\\\\bb\\sds\\\\\\ddsgf";
            series1.Name = "Series1";
            dataPoint1.LegendText = "직원급여(간부급)";
            dataPoint2.LegendText = "직원급여(평사원)";
            dataPoint3.LegendText = "지급 임차료";
            dataPoint4.LegendText = "지급 수수료";
            dataPoint5.LegendText = "판매촉진비";
            dataPoint6.LegendText = "건물관리비";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);
            series1.Points.Add(dataPoint6);
            this.chart2.Series.Add(series1);
            this.chart2.Size = new System.Drawing.Size(679, 708);
            this.chart2.TabIndex = 9;
            this.chart2.Text = "chart7";
            title1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Bold);
            title1.ForeColor = System.Drawing.Color.White;
            title1.Name = "Title1";
            title1.Text = "비용계정";
            this.chart2.Titles.Add(title1);
            // 
            // chart1
            // 
            this.chart1.BackColor = System.Drawing.Color.Transparent;
            this.chart1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.chart1.BackImage = "Resources\\back.jpg";
            this.chart1.BackImageWrapMode = System.Windows.Forms.DataVisualization.Charting.ChartImageWrapMode.Scaled;
            chartArea2.Area3DStyle.Enable3D = true;
            chartArea2.Area3DStyle.IsClustered = true;
            chartArea2.AxisX.LabelStyle.ForeColor = System.Drawing.Color.White;
            chartArea2.AxisX.LineColor = System.Drawing.Color.White;
            chartArea2.AxisX.MajorGrid.LineColor = System.Drawing.Color.White;
            chartArea2.AxisX.TitleForeColor = System.Drawing.Color.White;
            chartArea2.AxisY.LabelStyle.ForeColor = System.Drawing.Color.White;
            chartArea2.AxisY.LineColor = System.Drawing.Color.White;
            chartArea2.AxisY.MajorGrid.LineColor = System.Drawing.Color.White;
            chartArea2.AxisY.TitleForeColor = System.Drawing.Color.White;
            chartArea2.BackColor = System.Drawing.Color.Transparent;
            chartArea2.BackImageWrapMode = System.Windows.Forms.DataVisualization.Charting.ChartImageWrapMode.Scaled;
            chartArea2.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea2);
            legend2.BackColor = System.Drawing.Color.Transparent;
            legend2.ForeColor = System.Drawing.Color.White;
            legend2.Name = "Legend1";
            this.chart1.Legends.Add(legend2);
            this.chart1.Location = new System.Drawing.Point(13, 112);
            this.chart1.Name = "chart1";
            series2.ChartArea = "ChartArea1";
            series2.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            series2.Color = System.Drawing.Color.Blue;
            series2.LabelForeColor = System.Drawing.Color.White;
            series2.Legend = "Legend1";
            series2.MarkerColor = System.Drawing.Color.White;
            series2.Name = "Series1";
            dataPoint7.LegendText = "월평균 업무취급수수료";
            dataPoint8.LegendText = "직영매장 판매 수익";
            series2.Points.Add(dataPoint7);
            series2.Points.Add(dataPoint8);
            this.chart1.Series.Add(series2);
            this.chart1.Size = new System.Drawing.Size(679, 708);
            this.chart1.TabIndex = 8;
            this.chart1.Text = "chart7";
            title2.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Bold);
            title2.ForeColor = System.Drawing.Color.White;
            title2.Name = "Title1";
            title2.Text = "수익계정";
            this.chart1.Titles.Add(title2);
            // 
            // textBox5
            // 
            this.textBox5.BackColor = System.Drawing.Color.LimeGreen;
            this.textBox5.Font = new System.Drawing.Font("굴림", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.textBox5.Location = new System.Drawing.Point(640, 74);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(104, 20);
            this.textBox5.TabIndex = 107;
            this.textBox5.Text = "0";
            this.textBox5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.textBox5.TextChanged += new System.EventHandler(this.textBox5_TextChanged);
            this.textBox5.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox5_KeyPress);
            // 
            // textBox4
            // 
            this.textBox4.BackColor = System.Drawing.Color.LimeGreen;
            this.textBox4.Font = new System.Drawing.Font("굴림", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.textBox4.Location = new System.Drawing.Point(640, 44);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(104, 20);
            this.textBox4.TabIndex = 108;
            this.textBox4.Text = "0";
            this.textBox4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.textBox4.TextChanged += new System.EventHandler(this.textBox4_TextChanged);
            this.textBox4.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox4_KeyPress);
            // 
            // textBox3
            // 
            this.textBox3.BackColor = System.Drawing.Color.LimeGreen;
            this.textBox3.Font = new System.Drawing.Font("굴림", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.textBox3.Location = new System.Drawing.Point(640, 12);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(104, 20);
            this.textBox3.TabIndex = 109;
            this.textBox3.Text = "0";
            this.textBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.textBox3.TextChanged += new System.EventHandler(this.textBox3_TextChanged);
            this.textBox3.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox3_KeyPress);
            // 
            // textBox2
            // 
            this.textBox2.BackColor = System.Drawing.Color.LimeGreen;
            this.textBox2.Font = new System.Drawing.Font("굴림", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.textBox2.Location = new System.Drawing.Point(236, 57);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(104, 20);
            this.textBox2.TabIndex = 106;
            this.textBox2.Text = "0";
            this.textBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.textBox2.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
            this.textBox2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox2_KeyPress);
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.Color.LimeGreen;
            this.textBox1.Font = new System.Drawing.Font("굴림", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.textBox1.Location = new System.Drawing.Point(236, 31);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(104, 20);
            this.textBox1.TabIndex = 105;
            this.textBox1.Text = "0";
            this.textBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            this.textBox1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // textBox6
            // 
            this.textBox6.BackColor = System.Drawing.Color.LimeGreen;
            this.textBox6.Font = new System.Drawing.Font("굴림", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.textBox6.Location = new System.Drawing.Point(874, 12);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(104, 20);
            this.textBox6.TabIndex = 110;
            this.textBox6.Text = "0";
            this.textBox6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.textBox6.TextChanged += new System.EventHandler(this.textBox6_TextChanged);
            this.textBox6.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox6_KeyPress);
            // 
            // textBox7
            // 
            this.textBox7.BackColor = System.Drawing.Color.LimeGreen;
            this.textBox7.Font = new System.Drawing.Font("굴림", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.textBox7.Location = new System.Drawing.Point(874, 44);
            this.textBox7.Name = "textBox7";
            this.textBox7.Size = new System.Drawing.Size(104, 20);
            this.textBox7.TabIndex = 111;
            this.textBox7.Text = "0";
            this.textBox7.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.textBox7.TextChanged += new System.EventHandler(this.textBox7_TextChanged);
            this.textBox7.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox7_KeyPress);
            // 
            // textBox8
            // 
            this.textBox8.BackColor = System.Drawing.Color.LimeGreen;
            this.textBox8.Font = new System.Drawing.Font("굴림", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.textBox8.Location = new System.Drawing.Point(874, 74);
            this.textBox8.Name = "textBox8";
            this.textBox8.Size = new System.Drawing.Size(104, 20);
            this.textBox8.TabIndex = 112;
            this.textBox8.Text = "0";
            this.textBox8.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.textBox8.TextChanged += new System.EventHandler(this.textBox8_TextChanged);
            this.textBox8.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox8_KeyPress);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(101, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(129, 12);
            this.label1.TabIndex = 113;
            this.label1.Text = "월평균 업무취급수수료";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(125, 59);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(105, 12);
            this.label2.TabIndex = 114;
            this.label2.Text = "직영매장 판매수익";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(529, 16);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(99, 12);
            this.label3.TabIndex = 115;
            this.label3.Text = "직원급여(간부급)";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(529, 46);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(99, 12);
            this.label4.TabIndex = 116;
            this.label4.Text = "직원급여(평사원)";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(559, 76);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(69, 12);
            this.label5.TabIndex = 117;
            this.label5.Text = "지급 임차료";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(792, 16);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(65, 12);
            this.label6.TabIndex = 117;
            this.label6.Text = "지급수수료";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(792, 46);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(65, 12);
            this.label7.TabIndex = 117;
            this.label7.Text = "판매촉진비";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(792, 76);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(65, 12);
            this.label8.TabIndex = 117;
            this.label8.Text = "건물관리비";
            // 
            // FormPopupChartSample
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1384, 832);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox8);
            this.Controls.Add(this.textBox7);
            this.Controls.Add(this.textBox6);
            this.Controls.Add(this.textBox5);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.chart2);
            this.Controls.Add(this.chart1);
            this.Name = "FormPopupChartSample";
            this.Text = "KIWI";
            ((System.ComponentModel.ISupportInitialize)(this.chart2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataVisualization.Charting.Chart chart2;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox6;
        private System.Windows.Forms.TextBox textBox7;
        private System.Windows.Forms.TextBox textBox8;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
    }
}