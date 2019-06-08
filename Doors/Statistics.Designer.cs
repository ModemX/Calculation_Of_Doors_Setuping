namespace Doors
{
    partial class Statistics
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
            System.Windows.Forms.DataVisualization.Charting.Title title1 = new System.Windows.Forms.DataVisualization.Charting.Title();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea2 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend2 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series2 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series3 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series4 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Title title2 = new System.Windows.Forms.DataVisualization.Charting.Title();
            this.chartСпросЗаМесяц = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.chartСпросЗаКвартал = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.СоставитьОтчет = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.chartСпросЗаМесяц)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chartСпросЗаКвартал)).BeginInit();
            this.SuspendLayout();
            // 
            // chartСпросЗаМесяц
            // 
            chartArea1.Name = "ChartArea1";
            this.chartСпросЗаМесяц.ChartAreas.Add(chartArea1);
            legend1.Name = "Legend1";
            this.chartСпросЗаМесяц.Legends.Add(legend1);
            this.chartСпросЗаМесяц.Location = new System.Drawing.Point(13, 13);
            this.chartСпросЗаМесяц.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chartСпросЗаМесяц.Name = "chartСпросЗаМесяц";
            series1.ChartArea = "ChartArea1";
            series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            series1.CustomProperties = "PieDrawingStyle=Concave";
            series1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            series1.Legend = "Legend1";
            series1.Name = "Series1";
            this.chartСпросЗаМесяц.Series.Add(series1);
            this.chartСпросЗаМесяц.Size = new System.Drawing.Size(743, 664);
            this.chartСпросЗаМесяц.TabIndex = 0;
            this.chartСпросЗаМесяц.Text = "Тест";
            title1.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            title1.Name = "Title1";
            title1.Text = "Спрос товара за месяц, группировка по профилю";
            this.chartСпросЗаМесяц.Titles.Add(title1);
            // 
            // chartСпросЗаКвартал
            // 
            chartArea2.AxisX.MajorGrid.Enabled = false;
            chartArea2.AxisX.Title = "Месяц";
            chartArea2.AxisY.Title = "Количество дверей";
            chartArea2.Name = "ChartArea1";
            this.chartСпросЗаКвартал.ChartAreas.Add(chartArea2);
            legend2.Name = "Legend1";
            this.chartСпросЗаКвартал.Legends.Add(legend2);
            this.chartСпросЗаКвартал.Location = new System.Drawing.Point(764, 13);
            this.chartСпросЗаКвартал.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chartСпросЗаКвартал.Name = "chartСпросЗаКвартал";
            series2.ChartArea = "ChartArea1";
            series2.CustomProperties = "DrawSideBySide=False, DrawingStyle=Emboss";
            series2.Legend = "Legend1";
            series2.Name = "Месяц1";
            series3.ChartArea = "ChartArea1";
            series3.CustomProperties = "DrawingStyle=Emboss";
            series3.Legend = "Legend1";
            series3.Name = "Месяц2";
            series4.ChartArea = "ChartArea1";
            series4.CustomProperties = "DrawingStyle=Emboss";
            series4.Legend = "Legend1";
            series4.Name = "Месяц3";
            this.chartСпросЗаКвартал.Series.Add(series2);
            this.chartСпросЗаКвартал.Series.Add(series3);
            this.chartСпросЗаКвартал.Series.Add(series4);
            this.chartСпросЗаКвартал.Size = new System.Drawing.Size(535, 587);
            this.chartСпросЗаКвартал.TabIndex = 1;
            this.chartСпросЗаКвартал.Text = "chart2";
            title2.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            title2.Name = "Title1";
            title2.Text = "Статистика проданных дверей за квартал";
            this.chartСпросЗаКвартал.Titles.Add(title2);
            // 
            // СоставитьОтчет
            // 
            this.СоставитьОтчет.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.СоставитьОтчет.Location = new System.Drawing.Point(763, 607);
            this.СоставитьОтчет.Name = "СоставитьОтчет";
            this.СоставитьОтчет.Size = new System.Drawing.Size(536, 70);
            this.СоставитьОтчет.TabIndex = 2;
            this.СоставитьОтчет.Text = "Составить отчет";
            this.СоставитьОтчет.UseVisualStyleBackColor = true;
            this.СоставитьОтчет.Click += new System.EventHandler(this.СоставитьОтчет_Click);
            // 
            // Statistics
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1312, 690);
            this.Controls.Add(this.chartСпросЗаМесяц);
            this.Controls.Add(this.chartСпросЗаКвартал);
            this.Controls.Add(this.СоставитьОтчет);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Statistics";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Статистика продаж";
            ((System.ComponentModel.ISupportInitialize)(this.chartСпросЗаМесяц)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chartСпросЗаКвартал)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataVisualization.Charting.Chart chartСпросЗаМесяц;
        private System.Windows.Forms.DataVisualization.Charting.Chart chartСпросЗаКвартал;
        private System.Windows.Forms.Button СоставитьОтчет;
    }
}