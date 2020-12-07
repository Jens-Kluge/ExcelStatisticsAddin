
namespace ExcelStatisticsAddin
{
    partial class frmPlotHist
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPlotHist));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnExtendRg2 = new System.Windows.Forms.Button();
            this.grpInputData = new System.Windows.Forms.GroupBox();
            this.btnExtendRg1 = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.plotView1 = new OxyPlot.WindowsForms.PlotView();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.chkOverlay = new System.Windows.Forms.CheckBox();
            this.txtTitle = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtLegend = new System.Windows.Forms.TextBox();
            this.refedOutput = new VS.NET_RefeditControl.refedit();
            this.refedData = new VS.NET_RefeditControl.refedit();
            this.refedBins = new VS.NET_RefeditControl.refedit();
            this.groupBox1.SuspendLayout();
            this.grpInputData.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnExtendRg2);
            this.groupBox1.Controls.Add(this.refedBins);
            this.groupBox1.Location = new System.Drawing.Point(26, 92);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(236, 59);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Bins";
            // 
            // btnExtendRg2
            // 
            this.btnExtendRg2.Image = global::ExcelStatisticsAddin.Properties.Resources.fill_270_icon;
            this.btnExtendRg2.Location = new System.Drawing.Point(204, 21);
            this.btnExtendRg2.Name = "btnExtendRg2";
            this.btnExtendRg2.Size = new System.Drawing.Size(19, 23);
            this.btnExtendRg2.TabIndex = 5;
            this.toolTip1.SetToolTip(this.btnExtendRg2, "extend selection downwards");
            this.btnExtendRg2.UseVisualStyleBackColor = true;
            this.btnExtendRg2.Click += new System.EventHandler(this.btnExtendRg2_Click);
            // 
            // grpInputData
            // 
            this.grpInputData.Controls.Add(this.refedData);
            this.grpInputData.Controls.Add(this.btnExtendRg1);
            this.grpInputData.Location = new System.Drawing.Point(26, 23);
            this.grpInputData.Name = "grpInputData";
            this.grpInputData.Size = new System.Drawing.Size(236, 59);
            this.grpInputData.TabIndex = 1;
            this.grpInputData.TabStop = false;
            this.grpInputData.Text = "Data";
            // 
            // btnExtendRg1
            // 
            this.btnExtendRg1.Image = global::ExcelStatisticsAddin.Properties.Resources.fill_270_icon;
            this.btnExtendRg1.Location = new System.Drawing.Point(204, 20);
            this.btnExtendRg1.Name = "btnExtendRg1";
            this.btnExtendRg1.Size = new System.Drawing.Size(19, 23);
            this.btnExtendRg1.TabIndex = 4;
            this.toolTip1.SetToolTip(this.btnExtendRg1, "extend selection downwards");
            this.btnExtendRg1.UseVisualStyleBackColor = true;
            this.btnExtendRg1.Click += new System.EventHandler(this.btnExtendRg1_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(47, 453);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(94, 43);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "&Plot";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(158, 453);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(91, 43);
            this.button2.TabIndex = 3;
            this.button2.Text = "&Close";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.refedOutput);
            this.groupBox2.Location = new System.Drawing.Point(26, 170);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(236, 59);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Output Range";
            // 
            // plotView1
            // 
            this.plotView1.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.plotView1.Location = new System.Drawing.Point(311, 32);
            this.plotView1.Name = "plotView1";
            this.plotView1.PanCursor = System.Windows.Forms.Cursors.Hand;
            this.plotView1.Size = new System.Drawing.Size(700, 464);
            this.plotView1.TabIndex = 4;
            this.plotView1.Text = "plotView1";
            this.plotView1.ZoomHorizontalCursor = System.Windows.Forms.Cursors.SizeWE;
            this.plotView1.ZoomRectangleCursor = System.Windows.Forms.Cursors.SizeNWSE;
            this.plotView1.ZoomVerticalCursor = System.Windows.Forms.Cursors.SizeNS;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(32, 237);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 17);
            this.label1.TabIndex = 8;
            this.label1.Text = "X-Axis Units";
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.Location = new System.Drawing.Point(35, 260);
            this.numericUpDown1.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(70, 22);
            this.numericUpDown1.TabIndex = 9;
            this.numericUpDown1.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDown1.ValueChanged += new System.EventHandler(this.numericUpDown1_ValueChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(111, 262);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 17);
            this.label2.TabIndex = 10;
            this.label2.Text = "x bin size";
            // 
            // chkOverlay
            // 
            this.chkOverlay.AutoSize = true;
            this.chkOverlay.Location = new System.Drawing.Point(35, 299);
            this.chkOverlay.Name = "chkOverlay";
            this.chkOverlay.Size = new System.Drawing.Size(202, 21);
            this.chkOverlay.TabIndex = 11;
            this.chkOverlay.Text = "Overlay Weibull distribution";
            this.chkOverlay.UseVisualStyleBackColor = true;
            this.chkOverlay.CheckedChanged += new System.EventHandler(this.chkOverlay_CheckedChanged);
            // 
            // txtTitle
            // 
            this.txtTitle.Location = new System.Drawing.Point(35, 355);
            this.txtTitle.Name = "txtTitle";
            this.txtTitle.Size = new System.Drawing.Size(227, 22);
            this.txtTitle.TabIndex = 12;
            this.txtTitle.Text = "Histogram";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(32, 332);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 17);
            this.label3.TabIndex = 13;
            this.label3.Text = "Title";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(32, 380);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 17);
            this.label4.TabIndex = 14;
            this.label4.Text = "Legend";
            // 
            // txtLegend
            // 
            this.txtLegend.Location = new System.Drawing.Point(35, 400);
            this.txtLegend.Name = "txtLegend";
            this.txtLegend.Size = new System.Drawing.Size(227, 22);
            this.txtLegend.TabIndex = 15;
            this.txtLegend.Text = "Data Series";
            // 
            // refedOutput
            // 
            this.refedOutput._Excel = null;
            this.refedOutput.AllowCollapsedFormResize = false;
            this.refedOutput.Location = new System.Drawing.Point(21, 21);
            this.refedOutput.Name = "refedOutput";
            this.refedOutput.Size = new System.Drawing.Size(182, 22);
            this.refedOutput.TabIndex = 2;
            // 
            // refedData
            // 
            this.refedData._Excel = null;
            this.refedData.AllowCollapsedFormResize = false;
            this.refedData.Location = new System.Drawing.Point(21, 21);
            this.refedData.Name = "refedData";
            this.refedData.Size = new System.Drawing.Size(182, 22);
            this.refedData.TabIndex = 0;
            // 
            // refedBins
            // 
            this.refedBins._Excel = null;
            this.refedBins.AllowCollapsedFormResize = false;
            this.refedBins.Location = new System.Drawing.Point(21, 21);
            this.refedBins.Name = "refedBins";
            this.refedBins.Size = new System.Drawing.Size(182, 22);
            this.refedBins.TabIndex = 1;
            // 
            // frmPlotHist
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1102, 527);
            this.Controls.Add(this.txtLegend);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtTitle);
            this.Controls.Add(this.chkOverlay);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.numericUpDown1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.plotView1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.grpInputData);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmPlotHist";
            this.Text = "Plot Histogram";
            this.Load += new System.EventHandler(this.frmPlotHist_Load);
            this.groupBox1.ResumeLayout(false);
            this.grpInputData.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox grpInputData;
        public VS.NET_RefeditControl.refedit refedData;
        public VS.NET_RefeditControl.refedit refedBins;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button btnExtendRg2;
        private System.Windows.Forms.Button btnExtendRg1;
        private System.Windows.Forms.GroupBox groupBox2;
        public VS.NET_RefeditControl.refedit refedOutput;
        private OxyPlot.WindowsForms.PlotView plotView1;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown numericUpDown1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox chkOverlay;
        private System.Windows.Forms.TextBox txtTitle;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtLegend;
    }
}