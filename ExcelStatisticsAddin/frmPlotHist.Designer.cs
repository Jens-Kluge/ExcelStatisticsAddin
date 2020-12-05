
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPlotHist));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnExtendRg2 = new System.Windows.Forms.Button();
            this.grpInputData = new System.Windows.Forms.GroupBox();
            this.btnExtendRg1 = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.histogram1 = new ExcelStatisticsAddin.Histogram();
            this.refedData = new VS.NET_RefeditControl.refedit();
            this.refedBins = new VS.NET_RefeditControl.refedit();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.refedOutput = new VS.NET_RefeditControl.refedit();
            this.groupBox1.SuspendLayout();
            this.grpInputData.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnExtendRg2);
            this.groupBox1.Controls.Add(this.refedBins);
            this.groupBox1.Location = new System.Drawing.Point(26, 97);
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
            this.btnExtendRg1.UseVisualStyleBackColor = true;
            this.btnExtendRg1.Click += new System.EventHandler(this.btnExtendRg1_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(47, 274);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(94, 43);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "&OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(158, 274);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(91, 43);
            this.button2.TabIndex = 3;
            this.button2.Text = "&Cancel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // histogram1
            // 
            this.histogram1.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.histogram1.DisplayColor = System.Drawing.Color.Black;
            this.histogram1.Location = new System.Drawing.Point(307, 23);
            this.histogram1.Name = "histogram1";
            this.histogram1.Offset = 20;
            this.histogram1.Size = new System.Drawing.Size(383, 331);
            this.histogram1.TabIndex = 4;
            this.histogram1.Load += new System.EventHandler(this.histogram1_Load);
            // 
            // refedData
            // 
            this.refedData._Excel = null;
            this.refedData.AllowCollapsedFormResize = false;
            this.refedData.Location = new System.Drawing.Point(21, 21);
            this.refedData.Name = "refedData";
            this.refedData.Size = new System.Drawing.Size(182, 22);
            this.refedData.TabIndex = 2;
            // 
            // refedBins
            // 
            this.refedBins._Excel = null;
            this.refedBins.AllowCollapsedFormResize = false;
            this.refedBins.Location = new System.Drawing.Point(21, 21);
            this.refedBins.Name = "refedBins";
            this.refedBins.Size = new System.Drawing.Size(182, 22);
            this.refedBins.TabIndex = 3;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.refedOutput);
            this.groupBox2.Location = new System.Drawing.Point(26, 174);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(236, 59);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Output Range";
            // 
            // refedOutput
            // 
            this.refedOutput._Excel = null;
            this.refedOutput.AllowCollapsedFormResize = false;
            this.refedOutput.Location = new System.Drawing.Point(21, 21);
            this.refedOutput.Name = "refedOutput";
            this.refedOutput.Size = new System.Drawing.Size(182, 22);
            this.refedOutput.TabIndex = 3;
            // 
            // frmPlotHist
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(744, 406);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.histogram1);
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
            this.ResumeLayout(false);

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
        private Histogram histogram1;
        private System.Windows.Forms.GroupBox groupBox2;
        public VS.NET_RefeditControl.refedit refedOutput;
    }
}