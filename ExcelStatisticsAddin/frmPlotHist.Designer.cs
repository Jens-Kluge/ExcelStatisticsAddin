
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
            this.grpInputData = new System.Windows.Forms.GroupBox();
            this.refedData = new VS.NET_RefeditControl.refedit();
            this.refedit2 = new VS.NET_RefeditControl.refedit();
            this.btnOK = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.grpInputData.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.refedit2);
            this.groupBox1.Location = new System.Drawing.Point(26, 97);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(228, 59);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Bins";
            // 
            // grpInputData
            // 
            this.grpInputData.Controls.Add(this.refedData);
            this.grpInputData.Location = new System.Drawing.Point(26, 23);
            this.grpInputData.Name = "grpInputData";
            this.grpInputData.Size = new System.Drawing.Size(228, 59);
            this.grpInputData.TabIndex = 1;
            this.grpInputData.TabStop = false;
            this.grpInputData.Text = "Data";
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
            // refedit2
            // 
            this.refedit2._Excel = null;
            this.refedit2.AllowCollapsedFormResize = false;
            this.refedit2.Location = new System.Drawing.Point(21, 21);
            this.refedit2.Name = "refedit2";
            this.refedit2.Size = new System.Drawing.Size(182, 22);
            this.refedit2.TabIndex = 3;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(279, 28);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 43);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "&OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(279, 106);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 43);
            this.button2.TabIndex = 3;
            this.button2.Text = "&Cancel";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(381, 201);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.grpInputData);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Plot Histogram";
            this.groupBox1.ResumeLayout(false);
            this.grpInputData.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox grpInputData;
        private VS.NET_RefeditControl.refedit refedData;
        private VS.NET_RefeditControl.refedit refedit2;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button button2;
    }
}