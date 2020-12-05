
namespace ExcelStatisticsAddin
{
    partial class frmDistFit
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmDistFit));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnFill = new System.Windows.Forms.Button();
            this.refedit1 = new VS.NET_RefeditControl.refedit();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.refedit2 = new VS.NET_RefeditControl.refedit();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.grpMethod = new System.Windows.Forms.GroupBox();
            this.rbDensity = new System.Windows.Forms.RadioButton();
            this.rbMoments = new System.Windows.Forms.RadioButton();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.grpMethod.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnFill);
            this.groupBox1.Controls.Add(this.refedit1);
            this.groupBox1.Location = new System.Drawing.Point(22, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(264, 83);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Input Range";
            // 
            // btnFill
            // 
            this.btnFill.Image = global::ExcelStatisticsAddin.Properties.Resources.fill_270_icon;
            this.btnFill.Location = new System.Drawing.Point(231, 36);
            this.btnFill.Name = "btnFill";
            this.btnFill.Size = new System.Drawing.Size(20, 22);
            this.btnFill.TabIndex = 6;
            this.toolTip1.SetToolTip(this.btnFill, "Extend selection downwards");
            this.btnFill.UseVisualStyleBackColor = true;
            this.btnFill.Click += new System.EventHandler(this.btnFill_Click);
            // 
            // refedit1
            // 
            this.refedit1._Excel = null;
            this.refedit1.AllowCollapsedFormResize = false;
            this.refedit1.Location = new System.Drawing.Point(11, 36);
            this.refedit1.Name = "refedit1";
            this.refedit1.Size = new System.Drawing.Size(220, 22);
            this.refedit1.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.refedit2);
            this.groupBox2.Location = new System.Drawing.Point(22, 102);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(264, 83);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Output Range";
            // 
            // refedit2
            // 
            this.refedit2._Excel = null;
            this.refedit2.AllowCollapsedFormResize = false;
            this.refedit2.Location = new System.Drawing.Point(11, 36);
            this.refedit2.Name = "refedit2";
            this.refedit2.Size = new System.Drawing.Size(220, 22);
            this.refedit2.TabIndex = 1;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(313, 23);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(82, 47);
            this.btnOK.TabIndex = 3;
            this.btnOK.Text = "&OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(313, 89);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(82, 47);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // grpMethod
            // 
            this.grpMethod.Controls.Add(this.rbDensity);
            this.grpMethod.Controls.Add(this.rbMoments);
            this.grpMethod.Location = new System.Drawing.Point(22, 206);
            this.grpMethod.Name = "grpMethod";
            this.grpMethod.Size = new System.Drawing.Size(264, 85);
            this.grpMethod.TabIndex = 5;
            this.grpMethod.TabStop = false;
            this.grpMethod.Text = "Method";
            // 
            // rbDensity
            // 
            this.rbDensity.AutoSize = true;
            this.rbDensity.Location = new System.Drawing.Point(21, 48);
            this.rbDensity.Name = "rbDensity";
            this.rbDensity.Size = new System.Drawing.Size(174, 21);
            this.rbDensity.TabIndex = 1;
            this.rbDensity.TabStop = true;
            this.rbDensity.Text = "Energy density method";
            this.rbDensity.UseVisualStyleBackColor = true;
            // 
            // rbMoments
            // 
            this.rbMoments.AutoSize = true;
            this.rbMoments.Location = new System.Drawing.Point(21, 21);
            this.rbMoments.Name = "rbMoments";
            this.rbMoments.Size = new System.Drawing.Size(153, 21);
            this.rbMoments.TabIndex = 0;
            this.rbMoments.TabStop = true;
            this.rbMoments.Text = "Method of moments";
            this.rbMoments.UseVisualStyleBackColor = true;
            // 
            // frmDistFit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(413, 313);
            this.Controls.Add(this.grpMethod);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmDistFit";
            this.Text = "Weibull Fit";
            this.Load += new System.EventHandler(this.frmDistFit_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.grpMethod.ResumeLayout(false);
            this.grpMethod.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        public VS.NET_RefeditControl.refedit refedit1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        public VS.NET_RefeditControl.refedit refedit2;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.GroupBox grpMethod;
        private System.Windows.Forms.RadioButton rbDensity;
        private System.Windows.Forms.RadioButton rbMoments;
        private System.Windows.Forms.Button btnFill;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}