namespace Winform
{
    partial class ExcelFrm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnSaveAsExcel = new System.Windows.Forms.Button();
            this.cmbProductType = new System.Windows.Forms.ComboBox();
            this.lblProductType = new System.Windows.Forms.Label();
            this.rdbtnXls = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdbtnXlsx = new System.Windows.Forms.RadioButton();
            this.btnSaveName = new System.Windows.Forms.Button();
            this.lblFilename = new System.Windows.Forms.Label();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnSaveAsExcel
            // 
            this.btnSaveAsExcel.Location = new System.Drawing.Point(284, 124);
            this.btnSaveAsExcel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnSaveAsExcel.Name = "btnSaveAsExcel";
            this.btnSaveAsExcel.Size = new System.Drawing.Size(79, 29);
            this.btnSaveAsExcel.TabIndex = 0;
            this.btnSaveAsExcel.Text = "导出";
            this.btnSaveAsExcel.UseVisualStyleBackColor = true;
            this.btnSaveAsExcel.Click += new System.EventHandler(this.btnSaveAsExcel_Click);
            // 
            // cmbProductType
            // 
            this.cmbProductType.FormattingEnabled = true;
            this.cmbProductType.Location = new System.Drawing.Point(87, 11);
            this.cmbProductType.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cmbProductType.Name = "cmbProductType";
            this.cmbProductType.Size = new System.Drawing.Size(278, 25);
            this.cmbProductType.TabIndex = 1;
            // 
            // lblProductType
            // 
            this.lblProductType.AutoSize = true;
            this.lblProductType.Location = new System.Drawing.Point(14, 14);
            this.lblProductType.Name = "lblProductType";
            this.lblProductType.Size = new System.Drawing.Size(68, 17);
            this.lblProductType.TabIndex = 2;
            this.lblProductType.Text = "产出物类型";
            // 
            // rdbtnXls
            // 
            this.rdbtnXls.AutoSize = true;
            this.rdbtnXls.Location = new System.Drawing.Point(27, 24);
            this.rdbtnXls.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.rdbtnXls.Name = "rdbtnXls";
            this.rdbtnXls.Size = new System.Drawing.Size(113, 21);
            this.rdbtnXls.TabIndex = 3;
            this.rdbtnXls.Tag = "xls";
            this.rdbtnXls.Text = "Excel 2003(.xls)";
            this.rdbtnXls.UseVisualStyleBackColor = true;
            this.rdbtnXls.CheckedChanged += new System.EventHandler(this.ExcelFormat_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rdbtnXlsx);
            this.groupBox1.Controls.Add(this.rdbtnXls);
            this.groupBox1.Location = new System.Drawing.Point(17, 75);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox1.Size = new System.Drawing.Size(175, 84);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "导出格式";
            // 
            // rdbtnXlsx
            // 
            this.rdbtnXlsx.AutoSize = true;
            this.rdbtnXlsx.Checked = true;
            this.rdbtnXlsx.Location = new System.Drawing.Point(27, 48);
            this.rdbtnXlsx.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.rdbtnXlsx.Name = "rdbtnXlsx";
            this.rdbtnXlsx.Size = new System.Drawing.Size(119, 21);
            this.rdbtnXlsx.TabIndex = 3;
            this.rdbtnXlsx.TabStop = true;
            this.rdbtnXlsx.Tag = "xlsx";
            this.rdbtnXlsx.Text = "Excel 2007(.xlsx)";
            this.rdbtnXlsx.UseVisualStyleBackColor = true;
            this.rdbtnXlsx.CheckedChanged += new System.EventHandler(this.ExcelFormat_CheckedChanged);
            // 
            // btnSaveName
            // 
            this.btnSaveName.Location = new System.Drawing.Point(331, 43);
            this.btnSaveName.Name = "btnSaveName";
            this.btnSaveName.Size = new System.Drawing.Size(34, 23);
            this.btnSaveName.TabIndex = 6;
            this.btnSaveName.Text = "...";
            this.btnSaveName.UseVisualStyleBackColor = true;
            this.btnSaveName.Click += new System.EventHandler(this.btnSaveName_Click);
            // 
            // lblFilename
            // 
            this.lblFilename.AutoSize = true;
            this.lblFilename.Location = new System.Drawing.Point(25, 46);
            this.lblFilename.Name = "lblFilename";
            this.lblFilename.Size = new System.Drawing.Size(56, 17);
            this.lblFilename.TabIndex = 2;
            this.lblFilename.Text = "输出路径";
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(87, 43);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(237, 23);
            this.txtFileName.TabIndex = 5;
            // 
            // ExcelFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(375, 172);
            this.Controls.Add(this.btnSaveName);
            this.Controls.Add(this.txtFileName);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.lblFilename);
            this.Controls.Add(this.lblProductType);
            this.Controls.Add(this.cmbProductType);
            this.Controls.Add(this.btnSaveAsExcel);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "ExcelFrm";
            this.Text = "建模产出物导出";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSaveAsExcel;
        private System.Windows.Forms.ComboBox cmbProductType;
        private System.Windows.Forms.Label lblProductType;
        private System.Windows.Forms.RadioButton rdbtnXls;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdbtnXlsx;
        private System.Windows.Forms.Button btnSaveName;
        private System.Windows.Forms.Label lblFilename;
        private System.Windows.Forms.TextBox txtFileName;

    }
}

