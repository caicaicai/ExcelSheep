namespace ExcelSheep
{
    partial class ExcelSheep
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
            this.selectBtn = new System.Windows.Forms.Button();
            this.fileName = new System.Windows.Forms.Label();
            this.addFilterBtn = new System.Windows.Forms.Button();
            this.addExportBtn = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkHasHeader = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.logBox = new System.Windows.Forms.TextBox();
            this.dataGroupBox = new System.Windows.Forms.GroupBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.filterBtn = new System.Windows.Forms.Button();
            this.handelDataBtn = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.dataListCBox = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.dataGroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // selectBtn
            // 
            this.selectBtn.Location = new System.Drawing.Point(6, 20);
            this.selectBtn.Name = "selectBtn";
            this.selectBtn.Size = new System.Drawing.Size(75, 23);
            this.selectBtn.TabIndex = 0;
            this.selectBtn.Text = "选择文件";
            this.selectBtn.UseVisualStyleBackColor = true;
            this.selectBtn.Click += new System.EventHandler(this.selectBtn_Click);
            // 
            // fileName
            // 
            this.fileName.AutoSize = true;
            this.fileName.BackColor = System.Drawing.SystemColors.Info;
            this.fileName.Font = new System.Drawing.Font("宋体", 13F);
            this.fileName.Location = new System.Drawing.Point(87, 21);
            this.fileName.MaximumSize = new System.Drawing.Size(155, 22);
            this.fileName.MinimumSize = new System.Drawing.Size(155, 22);
            this.fileName.Name = "fileName";
            this.fileName.Size = new System.Drawing.Size(155, 22);
            this.fileName.TabIndex = 1;
            this.fileName.Text = "请先添加基础数据";
            // 
            // addFilterBtn
            // 
            this.addFilterBtn.Location = new System.Drawing.Point(316, 55);
            this.addFilterBtn.Name = "addFilterBtn";
            this.addFilterBtn.Size = new System.Drawing.Size(75, 23);
            this.addFilterBtn.TabIndex = 4;
            this.addFilterBtn.Text = "过滤数据";
            this.addFilterBtn.UseVisualStyleBackColor = true;
            this.addFilterBtn.Click += new System.EventHandler(this.addFilterBtn_Click);
            // 
            // addExportBtn
            // 
            this.addExportBtn.Location = new System.Drawing.Point(397, 55);
            this.addExportBtn.Name = "addExportBtn";
            this.addExportBtn.Size = new System.Drawing.Size(75, 23);
            this.addExportBtn.TabIndex = 5;
            this.addExportBtn.Text = "设置导出";
            this.addExportBtn.UseVisualStyleBackColor = true;
            this.addExportBtn.Click += new System.EventHandler(this.addExportBtn_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.dataListCBox);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.handelDataBtn);
            this.groupBox1.Controls.Add(this.filterBtn);
            this.groupBox1.Controls.Add(this.chkHasHeader);
            this.groupBox1.Controls.Add(this.selectBtn);
            this.groupBox1.Controls.Add(this.addExportBtn);
            this.groupBox1.Controls.Add(this.fileName);
            this.groupBox1.Controls.Add(this.addFilterBtn);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(640, 84);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "操作";
            // 
            // chkHasHeader
            // 
            this.chkHasHeader.AutoSize = true;
            this.chkHasHeader.Location = new System.Drawing.Point(316, 24);
            this.chkHasHeader.Name = "chkHasHeader";
            this.chkHasHeader.Size = new System.Drawing.Size(216, 16);
            this.chkHasHeader.TabIndex = 6;
            this.chkHasHeader.Text = "数据包含标题(标题字段不允许重复)";
            this.chkHasHeader.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.logBox);
            this.groupBox2.Location = new System.Drawing.Point(12, 102);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(640, 107);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "日志";
            // 
            // logBox
            // 
            this.logBox.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.logBox.Location = new System.Drawing.Point(7, 21);
            this.logBox.Multiline = true;
            this.logBox.Name = "logBox";
            this.logBox.ReadOnly = true;
            this.logBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.logBox.Size = new System.Drawing.Size(627, 93);
            this.logBox.TabIndex = 0;
            this.logBox.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.logBox_MouseDoubleClick);
            // 
            // dataGroupBox
            // 
            this.dataGroupBox.Controls.Add(this.dataGridView1);
            this.dataGroupBox.Location = new System.Drawing.Point(12, 215);
            this.dataGroupBox.Name = "dataGroupBox";
            this.dataGroupBox.Size = new System.Drawing.Size(640, 437);
            this.dataGroupBox.TabIndex = 8;
            this.dataGroupBox.TabStop = false;
            this.dataGroupBox.Text = "数据列表";
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(6, 20);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(628, 411);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.DataSourceChanged += new System.EventHandler(this.dataGridView1_DataSourceChanged);
            // 
            // filterBtn
            // 
            this.filterBtn.Location = new System.Drawing.Point(478, 55);
            this.filterBtn.Name = "filterBtn";
            this.filterBtn.Size = new System.Drawing.Size(75, 23);
            this.filterBtn.TabIndex = 8;
            this.filterBtn.Text = "处理数据";
            this.filterBtn.UseVisualStyleBackColor = true;
            this.filterBtn.Click += new System.EventHandler(this.filterBtn_Click);
            // 
            // handelDataBtn
            // 
            this.handelDataBtn.Location = new System.Drawing.Point(559, 55);
            this.handelDataBtn.Name = "handelDataBtn";
            this.handelDataBtn.Size = new System.Drawing.Size(75, 23);
            this.handelDataBtn.TabIndex = 9;
            this.handelDataBtn.Text = "导出数据";
            this.handelDataBtn.UseVisualStyleBackColor = true;
            this.handelDataBtn.Click += new System.EventHandler(this.handelDataBtn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 60);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 10;
            this.label1.Text = "查看的数据";
            // 
            // dataListCBox
            // 
            this.dataListCBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.dataListCBox.FormattingEnabled = true;
            this.dataListCBox.Location = new System.Drawing.Point(90, 57);
            this.dataListCBox.Name = "dataListCBox";
            this.dataListCBox.Size = new System.Drawing.Size(152, 20);
            this.dataListCBox.TabIndex = 11;
            this.dataListCBox.SelectedIndexChanged += new System.EventHandler(this.dataListCBox_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label2.Location = new System.Drawing.Point(0, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(640, 2);
            this.label2.TabIndex = 12;
            // 
            // ExcelSheep
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(660, 690);
            this.Controls.Add(this.dataGroupBox);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ExcelSheep";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel 筛选助手";
            this.Load += new System.EventHandler(this.ExcelSheep_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.dataGroupBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button selectBtn;
        private System.Windows.Forms.Label fileName;
        private System.Windows.Forms.Button addFilterBtn;
        private System.Windows.Forms.Button addExportBtn;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox logBox;
        private System.Windows.Forms.GroupBox dataGroupBox;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.CheckBox chkHasHeader;
        private System.Windows.Forms.Button filterBtn;
        private System.Windows.Forms.Button handelDataBtn;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox dataListCBox;
        private System.Windows.Forms.Label label2;

    }
}

