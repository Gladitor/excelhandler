namespace ExcelHandler
{
    partial class UserControl_combine
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

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_pickfile = new System.Windows.Forms.Button();
            this.checkBox_firstrowishead = new System.Windows.Forms.CheckBox();
            this.checkBox_allInOneSheet = new System.Windows.Forms.CheckBox();
            this.checkBox_combinBySheet = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBox_appointSheetName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.checkBox_appointsheet = new System.Windows.Forms.CheckBox();
            this.checkBox_horizontal = new System.Windows.Forms.CheckBox();
            this.checkBox_allAppointInOneSheet = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_pickfile
            // 
            this.btn_pickfile.Location = new System.Drawing.Point(768, 34);
            this.btn_pickfile.Name = "btn_pickfile";
            this.btn_pickfile.Size = new System.Drawing.Size(177, 33);
            this.btn_pickfile.TabIndex = 0;
            this.btn_pickfile.Text = "选择文件合并";
            this.btn_pickfile.UseVisualStyleBackColor = true;
            this.btn_pickfile.Click += new System.EventHandler(this.btn_pickfile_Click);
            // 
            // checkBox_firstrowishead
            // 
            this.checkBox_firstrowishead.AutoSize = true;
            this.checkBox_firstrowishead.Checked = true;
            this.checkBox_firstrowishead.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_firstrowishead.Location = new System.Drawing.Point(16, 40);
            this.checkBox_firstrowishead.Name = "checkBox_firstrowishead";
            this.checkBox_firstrowishead.Size = new System.Drawing.Size(142, 22);
            this.checkBox_firstrowishead.TabIndex = 1;
            this.checkBox_firstrowishead.Text = "第一行为列头";
            this.checkBox_firstrowishead.UseVisualStyleBackColor = true;
            // 
            // checkBox_allInOneSheet
            // 
            this.checkBox_allInOneSheet.AutoSize = true;
            this.checkBox_allInOneSheet.Location = new System.Drawing.Point(21, 41);
            this.checkBox_allInOneSheet.Name = "checkBox_allInOneSheet";
            this.checkBox_allInOneSheet.Size = new System.Drawing.Size(313, 22);
            this.checkBox_allInOneSheet.TabIndex = 2;
            this.checkBox_allInOneSheet.Text = "所有Sheet合并到1个EXCEL1个Sheet";
            this.checkBox_allInOneSheet.UseVisualStyleBackColor = true;
            // 
            // checkBox_combinBySheet
            // 
            this.checkBox_combinBySheet.AutoSize = true;
            this.checkBox_combinBySheet.Location = new System.Drawing.Point(21, 85);
            this.checkBox_combinBySheet.Name = "checkBox_combinBySheet";
            this.checkBox_combinBySheet.Size = new System.Drawing.Size(322, 22);
            this.checkBox_combinBySheet.TabIndex = 3;
            this.checkBox_combinBySheet.Text = "所有Sheet合并到1个EXCEL对应Sheet";
            this.checkBox_combinBySheet.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.checkBox_allAppointInOneSheet);
            this.groupBox1.Controls.Add(this.textBox_appointSheetName);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.checkBox_appointsheet);
            this.groupBox1.Controls.Add(this.checkBox_allInOneSheet);
            this.groupBox1.Controls.Add(this.checkBox_combinBySheet);
            this.groupBox1.Location = new System.Drawing.Point(179, 40);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(529, 466);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "类型";
            // 
            // textBox_appointSheetName
            // 
            this.textBox_appointSheetName.Location = new System.Drawing.Point(214, 241);
            this.textBox_appointSheetName.Multiline = true;
            this.textBox_appointSheetName.Name = "textBox_appointSheetName";
            this.textBox_appointSheetName.ReadOnly = true;
            this.textBox_appointSheetName.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox_appointSheetName.Size = new System.Drawing.Size(282, 129);
            this.textBox_appointSheetName.TabIndex = 6;
            this.textBox_appointSheetName.Text = "Sheet1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 244);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(188, 18);
            this.label1.TabIndex = 5;
            this.label1.Text = "Sheet集合(每行一个):";
            // 
            // checkBox_appointsheet
            // 
            this.checkBox_appointsheet.AutoSize = true;
            this.checkBox_appointsheet.Location = new System.Drawing.Point(21, 131);
            this.checkBox_appointsheet.Name = "checkBox_appointsheet";
            this.checkBox_appointsheet.Size = new System.Drawing.Size(340, 22);
            this.checkBox_appointsheet.TabIndex = 4;
            this.checkBox_appointsheet.Text = "按指定Sheet合并到1个EXCEL对应Sheet";
            this.checkBox_appointsheet.UseVisualStyleBackColor = true;
            this.checkBox_appointsheet.CheckedChanged += new System.EventHandler(this.checkBox_appointsheet_CheckedChanged);
            // 
            // checkBox_horizontal
            // 
            this.checkBox_horizontal.AutoSize = true;
            this.checkBox_horizontal.Location = new System.Drawing.Point(16, 81);
            this.checkBox_horizontal.Name = "checkBox_horizontal";
            this.checkBox_horizontal.Size = new System.Drawing.Size(106, 22);
            this.checkBox_horizontal.TabIndex = 5;
            this.checkBox_horizontal.Text = "横向合并";
            this.checkBox_horizontal.UseVisualStyleBackColor = true;
            // 
            // checkBox_allAppointInOneSheet
            // 
            this.checkBox_allAppointInOneSheet.AutoSize = true;
            this.checkBox_allAppointInOneSheet.Location = new System.Drawing.Point(21, 179);
            this.checkBox_allAppointInOneSheet.Name = "checkBox_allAppointInOneSheet";
            this.checkBox_allAppointInOneSheet.Size = new System.Drawing.Size(331, 22);
            this.checkBox_allAppointInOneSheet.TabIndex = 7;
            this.checkBox_allAppointInOneSheet.Text = "按指定Sheet合并到1个EXCEL1个Sheet";
            this.checkBox_allAppointInOneSheet.UseVisualStyleBackColor = true;
            this.checkBox_allAppointInOneSheet.CheckedChanged += new System.EventHandler(this.checkBox_allAppointInOneSheet_CheckedChanged);
            // 
            // UserControl_combine
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.Controls.Add(this.checkBox_horizontal);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.checkBox_firstrowishead);
            this.Controls.Add(this.btn_pickfile);
            this.Name = "UserControl_combine";
            this.Size = new System.Drawing.Size(1020, 642);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_pickfile;
        private System.Windows.Forms.CheckBox checkBox_firstrowishead;
        private System.Windows.Forms.CheckBox checkBox_allInOneSheet;
        private System.Windows.Forms.CheckBox checkBox_combinBySheet;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox checkBox_appointsheet;
        private System.Windows.Forms.TextBox textBox_appointSheetName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox checkBox_horizontal;
        private System.Windows.Forms.CheckBox checkBox_allAppointInOneSheet;
    }
}
