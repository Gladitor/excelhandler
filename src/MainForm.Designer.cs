namespace ExcelHandler
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripDropDownButton_file = new System.Windows.Forms.ToolStripDropDownButton();
            this.tsmi_exit = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripDropDownButton_view = new System.Windows.Forms.ToolStripDropDownButton();
            this.tsmi_combine = new System.Windows.Forms.ToolStripMenuItem();
            this.panel_main = new System.Windows.Forms.Panel();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripDropDownButton_file,
            this.toolStripDropDownButton_view});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1018, 33);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripDropDownButton_file
            // 
            this.toolStripDropDownButton_file.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripDropDownButton_file.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmi_exit});
            this.toolStripDropDownButton_file.Image = ((System.Drawing.Image)(resources.GetObject("toolStripDropDownButton_file.Image")));
            this.toolStripDropDownButton_file.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripDropDownButton_file.Name = "toolStripDropDownButton_file";
            this.toolStripDropDownButton_file.Size = new System.Drawing.Size(64, 28);
            this.toolStripDropDownButton_file.Text = "文件";
            // 
            // tsmi_exit
            // 
            this.tsmi_exit.Name = "tsmi_exit";
            this.tsmi_exit.Size = new System.Drawing.Size(146, 34);
            this.tsmi_exit.Text = "退出";
            this.tsmi_exit.Click += new System.EventHandler(this.tsmi_exit_Click);
            // 
            // toolStripDropDownButton_view
            // 
            this.toolStripDropDownButton_view.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripDropDownButton_view.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmi_combine});
            this.toolStripDropDownButton_view.Image = ((System.Drawing.Image)(resources.GetObject("toolStripDropDownButton_view.Image")));
            this.toolStripDropDownButton_view.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripDropDownButton_view.Name = "toolStripDropDownButton_view";
            this.toolStripDropDownButton_view.Size = new System.Drawing.Size(64, 28);
            this.toolStripDropDownButton_view.Text = "视图";
            // 
            // tsmi_combine
            // 
            this.tsmi_combine.Name = "tsmi_combine";
            this.tsmi_combine.Size = new System.Drawing.Size(146, 34);
            this.tsmi_combine.Text = "合并";
            this.tsmi_combine.Click += new System.EventHandler(this.tsmi_combine_Click);
            // 
            // panel_main
            // 
            this.panel_main.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel_main.Location = new System.Drawing.Point(0, 33);
            this.panel_main.Name = "panel_main";
            this.panel_main.Size = new System.Drawing.Size(1018, 556);
            this.panel_main.TabIndex = 1;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.ClientSize = new System.Drawing.Size(1018, 589);
            this.Controls.Add(this.panel_main);
            this.Controls.Add(this.toolStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel辅助工具";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripDropDownButton toolStripDropDownButton_file;
        private System.Windows.Forms.ToolStripMenuItem tsmi_exit;
        private System.Windows.Forms.ToolStripDropDownButton toolStripDropDownButton_view;
        private System.Windows.Forms.ToolStripMenuItem tsmi_combine;
        private System.Windows.Forms.Panel panel_main;
    }
}