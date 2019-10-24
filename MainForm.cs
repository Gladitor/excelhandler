using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelHandler
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void tsmi_combine_Click(object sender, EventArgs e)
        {
            UserControl_combine ucc = new UserControl_combine();
            viewChange(ucc);
        }

        void viewChange(UserControl uc)
        {
            uc.Dock = DockStyle.Fill;
            this.panel_main.Controls.Clear();
            this.panel_main.Controls.Add(uc);
        }

        private void tsmi_exit_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            UserControl_combine ucc = new UserControl_combine();
            viewChange(ucc);
        }
    }
}
