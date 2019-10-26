using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections.Specialized;

namespace ExcelHandler
{
    public partial class UserControl_combine : UserControl
    {
        public UserControl_combine()
        {
            InitializeComponent();
        }

        //合并全部
        private void btn_pickfile_Click(object sender, EventArgs e)
        {
            bool horizontal = this.checkBox_horizontal.Checked;
            bool allInOneSheet = this.checkBox_allInOneSheet.Checked;
            bool combineBySheet = this.checkBox_combinBySheet.Checked;
            bool appointSheet = this.checkBox_appointsheet.Checked;
            bool appointSheetInOne = this.checkBox_allAppointInOneSheet.Checked;
            if(allInOneSheet && !combineBySheet && !appointSheet && !appointSheetInOne)
            {
                if(horizontal)
                {
                    AllInOneSheet_horizontal();
                }
                else
                {
                    AllInOneSheet();
                }
                MessageBox.Show("合并完成!","提示",MessageBoxButtons.OK,MessageBoxIcon.Asterisk);
            }
            else if(combineBySheet && !allInOneSheet && !appointSheet && !appointSheetInOne)
            {
                if (horizontal)
                { }
                else
                {
                    AllCombineBySheet();
                }
                MessageBox.Show("合并完成!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            else if(appointSheet && !allInOneSheet && !combineBySheet && !appointSheetInOne)
            {
                if (horizontal)
                { }
                else
                {
                    CombineAppointSheet();
                }
                MessageBox.Show("合并完成!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            else if(appointSheetInOne && !allInOneSheet && !combineBySheet && !appointSheet)
            {
                if (horizontal)
                { }
                else
                {
                    CombineAppointSheetInOne();
                }
                MessageBox.Show("合并完成!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            else
            {
                MessageBox.Show("请选择一种合并方式！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //所有Sheet合并成1个
        void AllInOneSheet()
        {
            bool firstRowIsHead = true;
            if (!checkBox_firstrowishead.Checked)
            {
                firstRowIsHead = false;
            }
            //System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.OpenFileDialog fileDialog = new System.Windows.Forms.OpenFileDialog();
            fileDialog.Multiselect = true;
            System.Windows.Forms.DialogResult result = fileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                string basestr = string.Empty;
                string targetfilename = string.Empty;
                string[] filenames = fileDialog.FileNames;
                basestr = filenames[0];
                targetfilename = basestr.Substring(0, basestr.LastIndexOf("\\") + 1) + "AllInOneSheet.xlsx";
                int num = 0;
                string targetsheet = "Sheet1";
                for (int i = 0; i < filenames.Length; i++)
                {
                    ExcelEdit ee = new ExcelEdit();
                    ee.Open(filenames[i]);
                    StringCollection sc = ee.ExcelSheetNames(filenames[i]);
                    for (int j = 0; j < sc.Count; j++)
                    {
                        DataTable dt = ExcelUtil.ExcelToDataTable(sc[j].Substring(0, sc[j].Length - 1), firstRowIsHead, filenames[i], false);
                        if (num == 0)
                        {
                            ExcelUtil.DataTableToExcel(dt, targetsheet, true, targetfilename, null);
                        }
                        else
                        {
                            ExcelUtil.appendInfoToFile(targetfilename, targetsheet, dt);
                        }
                        num++;
                    }
                }
            }
        }
        //按Sheet合并所有文件
        void AllCombineBySheet()
        {
            bool firstRowIsHead = true;
            if (!checkBox_firstrowishead.Checked)
            {
                firstRowIsHead = false;
            }
            //System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.OpenFileDialog fileDialog = new System.Windows.Forms.OpenFileDialog();
            fileDialog.Multiselect = true;
            System.Windows.Forms.DialogResult result = fileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                string basestr = string.Empty;
                string targetfilename = string.Empty;
                string[] filenames = fileDialog.FileNames;
                basestr = filenames[0];
                targetfilename = basestr.Substring(0, basestr.LastIndexOf("\\") + 1) + "AllInOneFileBySheet.xlsx";
                int num = 0;
                string targetsheet = "Sheet1";
                for (int i = 0; i < filenames.Length; i++)
                {
                    ExcelEdit ee = new ExcelEdit();
                    ee.Open(filenames[i]);
                    StringCollection sc = ee.ExcelSheetNames(filenames[i]);
                    for (int j = 0; j < sc.Count; j++)
                    {
                        targetsheet = sc[j].Substring(0, sc[j].Length - 1);
                        DataTable dt = ExcelUtil.ExcelToDataTable(sc[j].Substring(0, sc[j].Length - 1), firstRowIsHead, filenames[i], false);
                        if (num == 0)
                        {
                            string[] sheetnames = new string[sc.Count-1];
                            for (int s = 1; s < sc.Count; s++)
                            {
                                sheetnames[s - 1] = sc[s].Substring(0,sc[s].Length - 1);
                            }
                            ExcelUtil.DataTableToExcel(dt, targetsheet, true, targetfilename, sheetnames);
                        }
                        else
                        {
                            //ExcelUtil.appendInfoToFile(targetfilename, targetsheet, dt);
                            ExcelUtil.DataTableToExcel(dt, targetfilename, targetsheet);
                        }
                        num++;
                    }
                    ee.Close();
                }
            }
        }
        //合并指定Sheet到新文件对应Sheet
        void CombineAppointSheet()
        {
            string[] targetSheetNames = this.textBox_appointSheetName.Lines;
            bool firstRowIsHead = true;
            if (!checkBox_firstrowishead.Checked)
            {
                firstRowIsHead = false;
            }
            //System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.OpenFileDialog fileDialog = new System.Windows.Forms.OpenFileDialog();
            fileDialog.Multiselect = true;
            System.Windows.Forms.DialogResult result = fileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                string[] filenames = fileDialog.FileNames;
                string basestr = filenames[0];
                string targetfilename = basestr.Substring(0, basestr.LastIndexOf("\\") + 1) + "AllAppoint.xlsx";
                for (int i = 0; i < filenames.Length; i++)
                {
                    for (int j = 0; j < targetSheetNames.Length; j++)
                    {
                        string targetSheetName = targetSheetNames[j];
                        if(targetSheetName!="")
                        {
                            DataTable dt = ExcelUtil.ExcelToDataTable(targetSheetName, firstRowIsHead, filenames[i], false);
                            ExcelUtil.DataTableToExcel(dt, targetfilename, targetSheetName);
                        }
                    }
                }
            }
        }
        //合并指定Sheet到新文件1个Sheet
        void CombineAppointSheetInOne()
        {
            string[] targetSheetNames = this.textBox_appointSheetName.Lines;
            bool firstRowIsHead = true;
            if (!checkBox_firstrowishead.Checked)
            {
                firstRowIsHead = false;
            }
            //System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.OpenFileDialog fileDialog = new System.Windows.Forms.OpenFileDialog();
            fileDialog.Multiselect = true;
            System.Windows.Forms.DialogResult result = fileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                string targetSheetName = "Sheet1";
                string[] filenames = fileDialog.FileNames;
                string basestr = filenames[0];
                string targetfilename = basestr.Substring(0, basestr.LastIndexOf("\\") + 1) + "AllAppoint.xlsx";
                for (int i = 0; i < filenames.Length; i++)
                {
                    for (int j = 0; j < targetSheetNames.Length; j++)
                    {
                        if(targetSheetNames[j]!="")
                        {
                            DataTable dt = ExcelUtil.ExcelToDataTable(targetSheetNames[j], firstRowIsHead, filenames[i], false);
                            ExcelUtil.DataTableToExcel(dt, targetfilename, targetSheetName);
                        }
                    }
                }
            }
        }
        //所有Sheet合并成1个-horizontal
        void AllInOneSheet_horizontal()
        {
            bool firstRowIsHead = true;
            if (!checkBox_firstrowishead.Checked)
            {
                firstRowIsHead = false;
            }
            System.Windows.Forms.OpenFileDialog fileDialog = new System.Windows.Forms.OpenFileDialog();
            fileDialog.Multiselect = true;
            System.Windows.Forms.DialogResult result = fileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                string basestr = string.Empty;
                string targetfilename = string.Empty;
                string[] filenames = fileDialog.FileNames;
                basestr = filenames[0];
                targetfilename = basestr.Substring(0, basestr.LastIndexOf("\\") + 1) + "AllInOneSheet-horizontal.xlsx";
                string targetsheet = "Sheet1";
                DataTable allDt = new DataTable();
                for (int i = 0; i < filenames.Length; i++)
                {
                    ExcelEdit ee = new ExcelEdit();
                    ee.Open(filenames[i]);
                    StringCollection sc = ee.ExcelSheetNames(filenames[i]);
                    ee.Close();
                    for (int j = 0; j < sc.Count; j++)
                    {
                        DataTable dt = ExcelUtil.ExcelToDataTable(sc[j].Substring(0, sc[j].Length - 1), firstRowIsHead, filenames[i], false);
                        //合并datatable
                        allDt = DataTableHelper.UniteDataTable(allDt, dt,"");
                    }
                }
                ExcelUtil.DataTableToExcel(allDt, targetsheet, true, targetfilename, null);
            }
        }

        private void checkBox_appointsheet_CheckedChanged(object sender, EventArgs e)
        {
            if(((CheckBox)sender).Checked)
            {
                this.textBox_appointSheetName.ReadOnly = false;
            }
            else
            {
                this.textBox_appointSheetName.ReadOnly = true;
            }
        }

        private void checkBox_allAppointInOneSheet_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked)
            {
                this.textBox_appointSheetName.ReadOnly = false;
            }
            else
            {
                this.textBox_appointSheetName.ReadOnly = true;
            }
        }
    }
}
