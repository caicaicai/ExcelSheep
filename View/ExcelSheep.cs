using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelSheep.Controller;

using ExcelSheep.Utils;

namespace ExcelSheep
{
    
    public partial class ExcelSheep : Form
    {
        private ExcelSheepController controller;


        public ExcelSheep(ExcelSheepController controller)
        {

            InitializeComponent();
            this.controller = controller;
            this.controller.AttemptToLog += new LogEventHandler(logInfo);
            loadDataList();
        }

        private void ExcelSheep_Load(object sender, EventArgs e)
        {
            logBox.AppendText("[双击以清除日志]" + Environment.NewLine);
        }

        private void selectBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = false;
            fileDialog.Title = "请选择文件";
            fileDialog.Filter = "Excel Files|*.xlsx";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                //拿到选择的文件名，显示在选择器后面
                fileName.Text = fileDialog.SafeFileName;

                //将拿到的文件路径存入控制器
                this.controller.SourceDataPath = fileDialog.FileName;
                this.logInfo("选择了文件:" + fileDialog.FileName);

                this.controller.loadData(chkHasHeader.Checked);

                dataGridView1.DataSource = this.controller.BaseTable;

                this.logInfo("数据读取完成...");
            }
        }


        private void exportBtn_Click(object sender, EventArgs e)
        {


            if (dataGridView1.DataSource == null)
            {
                MessageBox.Show("缺少待操作数据，请先选择需要处理的数据！");
                return;
            }

            this.controller.handelData(dataGridView1);

            
        }

        public void logInfo(string Info)
        {
            logBox.AppendText(DateTime.Now.ToString("HH:mm:ss  ") + Info);
            logBox.AppendText(Environment.NewLine);
            logBox.ScrollToCaret();

            System.IO.StreamWriter sw = System.IO.File.AppendText(Directory.GetCurrentDirectory() + "/log.txt");
            sw.WriteLine(DateTime.Now.ToString("HH:mm:ss  ") + Info);
            sw.Close();

        }

        private void logInfo(object sender, LogEventArgs e)
        {
            logInfo(e.loginfo);
        }

        private void logBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            logBox.Clear();
        }

        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
            //Generate Row Number
            if (dataGridView1.DataSource != null)
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    row.HeaderCell.Value = string.Format("{0}", row.Index + 1);
                }
            }
        }

        private void addFilterBtn_Click(object sender, EventArgs e)
        {
            try
            {
                AddDataFilter adf = new AddDataFilter(this.controller);
                adf.ShowDialog();
            }
            catch (Exception ex)
            {
                //throw new Exception("Import failed. Original error: " + ex.Message);
                MessageBox.Show(ex.Message);
            }
        }

        private void addExportBtn_Click(object sender, EventArgs e)
        {
            try
            {
                AddExport ae = new AddExport(this.controller);
                ae.ShowDialog();
            }
            catch (Exception ex)
            {
                //throw new Exception("Import failed. Original error: " + ex.Message);
                MessageBox.Show(ex.Message);
            }
        }

        private void loadDataList()
        {
            dataListCBox.Items.Clear();

            foreach(string dl in controller.getDataList())
            {
                dataListCBox.Items.Add(dl);
            }
            dataListCBox.SelectedIndex = 0;
        }

        private void dataListCBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.DataSource = this.controller.getDataByName(dataListCBox.Text);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void filterBtn_Click(object sender, EventArgs e)
        {
            try
            {
                //generate filter data
                this.controller.generateFilterData();
                //generate export data
                this.controller.generateExportData();
                //reload datalist
                loadDataList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        private void handelDataBtn_Click(object sender, EventArgs e)
        {
            try
            {
                this.controller.exportFiles();
            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

    }
}
