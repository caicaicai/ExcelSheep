using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using ExcelSheep.Controller;
using ExcelSheep.Model;


namespace ExcelSheep
{
    public partial class AddExport : Form
    {

        private ExcelSheepController controller;
        private int _oldSelectedIndex = -1;
        //this is a copy of exportString that we are working on
        private ExportConfig _modifiedConfiguration;

        #region 构造和析构函数
        public AddExport(ExcelSheepController controller)
        {
            InitializeComponent();

            //控制器
            this.controller = controller;
            //载入筛选数据的栏目
            this.ColumnsCBox.DataSource = controller.ColumnNames;
            //筛选数据操作符初始化
            this.ComparisonOperatorsCBox.SelectedIndex = 0;

            LoadCurrentConfiguration();
        }

        ~AddExport()
        {

        } 
        #endregion

        #region 根据用户点击的逻辑拼接符，进行条件拼接
        private void AndBtn_Click(object sender, EventArgs ea)
        {
            try
            {
                string booleanOperator = " AND";
                this.SpliceFilterStr(booleanOperator);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void OrBtn_Click(object sender, EventArgs ea)
        {
            try
            {
                string booleanOperator = " OR";
                this.SpliceFilterStr(booleanOperator);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void SpliceFilterStr(string booleanOperator)
        {
            if (this.filterResult.Text == "")
            {
                booleanOperator = "";
            }
            if (ComparisonOperatorsCBox.SelectedItem.ToString() == "IN" || ComparisonOperatorsCBox.SelectedItem.ToString() == "NOT IN")
            {
                this.filterResult.Text += booleanOperator + " " + ColumnsCBox.SelectedItem.ToString() + " " + ComparisonOperatorsCBox.SelectedItem.ToString() + " (" + LiteralsTextBox.Text + ")";
            }
            else
            {
                this.filterResult.Text += booleanOperator + " " + ColumnsCBox.SelectedItem.ToString() + " " + ComparisonOperatorsCBox.SelectedItem.ToString() + " '" + LiteralsTextBox.Text + "'";
            }
        } 
        #endregion

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string url = "http://www.csharp-examples.net/dataview-rowfilter/";
            System.Diagnostics.Process.Start(url);
        }

        private void addBtn_Click(object sender, EventArgs e)
        {
            if (!SaveOldSelectedExportStrings())
            {
                return;
            }

            ExportItem ei = ExportConfig.GetDefaultServer();
            _modifiedConfiguration.configs.Add(ei);
            LoadExportStrings(_modifiedConfiguration);
            exportItemListBox.SelectedIndex = _modifiedConfiguration.configs.Count - 1;
            _oldSelectedIndex = exportItemListBox.SelectedIndex;
        }

        private void ComparisonOperatorsCBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComparisonOperatorsCBox.SelectedItem.ToString() == "IN" || ComparisonOperatorsCBox.SelectedItem.ToString() == "NOT IN")
            {
                LiteralsTextBox.Text = "'John', 'Jim'";
            }
            else
            {
                LiteralsTextBox.Text = "";
            }
        }

        private bool SaveOldSelectedExportStrings()
        {
            try
            {
                if (_oldSelectedIndex == -1 || _oldSelectedIndex >= _modifiedConfiguration.configs.Count)
                {
                    return true;
                }

                if (exportItemName.Text.Length == 0)
                {
                    MessageBox.Show("导出名称不能为空!");
                    return false;
                }

                ExportItem ei = new ExportItem { 
                    name = exportItemName.Text,
                    queryStr = filterResult.Text
                };
                _modifiedConfiguration.configs[_oldSelectedIndex] = ei;

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return false;
        }

        private void LoadExportStrings(ExportConfig configuration)
        {

            exportItemListBox.Items.Clear();

            foreach (ExportItem exportItem in configuration.configs)
            {
                exportItemListBox.Items.Add(exportItem.FriendlyName());
            }

        }

        private void LoadSelectedExportString()
        {

            if (exportItemListBox.SelectedIndex >= 0 && exportItemListBox.SelectedIndex < _modifiedConfiguration.configs.Count)
            {
                ExportItem ei = _modifiedConfiguration.configs[exportItemListBox.SelectedIndex];

                exportItemName.Text = ei.name;
                filterResult.Text = ei.queryStr;
                exportItemListBox.Visible = true;
            }
            else
            {
                exportItemListBox.Visible = false;
            }
        }

        private void LoadCurrentConfiguration()
        {
            _modifiedConfiguration = controller.GetConfiguration();
            LoadExportStrings(_modifiedConfiguration);
            _oldSelectedIndex = 0;
            exportItemListBox.SelectedIndex = 0;
            LoadSelectedExportString();
        }

        private void AddExport_FormClosed(object sender, FormClosedEventArgs e)
        {
            _modifiedConfiguration.configs.Clear();
        }

        private void exportItemListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_oldSelectedIndex == exportItemListBox.SelectedIndex)
            {
                // we are moving back to oldSelectedIndex or doing a force move
                return;
            }

            //尝试保存现有的配置
            if (!SaveOldSelectedExportStrings())
            {
                // why this won't cause stack overflow?
                exportItemListBox.SelectedIndex = _oldSelectedIndex;
                return;
            }

            //将当前选择的位置保存
            _oldSelectedIndex = exportItemListBox.SelectedIndex;
            //刷新配置列表
            LoadExportStrings(_modifiedConfiguration);
            //将刚才保存的位置重新渲染
            exportItemListBox.SelectedIndex = _oldSelectedIndex;
            //将选择的配置载入编辑器
            LoadSelectedExportString();
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            if (!SaveOldSelectedExportStrings())
            {
                return;
            }
            if (_modifiedConfiguration.configs.Count == 0)
            {
                MessageBox.Show("请至少添加一个导出");
                return;
            }
            controller.saveExport(_modifiedConfiguration.configs);
            this.Close();
        }

        private void removeBtn_Click(object sender, EventArgs e)
        {
            _oldSelectedIndex = exportItemListBox.SelectedIndex;
            if (_oldSelectedIndex >= 0 && _oldSelectedIndex < _modifiedConfiguration.configs.Count)
            {
                _modifiedConfiguration.configs.RemoveAt(_oldSelectedIndex);
            }
            if (_oldSelectedIndex >= _modifiedConfiguration.configs.Count)
            {
                // can be -1
                _oldSelectedIndex = _modifiedConfiguration.configs.Count - 1;
            }
            exportItemListBox.SelectedIndex = _oldSelectedIndex;
            LoadExportStrings(_modifiedConfiguration);
            exportItemListBox.SelectedIndex = _oldSelectedIndex;
            LoadSelectedExportString();
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
