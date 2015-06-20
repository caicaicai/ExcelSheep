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
namespace ExcelSheep
{
    public partial class AddDataFilter : Form
    {

        private ExcelSheepController controller;

        public AddDataFilter(ExcelSheepController controller)
        {
            InitializeComponent();
            this.controller = controller;
            filterResult.Text = controller.FilterString;
            ColumnsCBox.DataSource = controller.ColumnNames;
            //筛选数据操作符初始化
            this.ComparisonOperatorsCBox.SelectedIndex = 0;

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string url = "http://www.csharp-examples.net/dataview-rowfilter/";
            System.Diagnostics.Process.Start(url);
        }

        private void confimBtn_Click(object sender, EventArgs e)
        {
            this.controller.FilterString = filterResult.Text;
            this.Close();
        }

        private void filterResultReset_Click(object sender, EventArgs e)
        {
            this.filterResult.Text = "";
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            this.controller.log("取消编辑数据筛选");
            this.Close();
        }

        private void AndBtn_Click(object sender, EventArgs ea)
        {
            try
            {
                string booleanOperator = " AND";
                this.SpliceFilterStr(booleanOperator);
            }
            catch(Exception ex)
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
            catch(Exception ex)
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

        private void ComparisonOperatorsCBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComparisonOperatorsCBox.SelectedItem.ToString() == "IN" || ComparisonOperatorsCBox.SelectedItem.ToString() == "NOT IN")
            {
                LiteralsTextBox.Text = "'John', 'Jim', 'Tom'";
            }
            else
            {
                LiteralsTextBox.Text = "";
            }
        }

    }
}
