using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ExcelSheep.Utils;
using System.Data;
using System.Windows.Forms;
using System.ComponentModel;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using ExcelSheep.Model;


namespace ExcelSheep.Controller
{

    #region 定义新的事件类型
    public class LogEventArgs : EventArgs
    {
        public string loginfo { get; set; }
    }
    public delegate void LogEventHandler(object sender, LogEventArgs e);
    #endregion

    
    public class ExcelSheepController
    {
        #region Attributes
        //基础数据的文件路径
        private String sourceDataPath;
        //基础数据生成的基础表
        private DataTable baseTable;

        //数据筛数据表
        private DataTable filteredTable;

        //需要导出的数据表数组
        public Dictionary<string, DataTable> exportTables;


        private ExportConfig _config;
        #endregion

        #region 注册处理filterString的change事件
        public event EventHandler FilterStringChanged; 
        protected void OnFilterStringChangedChanged(EventArgs e)
        {
            EventHandler handler = FilterStringChanged;
            if (handler != null)
                handler(this, e);
        }
        #endregion

        #region 注册处理FilteredTable的change事件
        public event EventHandler FilteredTableChanged;
        protected void OnFilteredTableChangedChanged(EventArgs e)
        {
            EventHandler handler = FilteredTableChanged;
            if (handler != null)
                handler(this, e);
        }
        #endregion

        #region 注册日志事件
        public void log(string str)
        {
            LogEventArgs e = new LogEventArgs();
            e.loginfo = str;
            LogEventHandler hander = AttemptToLog;
            if (hander != null)
                hander(this, e);
        }

        public event LogEventHandler AttemptToLog; 
        #endregion

        #region 构造函数
        public ExcelSheepController()
        {
            //初始化
            this.sourceDataPath = "";
            this._config = ExportConfig.Load();
            this.exportTables = new Dictionary<string, DataTable>();
            this.filteredTable = new DataTable();
        } 
        #endregion

        #region GET / SET

        public string SourceDataPath
        {
            get
            {
                return this.sourceDataPath;
            }
            set
            {
                if (value != this.sourceDataPath)
                {
                    this.sourceDataPath = value;
                }
            }
        }

        public DataTable BaseTable
        {
            get
            {
                return this.baseTable;
            }
            set
            {
                if (value != this.baseTable)
                {
                    this.baseTable = value;
                }
            }

        }

        public string FilterString
        {
            get { return this._config.filterString; }
            set
            {
                if (value != this._config.filterString)
                {
                    this._config.filterString = value;
                    //保存
                    ExportConfig.Save(_config);
                    this._config = ExportConfig.Load();

                }
            }
        }

        public DataTable FilteredTable
        {
            get
            {
                return this.filteredTable;
            }
            set
            {
                if(value != this.filteredTable)
                {
                    this.filteredTable = value;
                    OnFilteredTableChangedChanged(EventArgs.Empty);
                }
            }
        }

        public string[] ColumnNames
        {
            get
            {
                if (this.BaseTable == null)
                {
                    throw new Exception("没有基础数据，无法获取数据的栏目列表");
                }
                return this.BaseTable.Columns.Cast<DataColumn>()
                                 .Select(x => x.ColumnName)
                                 .ToArray(); 
            }
        }

        //always return copy
        public ExportConfig GetConfiguration()
        {
            return ExportConfig.Load();
        }

        #endregion

        public void generateFilterData()
        {
            if (_config.filterString.Length > 0)
            {
                log(String.Format("开始数据过滤筛选。"));
                FilteredTable = DataTableSeleter(this.BaseTable, _config.filterString);
                
            }
            else
            {
                log("筛选条件为空，数据将不进行筛选。");
                if( filteredTable != null && filteredTable.Rows.Count > 0 )
                {
                    filteredTable.Clear();
                }
            }
        }

        public void generateExportData()
        {
            Console.WriteLine(_config.configs.Count);
            if(_config.configs.Count > 0)
            {
                DataTable tempSource;
                DataTable tempTagert;
                if (filteredTable == null || filteredTable.Rows.Count == 0)
                {
                    tempSource = BaseTable.Copy();
                    log(String.Format("筛选数据为空或者不可用，将使用原始数据生成导出"));
                }
                else
                {
                    tempSource = FilteredTable.Copy();
                    log(String.Format("筛选数据存在，将使用筛选数据生成导出。"));
                }

                log(String.Format("正在清理旧的导出数据。"));
                if(exportTables != null && exportTables.Count > 0)
                {
                    exportTables.Clear();
                }
                foreach(ExportItem ei in _config.configs)
                {
                    log(String.Format("正在导出数据项：{0}", ei.name));
                    tempTagert = DataTableSeleter(tempSource, ei.queryStr);
                    exportTables.Add(ei.name, tempTagert);
                }
                log(String.Format("导出完成。"));
            }
            else
            {
                throw new Exception("没有配置导出项，导出数据不可用。");
            }
        }

        #region 导出数据
        public void exportFiles()
        {
            if (this.exportTables == null || exportTables.Count == 0)
            {
                throw new Exception("没有任何数据可以用于导出！");
            }
            foreach (var item in exportTables)
            {
                log(String.Format("正在导出数据：{0}", item.Key));
                saveFile(item.Value, item.Key);

            }
        } 
        #endregion

        public void saveExport(List<ExportItem> exports)
        {
            _config.configs = exports;
            ExportConfig.Save(_config);
            this._config = ExportConfig.Load();
        }

        public List<string> getDataList()
        {
            List<string> dl = new List<string>{"源数据","筛选数据"};
            if(exportTables.Count > 0)
            {
                dl.AddRange(exportTables.Keys);
            }
            return dl;
        }

        public DataTable getDataByName(string name)
        {
            if(name == "源数据")
            {
                return BaseTable;
            }
            else if(name == "筛选数据")
            {
                return filteredTable;
            }
            else
            {
                if(exportTables.ContainsKey(name))
                {
                    return exportTables[name];
                }
                else
                {
                    throw new KeyNotFoundException("尝试访问不存在的导出数据");
                }
            }
        }

        public void loadData(bool chkHasHeader)
        {
            try
            {
                this.log("从选择的excel:"+ this.sourceDataPath +" 读取数据...");
                using (ExcelPackage pck = new ExcelPackage())
                {
                    using (var stream = File.OpenRead(this.sourceDataPath))
                    {
                        pck.Load(stream);
                    }

                    ExcelWorksheet ws = pck.Workbook.Worksheets.First();
                    this.baseTable = WorksheetToDataTable(ws, chkHasHeader);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Import failed. Original error: " + ex.Message);
                //MessageBox.Show("Import failed. Original error: " + ex.Message);
            }
        }

        public void handelData(DataGridView dataGridView)
        {
            var saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel File (*.xlsx)|*.xlsx";
            saveFileDialog1.FilterIndex = 1;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileInfo file = new FileInfo(saveFileDialog1.FileName);
                    if (file.Exists)
                    {
                        file.Delete();
                    }

                    using (ExcelPackage pck = new ExcelPackage(file))
                    {
                        ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Sheet1");
                        ws.Cells["A1"].LoadFromDataTable(((DataTable)baseTable), true);
                        ws.Cells.AutoFitColumns();

                        using (ExcelRange rng = ws.Cells[1, 1, 1, dataGridView.Columns.Count])
                        {
                            rng.Style.Font.Bold = true;
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(79, 129, 189));
                            rng.Style.Font.Color.SetColor(System.Drawing.Color.White);
                        }

                        pck.Save();
                    }

                    Console.WriteLine(string.Format("Excel file \"{0}\" generated successfully.", file.Name));
                }
                catch (Exception ex)
                {
                    throw new Exception("Failed to export to Excel. Original error: " + ex.Message);
                }
            }
        }

        public void saveFile(DataTable dt, string name)
        {
            try
            {
                FileInfo file = new FileInfo(Directory.GetCurrentDirectory() + "\\" + name + ".xlsx");
                if (file.Exists)
                {
                    file.Delete();
                }

                using (ExcelPackage pck = new ExcelPackage(file))
                {
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Sheet1");
                    ws.Cells["A1"].LoadFromDataTable(((DataTable)dt), true);
                    ws.Cells.AutoFitColumns();

                    using (ExcelRange rng = ws.Cells[1, 1, 1, dt.Columns.Count])
                    {
                        rng.Style.Font.Bold = true;
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(79, 129, 189));
                        rng.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    }

                    pck.Save();
                }

                Console.WriteLine(string.Format("Excel file \"{0}\" generated successfully.", file.Name));
                log(string.Format("Excel file \"{0}\" generated successfully.", file.Name));
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to export to Excel. Original error: " + ex.Message);
            }


        }

        private DataTable WorksheetToDataTable(ExcelWorksheet ws, bool hasHeader = true)
        {
            DataTable dt = new DataTable(ws.Name);
            int totalCols = ws.Dimension.End.Column;
            int totalRows = ws.Dimension.End.Row;
            int startRow = hasHeader ? 2 : 1;
            ExcelRange wsRow;
            DataRow dr;
            foreach (var firstRowCell in ws.Cells[1, 1, 1, totalCols])
            {
                dt.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column{0}", firstRowCell.Start.Column));
            }

            for (int rowNum = startRow; rowNum <= totalRows; rowNum++)
            {
                wsRow = ws.Cells[rowNum, 1, rowNum, totalCols];
                dr = dt.NewRow();
                foreach (var cell in wsRow)
                {
                    dr[cell.Start.Column - 1] = cell.Text;
                }

                dt.Rows.Add(dr);
            }

            return dt;
        }

        public DataTable DataTableSeleter(DataTable sourceTable, string selectStr)
        {

            this.log("正在筛选数据...");
            DataRow [] dr =  sourceTable.Select(selectStr);
            int count = dr.Count();
            if(count > 0)
            {
                DataTable dt = dr.CopyToDataTable();
                this.log(string.Format("筛选到数据 {0} 条...", count));
                return dt;
            }
            else
            {
                this.log("没有筛选到任何数据!");
                return new DataTable();
            }
            
        }

    }
}
