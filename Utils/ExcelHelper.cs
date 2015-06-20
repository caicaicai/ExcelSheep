using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using ExcelSheep.Utils;

using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ExcelSheep.Utils
{
    class ExcelHelper
    {
        public Excel.Application excelApp;
        public Excel._Workbook excelWB;
        public Excel._Worksheet excelSheet;
        //Excel.Range excelRng;


        #region 生成空的ExcelHelper对象
        /// <summary>
        /// 生成空的ExcelHepler对象
        /// </summary>
        /// <returns>ExcelHelper对象</returns>
        public ExcelHelper()
        {
            //Start Excel and get Application object.
            excelApp = new Excel.Application();
            excelApp.Visible = true;

            //Get a new workbook.
            excelWB = (Excel._Workbook)(excelApp.Workbooks.Add(Missing.Value));
            excelSheet = (Excel._Worksheet)excelWB.ActiveSheet;
        } 
        #endregion

        #region 根据EXCEL路径生成ExcelHelper对象
        /// <summary>
        /// 根据EXCEL路径生成ExcelHepler对象
        /// </summary>
        /// <param name="ExcelFilePath">EXCEL文件的路径</param>
        /// <returns>ExcelHelper对象</returns>
        public ExcelHelper(string ExcelFilePath)
        {
            //Start Excel and get Application object.
            excelApp = new Excel.Application();
            //excelApp.Visible = true;

            //Get a  workbook.
            excelWB = (Excel._Workbook)(excelApp.Workbooks.Open(ExcelFilePath));
            excelSheet = (Excel._Worksheet)excelWB.ActiveSheet;
        }
        #endregion

        public Dictionary<string, int> getSheetCount()
        {
            Dictionary<string,int> count = new Dictionary<string,int>();

            int rowsCount = excelSheet.UsedRange.Rows.Count;
            int colsCount = excelSheet.UsedRange.Columns.Count;
            Console.WriteLine("活动的行数为：{0},列数为：{1}", rowsCount,colsCount);

            count.Add("rows",rowsCount);
            count.Add("cols",colsCount);

            return count;
        }

        public void setFilter()
        {

            excelSheet.Cells.AutoFilter(1, "<10", Excel.XlAutoFilterOperator.xlOr, Type.Missing, Type.Missing);

            Excel.Range filteredRange = excelSheet.UsedRange.SpecialCells(
                               Excel.XlCellType.xlCellTypeVisible,
                               Type.Missing);
            //Start Excel and get Application object.
            Excel.Application App = new Excel.Application();

            //Get a new workbook.
            Excel._Workbook WB = (Excel._Workbook)(App.Workbooks.Add(Missing.Value));
            Excel._Worksheet Sheet = (Excel._Worksheet)WB.ActiveSheet;

            Sheet.UsedRange.Value = filteredRange.Value;

            Sheet.SaveAs("d:\\haha.xlsx");
        }

        /// <summary>
        /// This method takes DataSet as input paramenter and it exports the same to excel
        /// </summary>
        /// <param name="ds"></param>
        private void ExportDataSetToExcel(DataSet ds)
        {
            //Creae an Excel application instance
            Excel.Application excelApp = new Excel.Application();

            //Create an Excel workbook instance and open it from the predefined location
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(@"E:\Org.xlsx");

            foreach (DataTable table in ds.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;

                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }
            }

            excelWorkBook.Save();
            excelWorkBook.Close();
            excelApp.Quit();

        }

        public DataSet getDataSetFromExcel()
        {

            DataSet ds = new DataSet();
            try
            {
                foreach (Excel._Worksheet objSHT in excelWB.Worksheets)
                {
                    int rows = objSHT.UsedRange.Rows.Count;
                    int cols = objSHT.UsedRange.Columns.Count;
                    DataTable dt = new DataTable();
                    int noofrow = 1;

                    //If 1st Row Contains unique Headers for datatable include this part else remove it
                    //Start
                    for (int c = 1; c <= cols; c++)
                    {
                        string colname = ExcelConvert.ToName(c-1);
                        dt.Columns.Add(colname);
                        //noofrow = 2;
                    }
                    //END

                    for (int r = noofrow; r <= rows; r++)
                    {
                        DataRow dr = dt.NewRow();
                        for (int c = 1; c <= cols; c++)
                        {
                            dr[c - 1] = objSHT.Cells[r, c].Text;
                        }
                        dt.Rows.Add(dr);
                    }
                    ds.Tables.Add(dt);

                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("读取excel到dataset的时候出现异常，{0}", ex.Message);
            }

            return ds;

        }



    }
}
