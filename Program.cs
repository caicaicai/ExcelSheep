using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelSheep.Controller;

namespace ExcelSheep
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Directory.SetCurrentDirectory(Application.StartupPath);

            Console.WriteLine(Directory.GetCurrentDirectory());

            ExcelSheepController excelSheepController = new ExcelSheepController();

            Application.Run(new ExcelSheep(excelSheepController));
        }
    }
}
