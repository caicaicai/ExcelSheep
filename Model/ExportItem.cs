using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSheep.Model
{
    [Serializable]
    public class ExportItem
    {
        public string name;
        public string queryStr;

        public string FriendlyName()
        {
            if (string.IsNullOrEmpty(name))
            {
                return "新的导出";
            }
            if (string.IsNullOrEmpty(queryStr))
            {
                return name + ":(无筛选条件)";
            }
            else
            {
                return name + ":(" + queryStr + ")";
            }
        }

        public ExportItem()
        {

        }
    }
}
