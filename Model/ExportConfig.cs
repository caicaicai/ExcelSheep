using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSheep.Model
{
    [Serializable]
    public class ExportConfig
    {
        public List<ExportItem> configs;

        public string filterString;

        private static string CONFIG_FILE = "excelSheep-config.json";

        public static ExportConfig Load()
        {
            try
            {
                string configContent = File.ReadAllText(CONFIG_FILE);
                ExportConfig config = SimpleJson.SimpleJson.DeserializeObject<ExportConfig>(configContent, new JsonSerializerStrategy());
                return config;
            }
            catch (Exception e)
            {
                if (!(e is FileNotFoundException))
                {
                    Console.WriteLine(e);
                }
                return new ExportConfig
                {
                    configs = new List<ExportItem>()
                    {
                        GetDefaultServer()
                    },
                    filterString = ""
                };
            }
        }

        public static void Save(ExportConfig config)
        {

            try
            {
                using (StreamWriter sw = new StreamWriter(File.Open(CONFIG_FILE, FileMode.Create)))
                {
                    string jsonString = SimpleJson.SimpleJson.SerializeObject(config);
                    sw.Write(jsonString);
                    sw.Flush();
                }
            }
            catch (IOException e)
            {
                Console.Error.WriteLine(e);
            }
        }

        public static ExportItem GetDefaultServer()
        {
            return new ExportItem
            {
                name = "未配置的导出",
                queryStr = ""
            };
        }


        private class JsonSerializerStrategy : SimpleJson.PocoJsonSerializerStrategy
        {
            // convert string to int
            public override object DeserializeObject(object value, Type type)
            {
                if (type == typeof(Int32) && value.GetType() == typeof(string))
                {
                    return Int32.Parse(value.ToString());
                }
                return base.DeserializeObject(value, type);
            }
        }


    }
}
