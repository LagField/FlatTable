using System;
using System.Diagnostics;
using System.IO;
using Newtonsoft.Json;

namespace FlatTable
{
    public static class AppData
    {
        private static Config config;

        private static string ConfigFilePath
        {
            get { return AppDomain.CurrentDomain.BaseDirectory + "config.json"; }
        }

        public static void Init()
        {
            if (File.Exists(ConfigFilePath))
            {
                string configString = File.ReadAllText(ConfigFilePath);
                config = JsonConvert.DeserializeObject<Config>(configString);
            }
            else
            {
                config = new Config();
                config.excelFolderPath = AppDomain.CurrentDomain.BaseDirectory;
            }
        }

        public static string ExcelFolderPath
        {
            set
            {
                bool needRefreshConfigFile = config.excelFolderPath != value;

                config.excelFolderPath = value;
                if (needRefreshConfigFile)
                {
                    SaveConfigFile();
                }
            }
            get { return config.excelFolderPath; }
        }

        /// <summary>
        /// 根据当前内容重新写入配置文件
        /// </summary>
        public static void SaveConfigFile()
        {
            string json = JsonConvert.SerializeObject(config);
            File.WriteAllText(ConfigFilePath, json);
        }
    }

    [Serializable]
    public class Config
    {
        public string excelFolderPath;
    }
}