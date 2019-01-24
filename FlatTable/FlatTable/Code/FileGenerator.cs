using System.IO;
using Eto.Forms;

namespace FlatTable
{
    /// <summary>
    /// 负责生成二进制文件和对应的解码c#代码
    /// </summary>
    public class FileGenerator
    {
        public void GenerateFile(ExcelRowData[] rowDatas,string fileName)
        {
            if (!Directory.Exists(AppData.BinaryFileFolderPath))
            {
                MessageBox.Show("二进制文件路径不存在");
                return;
            }

            if (!Directory.Exists(AppData.CSharpFolderPath))
            {
                MessageBox.Show("C#文件路径不存在");
                return;
            }

            using (BinaryWriter bw = new BinaryWriter(File.Open(AppData.BinaryFileFolderPath + $"/{fileName}", FileMode.OpenOrCreate)))
            {
                
            }
        }
    }
}