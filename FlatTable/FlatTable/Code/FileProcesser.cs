using System.Diagnostics;
using System.IO;
using Eto.Forms;

namespace FlatTable
{
    public class FileProcesser
    {
        private ExcelReader excelReader;
        private FileGenerator fileGenerator;

        public FileProcesser()
        {
            excelReader = new ExcelReader();
            fileGenerator = new FileGenerator();
        }

        public void Process(string[] checkedFilePath)
        {
            if (checkedFilePath == null || checkedFilePath.Length == 0)
            {
                return;
            }

            for (int i = 0; i < checkedFilePath.Length; i++)
            {
                string filePath = checkedFilePath[i];

                ExcelRowData[] rowDatas = excelReader.ReadExcelDatas(filePath);
                if (rowDatas == null)
                {
                    return;
                }

                fileGenerator.GenerateFile(rowDatas, Path.GetFileNameWithoutExtension(checkedFilePath[i]));
            }

            MessageBox.Show("完成", "成功");
        }
    }
}