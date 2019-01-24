using System.Diagnostics;

namespace FlatTable
{
    public class FileProcesser
    {
        private ExcelReader excelReader;

        public FileProcesser()
        {
            excelReader = new ExcelReader();
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
                
                for (int j = 0; j < rowDatas.Length; j++)
                {
                    ExcelRowData rowData = rowDatas[j];
                    Debug.WriteLine(rowData.ToString());
                }
            }
        }
    }
}