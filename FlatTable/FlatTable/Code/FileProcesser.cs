using System.Diagnostics;

namespace FlatTable
{
    public class FileProcesser
    {
        public void Process(string[] checkedFilePath)
        {
            if (checkedFilePath == null || checkedFilePath.Length == 0)
            {
                return;
            }

            for (int i = 0; i < checkedFilePath.Length; i++)
            {
                string filePath = checkedFilePath[i];
                
                
            }
        }
    }
}