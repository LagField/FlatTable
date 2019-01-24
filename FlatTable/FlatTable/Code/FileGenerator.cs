using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Eto.Forms;

namespace FlatTable
{
    /// <summary>
    /// 负责生成二进制文件和对应的解码c#代码
    /// </summary>
    public class FileGenerator
    {
        private static Dictionary<string, Action<BinaryWriter, string>> valueTypeWriterDictionary =
            new Dictionary<string, Action<BinaryWriter, string>>
            {
                {
                    "int", (writer, valueString) =>
                    {
                        int value = int.Parse(valueString);
                        writer.Write(value);
                    }
                },
                {
                    "short", (writer, valueString) =>
                    {
                        short value = short.Parse(valueString);
                        writer.Write(value);
                    }
                },
                {
                    "float", (writer, valueString) =>
                    {
                        float value = float.Parse(valueString);
                        writer.Write(value);
                    }
                },
                {
                    "bool", (writer, valueString) =>
                    {
                        bool value = bool.Parse(valueString);
                        writer.Write(value);
                    }
                },
                {
                    "string", (writer, valueString) =>
                    {
                        byte[] bytes = Encoding.UTF8.GetBytes(valueString);
                        int byteLength = bytes.Length;
                        writer.Write(byteLength);
                        writer.Write(bytes);
                    }
                },
            };

        public void GenerateFile(ExcelRowData[] rowDatas, string fileName)
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

            using (BinaryWriter bw =
                new BinaryWriter(File.Open(AppData.BinaryFileFolderPath + $"/{fileName}.bytes", FileMode.OpenOrCreate)))
            {
                for (int i = 0; i < rowDatas.Length; i++)
                {
                    ExcelRowData rowData = rowDatas[i];
                    for (int j = 0; j < rowData.rowCellDatas.Length; j++)
                    {
                        RowCellData cellData = rowData.rowCellDatas[j];
                        WriteCellBinary(bw, cellData);
                    }
                }
            }
        }

        private void WriteCellBinary(BinaryWriter bw, RowCellData cellData)
        {
        }
    }
}