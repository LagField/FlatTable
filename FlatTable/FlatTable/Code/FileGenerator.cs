using System;
using System.Collections.Generic;
using System.Diagnostics;
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
        private static Dictionary<string, Action<BinaryWriter, string>> typeBinaryWriterDictionary =
            new Dictionary<string, Action<BinaryWriter, string>>
            {
                {
                    "int", (writer, valueString) =>
                    {
                        if (int.TryParse(valueString, out int value))
                        {
                            writer.Write(value);
                        }
                        else
                        {
                            writer.Write(0);
                        }
                    }
                },
                {
                    "short", (writer, valueString) =>
                    {
                        if (short.TryParse(valueString, out short value))
                        {
                            writer.Write(value);
                        }
                        else
                        {
                            writer.Write((short)0);
                        }
                    }
                },
                {
                    "float", (writer, valueString) =>
                    {
                        if (float.TryParse(valueString, out float value))
                        {
                            writer.Write(value);
                        }
                        else
                        {
                            writer.Write((float)0);
                        }
                    }
                },
                {
                    "bool", (writer, valueString) =>
                    {
                        if (bool.TryParse(valueString, out bool value))
                        {
                            writer.Write(value);
                        }
                        else
                        {
                            writer.Write(false);
                        }
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
//            for (int i = 0; i < rowDatas.Length; i++)
//            {
//                Debug.WriteLine(rowDatas[i]);
//            }
//            
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
            if (typeBinaryWriterDictionary.TryGetValue(cellData.typeName, out Action<BinaryWriter, string> handler))
            {
                handler(bw, cellData.value);
            }
        }
    }
}