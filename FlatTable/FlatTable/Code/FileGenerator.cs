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
        private static Dictionary<string, Action<BinaryWriter, string>> typeBinaryWriterDic =
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
                            if (string.IsNullOrEmpty(valueString))
                            {
                                writer.Write(0);
                            }
                            else
                            {
                                throw new WriteFileException {errorMsg = $"无法解析 int.值: {valueString}"};
                            }
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
                            if (string.IsNullOrEmpty(valueString))
                            {
                                writer.Write((short) 0);
                            }
                            else
                            {
                                throw new WriteFileException {errorMsg = $"无法解析 short.值: {valueString}"};
                            }
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
                            if (string.IsNullOrEmpty(valueString))
                            {
                                writer.Write((float) 0);
                            }
                            else
                            {
                                throw new WriteFileException {errorMsg = $"无法解析 float.值: {valueString}"};
                            }
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
                            if (string.IsNullOrEmpty(valueString))
                            {
                                writer.Write(false);
                            }
                            else
                            {
                                throw new WriteFileException {errorMsg = $"无法解析 bool.值: {valueString}"};
                            }
                        }
                    }
                },
                {
                    "string", (writer, valueString) =>
                    {
                        byte[] bytes = Encoding.UTF8.GetBytes(valueString);
                        ushort byteLength = (ushort) bytes.Length;
                        writer.Write(byteLength);
                        writer.Write(bytes);
                    }
                },
            };

        private static Dictionary<string, string> bitConvertMethodNameDic = new Dictionary<string, string>
        {
            {"int", "ToInt32"},
            {"short", "ToInt16"},
            {"float", "ToSingle"},
            {"bool", "ToBoolean"},
        };

        private static Dictionary<string, int> sizeOfTypeDic = new Dictionary<string, int>
        {
            {"int", sizeof(int)},
            {"short", sizeof(short)},
            {"float", sizeof(float)},
            {"bool", sizeof(bool)},
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

            try
            {
                WriteBinaryFile(rowDatas, fileName);
                WriteCSharpFile(rowDatas, fileName);
            }
            catch (WriteFileException writeFileException)
            {
                Debug.WriteLine(writeFileException.errorMsg);
                MessageBox.Show($"写文件错误:\n{writeFileException.errorMsg}", MessageBoxType.Error);
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
                MessageBox.Show($"写文件未知错误:\n{e.Message}", MessageBoxType.Error);
                throw;
            }
        }

        private void WriteBinaryFile(ExcelRowData[] rowDatas, string fileName)
        {
            //二进制文件按照行列顺序，直接写到文件里就行
            using (BinaryWriter bw = new BinaryWriter(File.Open(AppData.BinaryFileFolderPath + $"/{fileName}Table.bytes",
                FileMode.OpenOrCreate)))
            {
                for (int i = 0; i < rowDatas.Length; i++)
                {
                    ExcelRowData rowData = rowDatas[i];
                    for (int j = 0; j < rowData.rowCellDatas.Length; j++)
                    {
                        RowCellData cellData = rowData.rowCellDatas[j];
                        if (typeBinaryWriterDic.TryGetValue(cellData.typeName, out Action<BinaryWriter, string> writeHandler))
                        {
                            try
                            {
                                writeHandler(bw, cellData.value);
                            }
                            catch (WriteFileException e)
                            {
                                bw.Dispose();
                                throw;
                            }
                        }
                    }
                }
            }
        }

        private void WriteCSharpFile(ExcelRowData[] rowDatas, string fileName)
        {
            FormatWriter fw = new FormatWriter();

            WriteUsing(fw);

            fw.WriteLine("namespace FlatTable");
            fw.BeginBlock();

            fw.WriteLine($"public class {fileName}Table : TableBase");

            fw.BeginBlock();
            WriteValueClass(fw, rowDatas[0]);
            fw.EndBlock();

            fw.BreakLine();

            WriteClassFields(fw, fileName);
            fw.BreakLine();

            //写该表读取的方式
            WriteDecodeFunction(fw, rowDatas);

            WriteDisposeCode(fw);

            fw.EndBlock();

            fw.EndBlock();

            File.WriteAllText($"{AppData.CSharpFolderPath}/{fileName}Table.cs", fw.ToString());
        }

        private void WriteUsing(FormatWriter fw)
        {
            fw.WriteLine("using System;");
            fw.WriteLine("using System.Collections.Generic;");
            fw.BreakLine();
        }

        private void WriteValueClass(FormatWriter fw, ExcelRowData firstRowData)
        {
            //只需要第一行数据，就知道所有数据如何读取
            //写该表Value的结构
            fw.WriteLine("public class Value");
            fw.BeginBlock();

            Dictionary<string, int> arraySizeDictionary = new Dictionary<string, int>();
            for (int i = 0; i < firstRowData.rowCellDatas.Length; i++)
            {
                RowCellData cellData = firstRowData.rowCellDatas[i];

                if (!cellData.isArray)
                {
                    string typeName = cellData.typeName;
                    string fieldName = cellData.fieldName;
                    fw.WriteLine($"public {typeName} {fieldName};");
                }
                else
                {
                    string fieldName = cellData.arrayFieldNameWithoutIndex;
                    if (arraySizeDictionary.TryGetValue(fieldName, out int size))
                    {
                        arraySizeDictionary[fieldName] = size + 1;
                    }
                    else
                    {
                        arraySizeDictionary.Add(fieldName, 1);
                    }
                }
            }

            HashSet<string> writedArrayNameSet = new HashSet<string>();
            for (int i = 0; i < firstRowData.rowCellDatas.Length; i++)
            {
                RowCellData cellData = firstRowData.rowCellDatas[i];
                if (cellData.isArray)
                {
                    string fieldName = cellData.arrayFieldNameWithoutIndex;
                    if (writedArrayNameSet.Contains(fieldName))
                    {
                        continue;
                    }

                    string typeName = cellData.typeName;
                    if (arraySizeDictionary.TryGetValue(fieldName, out int size))
                    {
                        fw.WriteLine($"public {typeName}[] {fieldName} = new {typeName}[{size}];");
                        writedArrayNameSet.Add(fieldName);
                    }
                    else
                    {
                        throw new WriteFileException {errorMsg = $"写文件发生错误，arraySizeDictionary中没有定义{fieldName}的大小."};
                    }
                }
            }
        }

        private void WriteClassFields(FormatWriter fw, string fileName)
        {
            fw.WriteLine($"public static {fileName}Table ins;");
            fw.WriteLine("public List<Value> list = new List<Value>();");
            fw.WriteLine("public Dictionary<int,Value> map = new Dictionary<int, Value>();");
            fw.WriteLine("public override string FileName");
            fw.BeginBlock();
            fw.WriteLine($"get {{ return \"{fileName}Table\"; }}");
            fw.EndBlock();
        }

        private void WriteDecodeFunction(FormatWriter fw, ExcelRowData[] rowDatas)
        {
            fw.WriteLine("public override void Decode(byte[] bytes)");
            fw.BeginBlock();
            fw.WriteLine("ins = this;");
            fw.WriteLine("int readingPosition = 0;");
            fw.WriteLine("ushort stringByteLength = 0;");
            int rowCount = rowDatas.Length;
            fw.WriteLine($"for (int i = 0; i < {rowCount}; i++)");
            fw.BeginBlock();

            ExcelRowData rowData = rowDatas[0];
            fw.WriteLine("Value v = new Value();");

            RowCellData[] cellDatas = rowData.rowCellDatas;
            for (int j = 0; j < cellDatas.Length; j++)
            {
                RowCellData cellData = cellDatas[j];

                if (!cellData.isArray)
                {
                    WriteNonArrayTypeDecode(fw, cellData.fieldName, cellData.typeName);
                }
            }

            //写数组的decode代码
            int cellIndex = 0;
            while (true)
            {
                if (cellIndex >= cellDatas.Length)
                {
                    break;
                }

                RowCellData cellData = cellDatas[cellIndex];
                if (!cellData.isArray)
                {
                    cellIndex++;
                    continue;
                }

                if (cellIndex == cellDatas.Length - 1)
                {
                    WriteArrayTypeDecode(fw, cellData.arrayFieldNameWithoutIndex, cellData.typeName, 1);
                    break;
                }

                //因为数据是有序的，数组全部排在最后，且按照顺序和index排序，所以往后数一下当前数组的长度就行
                string currentArrayFieldName = cellData.arrayFieldNameWithoutIndex;
                int arrayLength = 1;
                for (int j = cellIndex + 1; j < cellDatas.Length; j++)
                {
                    RowCellData compareCellData = cellDatas[j];
                    if (compareCellData.arrayFieldNameWithoutIndex != currentArrayFieldName)
                    {
                        break;
                    }

                    arrayLength++;
                }

                WriteArrayTypeDecode(fw, cellData.arrayFieldNameWithoutIndex, cellData.typeName, arrayLength);
                cellIndex += arrayLength;
            }

            fw.WriteLine("list.Add(v);");
            fw.WriteLine("map.Add(v.id, v);");
            fw.EndBlock();
            fw.EndBlock();
        }

        private void WriteNonArrayTypeDecode(FormatWriter fw, string fieldName, string typeName)
        {
            if (typeName != "string")
            {
                fw.WriteLine($"v.{fieldName} = BitConverter.{bitConvertMethodNameDic[typeName]}(bytes, readingPosition);");
                fw.WriteLine($"readingPosition += {sizeOfTypeDic[typeName]};");
            }
            else
            {
                fw.WriteLine("stringByteLength = BitConverter.ToUInt16(bytes, readingPosition);");
                fw.WriteLine("readingPosition += 2;");
                fw.WriteLine($"v.{fieldName} = System.Text.Encoding.UTF8.GetString(bytes, readingPosition, stringByteLength);");
                fw.WriteLine("readingPosition += stringByteLength;");
            }
        }

        private void WriteArrayTypeDecode(FormatWriter fw, string fieldNameWithoutIndexName, string typeName, int arrayLength)
        {
            fw.WriteLine($"for (int j = 0; j < {arrayLength}; j++)");
            fw.BeginBlock();

            if (typeName != "string")
            {
                fw.WriteLine(
                    $"v.{fieldNameWithoutIndexName}[j] = BitConverter.{bitConvertMethodNameDic[typeName]}(bytes, readingPosition);");
                fw.WriteLine($"readingPosition += {sizeOfTypeDic[typeName]};");
            }
            else
            {
                fw.WriteLine("stringByteLength = BitConverter.ToUInt16(bytes, readingPosition);");
                fw.WriteLine("readingPosition += 2;");
                fw.WriteLine(
                    $"v.{fieldNameWithoutIndexName}[j] = System.Text.Encoding.UTF8.GetString(bytes, readingPosition, stringByteLength);");
                fw.WriteLine("readingPosition += stringByteLength;");
            }

            fw.EndBlock();
        }

        private void WriteDisposeCode(FormatWriter fw)
        {
            fw.WriteLine("public override void Dispose()");
            fw.BeginBlock();
            fw.WriteLine("if(list != null)");
            fw.BeginBlock();
            fw.WriteLine("list.Clear();");
            fw.EndBlock();
            fw.WriteLine("list = null;");

            fw.WriteLine("if(map != null)");
            fw.BeginBlock();
            fw.WriteLine("map.Clear();");
            fw.EndBlock();
            fw.WriteLine("map = null;");

            fw.WriteLine("ins = null;");
            fw.EndBlock();
        }
    }

    public class WriteFileException : Exception
    {
        public string errorMsg;
    }
}