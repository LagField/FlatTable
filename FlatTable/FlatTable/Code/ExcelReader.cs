using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Eto.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace FlatTable
{
    public class ExcelReader
    {
        private Application excelApp;
        private string currentReadingFileName;
        static HashSet<string> supportTypeSet = new HashSet<string> {"int", "short", "float", "bool", "string"};
        private const string ArrayIndexPattern = @"\[\d+\]";

        public ExcelReader()
        {
            excelApp = new ApplicationClass();
        }

        public ExcelRowData[] ReadExcelDatas(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return null;
            }

            Workbook workBook = excelApp.Workbooks.Open(filePath);
            Worksheet workSheet = (Worksheet) workBook.Worksheets.Item[1];
            Range usedRange = workSheet.UsedRange;
            object[,] valueTable = (object[,]) usedRange.Value2;
            currentReadingFileName = Path.GetFileNameWithoutExtension(filePath);

            try
            {
                GetValideRowAndColumnCount(valueTable, out int rowCount, out int columnCount);
                //如果少于4行(fieldname,include/exclude,typename,data)或者少于2列(id,otherdata)
                if (rowCount <= 4 || columnCount <= 1)
                {
                    throw new ParseExcelException {exceptionMsg = $"文档行列数量过少，无有效数据.行{rowCount} 列{columnCount}"};
                }

                //第一列必须是id且类型必须是int
                if (valueTable[1, 1].ToString() != "id" || valueTable[3, 1].ToString() != "int")
                {
                    throw new ParseExcelException {exceptionMsg = "文档第一列项目第一格名字不是id，或者类型不是int"};
                }

                ExcelRowData[] rowDatas = new ExcelRowData[rowCount - 3];
                //从第四行开始
                for (int i = 3; i < rowCount; i++)
                {
                    ExcelRowData rowData = new ExcelRowData();
                    List<RowCellData> cellDataList = new List<RowCellData>();
                    for (int j = 0; j < columnCount; j++)
                    {
                        RowCellData cellData = ParseCellData(valueTable, i, j);
                        if (cellData == null)
                        {
                            continue;
                        }

                        cellDataList.Add(cellData);
                    }

                    rowData.rowCellDatas = cellDataList.ToArray();
                    rowDatas[i - 3] = rowData;
                }

                DataValidate(rowDatas[0]);

                return rowDatas;
            }
            catch (ParseExcelException parseException)
            {
                Console.WriteLine("parse exception: " + parseException.exceptionMsg);
                MessageBox.Show($"文档{currentReadingFileName}解析发生错误\n{parseException.exceptionMsg}");
            }
            catch (Exception e)
            {
                Console.WriteLine("occur error: " + e);
                MessageBox.Show($"文档{currentReadingFileName}发生未知错误\n{e}");
            }
            finally
            {
                CloseExcelFile(usedRange, workSheet, workBook);
                currentReadingFileName = string.Empty;
            }

            return null;
        }

        /// <summary>
        /// 检查名称类型等是否有错误
        /// </summary>
        /// <param name="firstRow"></param>
        private void DataValidate(ExcelRowData firstRow)
        {
            FieldNameValidate(firstRow);
            ArrayTypeValidate(firstRow);
        }

        /// <summary>
        /// 检查是否有重复的field
        /// </summary>
        /// <param name="firstRow"></param>
        private void FieldNameValidate(ExcelRowData firstRow)
        {
            HashSet<string> occuredFieldNameSet = new HashSet<string>();

            for (int i = 0; i < firstRow.rowCellDatas.Length; i++)
            {
                RowCellData cellData = firstRow.rowCellDatas[i];
                if (cellData.isArray)
                {
                    string fieldName = cellData.fieldName;
                    MatchCollection matchCollection = Regex.Matches(fieldName, ArrayIndexPattern);
                    if (matchCollection.Count != 1)
                    {
                        throw new ParseExcelException
                            {exceptionMsg = $"文档列 {fieldName} 是数组类型，但是它有0个或多个index申明.\n" + "正确格式: variable[0]"};
                    }

                    string fieldNameWithoutIndex = fieldName.Replace(matchCollection[0].Value, "");
                    //数组类型的名字不能和其他普通变量名字一样
                    if (occuredFieldNameSet.Contains(fieldNameWithoutIndex))
                    {
                        throw new ParseExcelException {exceptionMsg = $"文档首行有重复的field: {fieldName}.该数组名和普通field名字有重复.\n"};
                    }
                }
                else
                {
                    string fieldName = cellData.fieldName;
                    if (occuredFieldNameSet.Contains(fieldName))
                    {
                        throw new ParseExcelException {exceptionMsg = $"文档首行有重复的field: {fieldName}.\n"};
                    }

                    occuredFieldNameSet.Add(fieldName);
                }
            }
        }

        /// <summary>
        /// 检查数组类型的列index是否是连续的
        /// </summary>
        private void ArrayTypeValidate(ExcelRowData firstRow)
        {
            Dictionary<string, List<int>> occuredArrayFields = new Dictionary<string, List<int>>();
            for (int i = 0; i < firstRow.rowCellDatas.Length; i++)
            {
                RowCellData cellData = firstRow.rowCellDatas[i];
                if (!cellData.isArray)
                {
                    continue;
                }

                string fieldName = cellData.fieldName;
                MatchCollection matchCollection = Regex.Matches(fieldName, ArrayIndexPattern);
                if (matchCollection.Count != 1)
                {
                    throw new ParseExcelException
                        {exceptionMsg = $"文档列 {fieldName} 是数组类型，但是它有0个或多个index申明.\n" + "正确格式: variable[0]"};
                }

                Match matchResult = matchCollection[0];
                if (!matchResult.Success)
                {
                    throw new ParseExcelException
                        {exceptionMsg = $"文档列 {fieldName} 是数组类型，但是无法解析它的index.\n" + "正确格式: variable[0]"};
                }

                string fieldNameWithoutIndex = fieldName.Replace(matchResult.Value, "");

                string arrayIndexString = matchResult.Value.Substring(1, matchResult.Value.Length - 2);
                if (!int.TryParse(arrayIndexString, out int arrayIndex))
                {
                    throw new ParseExcelException {exceptionMsg = $"文档列 {fieldName} 是数组类型，但是它的index无法转换为int."};
                }

                if (occuredArrayFields.ContainsKey(fieldNameWithoutIndex))
                {
                    List<int> occuredArrayIndexList = occuredArrayFields[fieldNameWithoutIndex];
                    if (occuredArrayIndexList.Contains(arrayIndex))
                    {
                        throw new ParseExcelException {exceptionMsg = $"文档列 {fieldName} 是数组类型，但是有重复的index: {arrayIndex}."};
                    }

                    occuredArrayIndexList.Add(arrayIndex);
                }
                else
                {
                    List<int> indexList = new List<int> {arrayIndex};
                    occuredArrayFields.Add(fieldNameWithoutIndex, indexList);
                }
            }

            foreach (var kv in occuredArrayFields)
            {
                string fieldNameWithoutIndex = kv.Key;
                List<int> arrayIndexList = kv.Value;

                arrayIndexList.Sort((a, b) => a - b);
                for (int i = 0; i < arrayIndexList.Count; i++)
                {
                    if (arrayIndexList[i] != i)
                    {
                        throw new ParseExcelException
                            {exceptionMsg = $"文档列 {fieldNameWithoutIndex} 是数组类型，但是它的index不是从0开始，或者不是连续的."};
                    }
                }
            }
        }

        private void GetValideRowAndColumnCount(object[,] valueTable, out int rowCount, out int columnCount)
        {
            rowCount = 0;
            columnCount = 0;
            for (int i = 0; i < valueTable.GetLength(1); i++)
            {
                object valueObject = valueTable[1, i + 1];
                if (valueObject == null || string.IsNullOrEmpty(valueObject.ToString()))
                {
                    break;
                }

                columnCount = i + 1;
            }

            for (int i = 0; i < valueTable.GetLength(0); i++)
            {
                object valueObject = valueTable[i + 1, 1];
                if (valueObject == null || string.IsNullOrEmpty(valueObject.ToString()))
                {
                    break;
                }

                rowCount = i + 1;
            }
        }

        private RowCellData ParseCellData(object[,] valueTable, int rowIndex, int columnIndex)
        {
            object valueObject = valueTable[rowIndex + 1, columnIndex + 1];

            bool isInclude = valueTable[2, columnIndex + 1].ToString().ToLower() != ColumnType.Exclude;
            if (!isInclude)
            {
                return null;
            }

            string valueString = valueObject == null ? "" : valueObject.ToString();
            string filedName = valueTable[1, columnIndex + 1].ToString().ToLower();
            string typeName = valueTable[3, columnIndex + 1].ToString().ToLower();
            bool isArrayType = typeName.Contains("[]");

            if (isArrayType)
            {
                typeName = typeName.Replace("[]", "");
            }

            if (!supportTypeSet.Contains(typeName))
            {
                //提示
                throw new ParseExcelException
                {
                    exceptionMsg = $"行{rowIndex + 1} 列{columnIndex + 1} 的类型信息填写错误: {typeName}\n" +
                                   $"目前支持的类型有： {string.Join(",", Array.ConvertAll(supportTypeSet.ToArray(), i => i))},以及他们的数组类型"
                };
            }

            RowCellData cellData = new RowCellData
            {
                value = valueString,
                fieldName = filedName,
                typeName = typeName,
                isArray = isArrayType
            };
            return cellData;
        }

        private void CloseExcelFile(Range usedRange, Worksheet workSheet, Workbook workBook)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(usedRange);
            Marshal.ReleaseComObject(workSheet);

            workBook.Close();
            Marshal.ReleaseComObject(workBook);
        }

        ~ExcelReader()
        {
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
        }
    }

    public class ExcelRowData
    {
        public RowCellData[] rowCellDatas;

        public override string ToString()
        {
            if (rowCellDatas == null || rowCellDatas.Length == 0)
            {
                return "empty row";
            }

            string result = "";
            for (int i = 0; i < rowCellDatas.Length; i++)
            {
                RowCellData cellData = rowCellDatas[i];
                result += cellData + "  ";
            }

            return result;
        }
    }

    public class RowCellData
    {
        public string fieldName;
        public string typeName;
        public string value;
        public bool isArray;

        public override string ToString()
        {
            return $"[field Name: {fieldName}, type name: {typeName}, value: {value}, isArray: {isArray.ToString()}]";
        }
    }

    public struct ColumnType
    {
        public const string Exclude = "exclude";
        public const string Include = "include";
    }

    public class ParseExcelException : Exception
    {
        public string exceptionMsg;
    }
}