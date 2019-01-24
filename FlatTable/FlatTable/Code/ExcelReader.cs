using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using Eto.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace FlatTable
{
    public class ExcelReader
    {
        private Application excelApp;
        static HashSet<string> supportTypeSet = new HashSet<string> {"int", "short", "float", "bool", "string"};

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

            try
            {
                GetValideRowAndColumnCount(valueTable, out int rowCount, out int columnCount);
                //如果少于4行(fieldname,include/exclude,typename,data)或者少于2列(id,otherdata)
                if (rowCount <= 4 || columnCount <= 1)
                {
                    throw new ParseCellException {exceptionMsg = $"文档行列数量过少，无有效数据.行{rowCount} 列{columnCount}"};
                }

                //第一列必须是id且类型必须是int
                if (valueTable[1, 1].ToString() != "id" || valueTable[3, 1].ToString() != "int")
                {
                    throw new ParseCellException {exceptionMsg = "文档第一列项目第一格名字不是id，或者类型不是int"};
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

                return rowDatas;
            }
            catch (ParseCellException parseException)
            {
                Console.WriteLine("parse exception: " + parseException.exceptionMsg);
                MessageBox.Show($"解析发生错误\n{parseException.exceptionMsg}");
            }
            catch (Exception e)
            {
                Console.WriteLine("occur error: " + e);
                MessageBox.Show($"发生错误\n{e}");
            }
            finally
            {
                CloseExcelFile(usedRange, workSheet, workBook);
            }

            return null;
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

            bool isInclude = valueTable[2, columnIndex + 1].ToString() != ColumnType.Exclude;
            if (!isInclude)
            {
                return null;
            }

            string valueString = valueObject == null ? "" : valueObject.ToString();
            string filedName = valueTable[1, columnIndex + 1].ToString();
            string typeName = valueTable[3, columnIndex + 1].ToString();
            bool isArrayType = typeName.Contains("[]");

            if (isArrayType)
            {
                typeName = typeName.Replace("[]", "");
            }

            if (!supportTypeSet.Contains(typeName))
            {
                //提示
                throw new ParseCellException {exceptionMsg = $"行{rowIndex + 1} 列{columnIndex + 1} 的类型信息填写错误: {typeName}"};
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
        public const string Exclude = "Exclude";
        public const string Include = "Include";
    }

    public class ParseCellException : Exception
    {
        public string exceptionMsg;
    }
}