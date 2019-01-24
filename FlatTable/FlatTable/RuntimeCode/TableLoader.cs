using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace FlatTable
{
    public class TableLoader
    {
        private static byte[] cacheBytes;
        public static byte[] GetByteArray(int length)
        {
            if (cacheBytes == null)
            {
                cacheBytes = new byte[1024];
            }

            if (length > cacheBytes.Length)
            {
                Array.Resize(ref cacheBytes,length);
            }

            return cacheBytes;
        }

        public static void Test()
        {
            const string filePath = @"F:\U3dProjects\FlatTable\Test\BinaryFile\Test.bytes";
            byte[] bytes = File.ReadAllBytes(filePath);
            TestTable tt = new TestTable();
            tt.Decode(bytes);
            List<TestTable.Value> valueList = tt.list;
            for (int i = 0; i < valueList.Count; i++)
            {
                Debug.WriteLine(valueList[i].id);
                Debug.WriteLine(valueList[i].hero_name);
                Debug.WriteLine(valueList[i].speed);
                Debug.WriteLine(valueList[i].damage);
                Debug.WriteLine(valueList[i].is_lock);

                for (int j = 0; j < valueList[i].resource.Length; j++)
                {
                    Debug.WriteLine(valueList[i].resource[j]);
                }
            }
        }
    }
}