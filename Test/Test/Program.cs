using System.Collections.Generic;
using System.Diagnostics;
using FlatTable;

namespace Test
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            TableLoader.fileLoadPath = @"D:\CSProjects\FlatTable\FlatTable\Test\BinaryFile";
            TableLoader.Load<TestTable>();
            TableLoader.Load<AnotherTestTable>();
            
            if (AnotherTestTable.ins == null)
            {
                return;
            }
            List<AnotherTestTable.Value> valueList = AnotherTestTable.ins.list;
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