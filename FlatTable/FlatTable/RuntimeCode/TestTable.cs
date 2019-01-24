using System;
using System.Collections.Generic;

namespace FlatTable
{
    public class TestTable
    {
        public class Value
        {
            public int id;
            public string hero_name;
            public float speed;
            public int damage;
            public short local_id;
            public bool is_lock;
            public int[] skin_id = new int[3];
            public string[] resource = new string[5];
        }

        public List<Value> list = new List<Value>();
        public Dictionary<int,Value> map = new Dictionary<int, Value>();
        public const string FileName = "Test";

        public void Decode(byte[] bytes)
        {
            int readingPosition = 0;
            ushort stringByteLength = 0;
            for (int i = 0; i < 3; i++)
            {
                Value v = new Value();
                v.id = BitConverter.ToInt32(bytes, readingPosition);
                readingPosition += 4;
                stringByteLength = BitConverter.ToUInt16(bytes, readingPosition);
                readingPosition += 2;
                v.hero_name = System.Text.Encoding.UTF8.GetString(bytes, readingPosition, stringByteLength);
                readingPosition += stringByteLength;
                v.speed = BitConverter.ToSingle(bytes, readingPosition);
                readingPosition += 4;
                v.damage = BitConverter.ToInt32(bytes, readingPosition);
                readingPosition += 4;
                v.local_id = BitConverter.ToInt16(bytes, readingPosition);
                readingPosition += 2;
                v.is_lock = BitConverter.ToBoolean(bytes, readingPosition);
                readingPosition += 1;
                for (int j = 0; j < 3; j++)
                {
                    v.skin_id[j] = BitConverter.ToInt32(bytes, readingPosition);
                    readingPosition += 4;
                }
                for (int j = 0; j < 5; j++)
                {
                    stringByteLength = BitConverter.ToUInt16(bytes, readingPosition);
                    readingPosition += 2;
                    v.resource[j] = System.Text.Encoding.UTF8.GetString(bytes, readingPosition, stringByteLength);
                    readingPosition += stringByteLength;
                }
                list.Add(v);
                map.Add(v.id, v);
            }
        }
    }
}
