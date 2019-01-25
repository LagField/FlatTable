using System;
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
                cacheBytes = new byte[length];
            }

            return cacheBytes;
        }

        public static string fileLoadPath;
        public static Func<string, byte[]> customLoader;

        public static void Load<T>() where T : TableBase, new()
        {
            T t = new T();

            if (customLoader != null)
            {
                byte[] bytes = customLoader(t.FileName);
                t.Decode(bytes);
                return;
            }
            
                string fileName = t.FileName + ".bytes";
                string loadPath = Path.Combine(fileLoadPath, fileName);
                if (!File.Exists(loadPath))
                {
                    Debug.WriteLine(string.Format("file not found: {0}", loadPath));
                    t.Dispose();
                    return;
                }

                using (FileStream fs = File.OpenRead(loadPath))
                {
                    int byteLength = (int) fs.Length;
                    fs.Position = 0;
                    byte[] bytes = GetByteArray(byteLength);
                    fs.Read(bytes, 0, byteLength);
                    t.Decode(bytes);
                }
        }
    }
}