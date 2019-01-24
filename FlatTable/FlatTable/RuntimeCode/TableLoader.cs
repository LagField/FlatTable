using System;

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

        public void Test()
        {
            const string filePath = ""
        }
    }
}