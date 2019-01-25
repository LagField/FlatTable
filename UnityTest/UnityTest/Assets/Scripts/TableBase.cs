using System;

namespace FlatTable
{
    public abstract class TableBase : IDisposable
    {
        public abstract string FileName { get; }
        public abstract void Dispose();
        public abstract void Decode(byte[] bytes);
    }
}