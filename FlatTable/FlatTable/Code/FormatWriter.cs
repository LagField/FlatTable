using System;
using System.Text;

namespace FlatTable
{
    public class FormatWriter
    {
        private StringBuilder sb;
        private int indent;

        public FormatWriter()
        {
            sb = new StringBuilder(1024);
            indent = 0;
        }

        public void AddIndent(int value)
        {
            indent += value;
            indent = Math.Max(0, indent);
        }

        public void WriteLine(string line)
        {
            WriteIndent();
            sb.AppendLine(line);
        }

        public void BreakLine(int lineCount = 1)
        {
            sb.AppendLine("");
        }

        public void WriteComment(string comment)
        {
            WriteIndent();
            sb.Append("//");
            sb.Append(comment);
        }

        /// <summary>
        /// 写入一个{换行，indent +4
        /// </summary>
        public void BeginBlock()
        {
            WriteLine("{");
            AddIndent(4);
        }

        public void EndBlock()
        {
            AddIndent(-4);
            WriteLine("}");
        }

        public override string ToString()
        {
            if (sb == null)
            {
                return String.Empty;
            }

            return sb.ToString();
        }

        private void WriteIndent()
        {
            for (int i = 0; i < indent; i++)
            {
                sb.Append(" ");
            }
        }
    }
}