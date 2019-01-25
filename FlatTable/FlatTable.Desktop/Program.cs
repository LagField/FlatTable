using System;
using Eto.Forms;

namespace FlatTable.Desktop
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
			new Application(Eto.Platform.Detect).Run(new FlatTableForm());
        }
    }
}