using System;
using Eto.Forms;
using Eto.Drawing;

namespace FlatTable.Desktop
{
	class Program
	{
		[STAThread]
		static void Main(string[] args)
		{
//			new Application(Eto.Platform.Detect).Run(new MainForm());
			new Application(Eto.Platform.Detect).Run(new FlatTableForm());
		}
	}
}