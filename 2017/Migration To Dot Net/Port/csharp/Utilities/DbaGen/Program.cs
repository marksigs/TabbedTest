using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace DbaGen
{
	static class Program
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			string[] arguments = Environment.GetCommandLineArgs();

			string root = null;

			if (arguments.Length > 1)
			{
				root = arguments[1];
			}
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new DbaGenForm(root));
		}
	}
}
