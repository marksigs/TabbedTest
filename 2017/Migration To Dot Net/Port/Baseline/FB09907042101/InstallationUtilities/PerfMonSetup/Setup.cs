using System;
using Vertex.Fsd.Omiga.Core;


namespace Vertex.Fsd.Omiga.PerfMonSetup
{
	/// <summary>
	/// Summary description for Setup.
	/// </summary>
	class Setup
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main(string[] args)
		{
			bool install = args.Length == 0 || (args.Length > 0 && args[0] == "-i");
			bool uninstall = args.Length > 0 && args[0] == "-u";

			Console.WriteLine("Performance Monitor Counters setup.");

			try
			{
				if (install)
				{
					Console.WriteLine("Installing omExperianWrapper performance counters.");
					Vertex.Fsd.Omiga.omExperianWrapper.Performance.Install();
					Console.WriteLine("Installed omExperianWrapper performance counters.");

 					Console.WriteLine("Installing omGemini performance counters.");
					Vertex.Fsd.Omiga.Gemini.Performance.Install();
					Console.WriteLine("Installed omGemini performance counters.");

					Console.WriteLine("Installing omGeminiPack performance counters.");
					Vertex.Fsd.Omiga.Gemini.Packs.Performance.Install();
					Console.WriteLine("Installed omGeminiPack performance counters.");

					Console.WriteLine("Installing omGeminiWS performance counters.");
					Vertex.Fsd.Omiga.Gemini.Web.Performance.Install();
					Console.WriteLine("Installed omGeminiWS performance counters.");
				}
				else if (uninstall)
				{
					Console.WriteLine("Uninstalling omExperianWrapper performance counters.");
					Vertex.Fsd.Omiga.omExperianWrapper.Performance.Uninstall();
					Console.WriteLine("Uninstalled omExperianWrapper performance counters.");

					Console.WriteLine("Uninstalling omGemini performance counters.");
					Vertex.Fsd.Omiga.Gemini.Performance.Uninstall();
					Console.WriteLine("Uninstalled omGemini performance counters.");

					Console.WriteLine("Uninstalling omGeminiPack performance counters.");
					Vertex.Fsd.Omiga.Gemini.Packs.Performance.Uninstall();
					Console.WriteLine("Uninstalled omGeminiPack performance counters.");

					Console.WriteLine("Uninstalling omGeminiWS performance counters.");
					Vertex.Fsd.Omiga.Gemini.Web.Performance.Uninstall();
					Console.WriteLine("Uninstalled omGeminiWS performance counters.");
				}
			}
			catch (OmigaException ex)
			{
				Console.WriteLine("Error: " + ex.FullMessage);
			}
			catch (Exception ex)
			{
				Console.WriteLine("Error: " + ex.Message);
			}
		}
	}
}
