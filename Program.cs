using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OrderPDF2CSV
{
	static class Program
	{
		/// <summary>
		/// Der Haupteinstiegspunkt für die Anwendung.
		/// </summary>
		static void Main()
		{
#if (!DEBUG)
			ServiceBase[] ServicesToRun;
			ServicesToRun = new ServiceBase[] 
			{ 
				new ServiceMain() 
			};
			ServiceBase.Run(ServicesToRun);
#else
			ConvertPDF2CV4DSP mainObj = new ConvertPDF2CV4DSP();
			string[] args = new string[0];
			mainObj.OnDebugStart(args);
			System.Console.Write("zum Beenden Return-Taste betätigen...");
			System.Console.Read();  // warten, bis Taste betätigt wird

			System.Console.Write("Programm wird beendet...");
			mainObj.OnDebugStop();
#endif
		}
	}
}
