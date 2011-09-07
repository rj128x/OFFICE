using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace Office
{
	static class Program
	{
		[DllImport("kernel32.dll")]
		static extern bool AttachConsole(int dwProcessId);
		private const int ATTACH_PARENT_PROCESS = -1;
		/// <summary>
		/// Главная точка входа для приложения.
		/// </summary>
		[STAThread]
		static void Main() {
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);

			/*// Attach to the parent process via AttachConsole SDK call
			AttachConsole(ATTACH_PARENT_PROCESS);
			Console.WriteLine("This is from the main program");*/

			Application.Run(new MenuForm());

			
		}
	}
}
