using System.Text.RegularExpressions;

namespace Migracao
{
	internal static class Program
	{
		/// <summary>
		///  The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			var j = "REC - 2099/01 - CELESTE RODRIGUES FERREIRA";
            var asdf = Regex.Split(j, "REC - ")[1].Split(' ')[0];
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
			Application.Run(new Form1());
		}
	}
}