using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace MBDRPC.Helpers
{
	/// <summary>
	/// A class that checks if an application is currently running.
	/// </summary>
	public class RunningAppChecker
	{
		/// <summary>
		/// Checks if the specified application is currently running.
		/// </summary>
		/// <param name="appName">The name of the application to check.</param>
		/// <returns>
		/// <c>true</c> if the application is running; otherwise, <c>false</c>
		/// </returns>
		public static bool IsAppRunning(string appName)
		{
			return Process.GetProcesses().Any(process => process.ProcessName.Equals(appName, StringComparison.OrdinalIgnoreCase));
		}

        /// <summary>
        /// Checks if the specified application is currently running and not exited.
        /// </summary>
        /// <param name="appName">The name of the application to check.</param>
        /// <returns>
        /// <c>true</c> if the application is running and not exited; otherwise, <c>false</c>
        /// </returns>
        public static bool IsAppRunningAndNotExited(string appName)
        {
            return Process.GetProcesses()
                          .Any( process => process.ProcessName.Equals( appName , StringComparison.OrdinalIgnoreCase ) &&
                                           ! process.HasExited );
        }


        /// <summary>
        /// Checks if any of the specified apps/processes are currently running.
        /// </summary>
        /// <param name="appNames">Names of the applications/processes to check.</param>
        public static bool IsOneAppRunning(params string[] appNames)
		{
			return appNames.Any(appName => IsAppRunning(appName));
		}


        /// <summary>
        /// Checks if any of the specified apps/processes ending with the specified name are currently running.
        /// </summary>
        /// <param name="appNames">Names of the applications/processes to check.</param>
        public static bool IsOneAppRunningEndingWith(params string[] appNames)
        {
            return appNames.Any(appName => IsAppRunning(appName) && appName.EndsWith(appName));
        }


        /// <summary>
        /// Checks if any of containing the specified name of the apps/processes are currently running.
        /// </summary>
        /// <param name="name">Name to check.</param>
        public static bool IsAnyAppRunningWithNameContaining( string name )
        {
            return Process.GetProcesses().Any(process => process.ProcessName.Contains(name));
        }


        /// <summary>
        /// Gets the name of the first process name containing one of the specified names.
        /// </summary>
        /// <param name="names">Names to check.</param>
        public static string GetFirstProcessNameContaining( params string[] names )
        {
            return Process
                  .GetProcesses()
                  .FirstOrDefault(process => names.Any(name => process.ProcessName.Contains(name)))?.ProcessName;
        }


        /// <summary>
        /// Gets the name of the first process name starting with one of the specified names.
        /// </summary>
        /// <param name="names">Names to check.</param>
        public static string GetFirstProcessNameStartingWith(params string[] names)
        {
            return Process
                  .GetProcesses()
                  .FirstOrDefault(process => names.Any(name => process.ProcessName.StartsWith(name)))?.ProcessName;
        }


        /// <summary>
        /// Gets the name of the first process name ending with one of the specified names.
        /// </summary>
        /// <param name="names">Names to check.</param>
        public static string GetFirstProcessNameEndingWith(params string[] names)
        {
            return Process
                  .GetProcesses()
                  .FirstOrDefault(process => names.Any(name => process.ProcessName.EndsWith(name)))?.ProcessName;
        }


        public static string GetProcessPublisher(int processId)
        {
            try
            {
                Process         process         = Process.GetProcessById(processId);
                string          mainModulePath  = process.MainModule.FileName;
                FileVersionInfo fileVersionInfo = FileVersionInfo.GetVersionInfo(mainModulePath);
                string          publisher       = fileVersionInfo.CompanyName;
                return publisher;
            }
            catch (Exception ex)
            {
                // Handle the case where the process's main module path could not be retrieved
                return null;
            }
        }

        /// <summary>
        /// Retrieves the start time of the specified process.
        /// </summary>
        /// <param name="processName">The name of the process.</param>
        public static DateTime GetProcessStartTime(string processName)
		{
			var process = Process.GetProcessesByName(processName);
			return process[0].StartTime;
		}




		/// <summary>
		/// Checks if the SSMS (SQL Server Management Studio) application is currently running.
		/// </summary>
		public static bool IsSsmsRunning()
		{
			return Process.GetProcesses().Any(process => process.ProcessName.Equals("ssms", StringComparison.OrdinalIgnoreCase));
		}

		/// <summary>
		/// Checks if the PgAdmin application is currently running.
		/// </summary>
		public static bool IsPgAdminRunning()
		{
			return Process.GetProcesses().Any(process => process.ProcessName.Equals("pgadmin4", StringComparison.OrdinalIgnoreCase));
		}

        /// <summary>
        /// Checks if the Microsoft Word application is currently running.
        /// </summary>
        public bool IsMicrosoftWordRunning()
        {
			return Process.GetProcesses().Any(process => process.ProcessName.Equals("winword", StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Checks if the Microsoft Excel application is currently running.
        /// </summary>
        public bool IsMicrosoftExcelRunning()
        {
            return Process.GetProcesses().Any(process => process.ProcessName.Equals("excel", StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Checks if the Microsoft PowerPoint application is currently running.
        /// </summary>
        public bool IsMicrosoftPowerPointRunning()
        {
            return Process.GetProcesses().Any(process => process.ProcessName.Equals("powerpnt", StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Checks if the Microsoft Whiteboard application is currently running.
        /// </summary>
        public bool IsMicrosoftWhiteboardRunning()
        {
            return Process.GetProcesses().Any(process => process.ProcessName.Equals("MicrosoftWhiteboard", StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Checks if the Microsoft OneDrive application is currently running.
        /// </summary>
        public bool IsMicrosoftOneDriveRunning()
        {
            return Process.GetProcesses().Any(process => process.ProcessName.Equals("onedrive", StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Checks if the Microsoft Outlook application is currently running.
        /// </summary>
        public bool IsMicrosoftOutlookRunning()
        {
            return Process.GetProcesses().Any(process => process.ProcessName.Equals("OUTLOOK", StringComparison.OrdinalIgnoreCase));
        }


        /// <summary>
        /// Checks if the Microsoft Publisher application is currently running.
        /// </summary>
        public bool IsMicrosoftPublisherRunning()
        {
            return Process.GetProcesses().Any(process => process.ProcessName.Equals("mspub", StringComparison.OrdinalIgnoreCase));
        }

    }
}