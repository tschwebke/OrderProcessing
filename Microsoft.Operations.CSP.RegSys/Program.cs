using System;
using System.Diagnostics;
using System.Reflection;
using System.Security.Permissions;
using System.ServiceProcess;

namespace Microsoft.Operations.CSP.RegSys
{
    internal class Program
    {
        [SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.ControlAppDomain)]
        private static void Main(string[] args)
        {
            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.UnhandledException += new UnhandledExceptionEventHandler(MyHandler);

            if (args.Length > 0)
            {
                if (args[0].Trim().ToLower().Contains("/i")) // Install service
                { System.Configuration.Install.ManagedInstallerClass.InstallHelper(new string[] { "/i", Assembly.GetExecutingAssembly().Location }); }
                else if (args[0].Trim().ToLower().Contains("/u")) // Uninstall service
                { System.Configuration.Install.ManagedInstallerClass.InstallHelper(new string[] { "/u", Assembly.GetExecutingAssembly().Location }); }
            }
            else // normal running
            {
                ServiceBase[] servicesToRun = new ServiceBase[]
                {
                    new Service1()
                };
                servicesToRun.Run(args);
            }
        }

        private static void MyHandler(object sender, UnhandledExceptionEventArgs args)
        {
            Exception e = (Exception)args.ExceptionObject;

            EventLog eventLog1 = new EventLog();
            eventLog1.Source = "Application";
            eventLog1.WriteEntry(e.Message + e.Source, EventLogEntryType.Error);
        }
    }
}