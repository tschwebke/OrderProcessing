using System;
using System.Collections.Generic;
using System.Reflection;
using System.ServiceProcess;

namespace Microsoft.Operations
{
    /// <summary>
    /// Common class for handling everything related to Console+Service hybrid mode.
    /// </summary>
    public static class ServiceBaseExtensions
    {
        /// <summary>
        /// Runs a given collection of services and hosts them in either Console Application or
        /// Windows Service, depending on how the application is executed (in interactive mode or not).
        ///
        /// The output type of the application must be set to Console: Project properties -&gt;
        /// Output type: Console Application.
        ///
        /// create references to:
        /// 1. System.Configuration.Install
        /// 2. System.Management
        /// 3. System.ServiceProcess
        /// 4. The core program which contains the extension for ServiceBase.
        ///
        /// copy from an existing Windows Service:
        ///
        /// Project Installer.cs Service1.cs
        ///
        /// modify:
        ///
        /// 1. Program.cs, to include stuff from Main function
        /// 2. Namespace, from the one copied, to reflect the current. NOTE: Search ALL files which
        ///    were copied.
        /// 3. Modify the designer to change the service name (in TWO Places)
        ///
        /// Code Credits for idea:
        /// ref: http://stackoverflow.com/questions/125964/easier-way-to-start-debugging-a-windows-service-in-c-sharp
        /// ref: http://einaregilsson.com/run-windows-service-as-a-console-program/
        /// ref: http://pastebin.com/F0fhhG2R
        ///
        /// NOTE: During debugging sessions you may see a few orphaned 'conhost.exe' floating around
        ///       (found on all OS's from Windows 7 to 10). Tests show that these don't appear when
        /// running as a service.
        /// </summary>
        /// <param name="servicesToRun">Array of services to run.</param>
        /// <param name="args">Arguments to be passed to the services (for interactive console)</param>
        public static void Run(this ServiceBase[] servicesToRun, string[] args)
        {
            if (!Environment.UserInteractive)
            {
                ServiceBase.Run(servicesToRun); // headless operation
                return;
            }

            Console.WriteLine("Running services in interactive mode... {0}", DateTime.Now);
            Console.WriteLine();

            CallServiceBaseMethod(servicesToRun, "OnStart", new object[] { args }, "Starting");

            Console.WriteLine("Press any key to stop the services...");
            Console.ReadKey();

            CallServiceBaseMethod(servicesToRun, "OnStop", null, "Stopping");

            Console.WriteLine();
            Console.WriteLine("Press any key to exit... ");
            Console.ReadKey();
        }

        /// <summary>
        /// Invokes the internal method for the requested service, using reflection.
        /// </summary>
        private static void CallServiceBaseMethod(IEnumerable<ServiceBase> services, string methodName, object[] methodArgs, string consoleMessage)
        {
            MethodInfo onStopMethod = typeof(ServiceBase).GetMethod(methodName, BindingFlags.Instance | BindingFlags.NonPublic);
            foreach (ServiceBase service in services)
            {
                Console.Write("{0} '{1}' ... ", consoleMessage, service.ServiceName);
                onStopMethod.Invoke(service, methodArgs);
                Console.Write("OK");
            }
            Console.WriteLine();
        }
    }
}

/*
SAMPLE USAGE:
*/

// using BullionBars;
//namespace MyNamespace
//{
//    static class Program
//    {
//        static void Main(string[] args)
//        {
//ServiceBase[] servicesToRun = new ServiceBase[]
//{
//    new Service1()
//};
//servicesToRun.Run(args);
//        }
//    }
//}