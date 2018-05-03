using System;
using System.ComponentModel;
using System.Net;
using System.ServiceProcess;
using System.Timers;

namespace Microsoft.Operations.CSP.RegSys
{
    public partial class Service1 : ServiceBase
    {
        public static string EmailPassword = "L3wn-Z~Xcs36GK;";

        // TODO: Remove passwords dependencies.
        public static string EmailUserIdentityEmail = "svcasfp@microsoft.com";

        public static string ProjectName = "CSP";
        public static string ServerName = "http://vstfpg07:8080/tfs/Operations";

        public BackgroundWorker ProcessWizardEntries, InvoiceSending, ActivationCompleted; // for main Process, allows non-colliding threading

        private static Timer every_minute;

        private ICredentials credentials = CredentialCache.DefaultNetworkCredentials;

        public Service1()
        {
            InitializeComponent();
            ActivationCompleted = new BackgroundWorker(); ActivationCompleted.DoWork += new DoWorkEventHandler(ActivationCompleted_DoWork); ActivationCompleted.WorkerReportsProgress = true;
            InvoiceSending = new BackgroundWorker(); InvoiceSending.DoWork += new DoWorkEventHandler(InvoiceSending_DoWork); InvoiceSending.WorkerReportsProgress = true;
            ProcessWizardEntries = new BackgroundWorker(); ProcessWizardEntries.DoWork += new DoWorkEventHandler(ProcessWizardEntries_DoWork); ProcessWizardEntries.WorkerReportsProgress = true;
        }

        protected override void OnStart(string[] args)
        {
            every_minute = new System.Timers.Timer(1000 * 60 * 1); // Every minute, check inbox for new emails
            every_minute.Elapsed += Timer_ReviewEntries_Tick;
            every_minute.Enabled = true; //

            Timer_ReviewEntries_Tick(null, null);
        }

        protected override void OnStop()
        {
            every_minute.Enabled = false;
        }

        /// <summary>
        /// IMPORTANT NOTE: Although this timer ticks every 1 minute, Each thread has it's own
        /// inbuilt time regulator (so the work won't necessarily be running fully on the interval)
        /// Check the internals of each thread for timing details.
        /// </summary>

        private void Timer_ReviewEntries_Tick(object sender, ElapsedEventArgs e)
        {
            // Private testing
            if (Environment.UserName.ToLower() == "warren" || Environment.UserName.ToLower() == "thads")
            {
                every_minute.Enabled = true;
                if (DateTime.UtcNow.Minute % 2 == 0) // every 2 minutes, look for new orders to process.
                {
                    try
                    {
                        ProcessWizardEntries_InvokeThread();
                    }
                    catch (Exception ex)
                    {
                        Email.NotifyAdministrator("thads@microsoft.com", "Error with processing RegSys data!", ex.Message + ":" + ex.InnerException);
                    }
                }

                if (DateTime.UtcNow.Minute % 5 == 0) // every 5 minutes, look for new invoices to move.
                {
                    try
                    {
                        InvoiceSending_InvokeThread();
                    }
                    catch (Exception ex)
                    {
                        Email.NotifyAdministrator("thads@microsoft.com", "Error with Invoice sending!", ex.Message + ":" + ex.InnerException);
                    }
                }
                if (DateTime.UtcNow.Minute % 3 == 0) // every 5 minutes, look for new activation emails to send.
                {
                    try
                    {
                        ActivationCompleted_InvokeThread();
                    }
                    catch (Exception ex)
                    {
                        Email.NotifyAdministrator("thads@microsoft.com", "Error with Activation email processing!", ex.Message + ":" + ex.InnerException);
                    }
                }
            }
            else
            {
                // PRODUCTION SEQUENCE, being run by a service account!
                if (DateTime.UtcNow.Minute % 2 == 0) // every 2 minutes, look for new orders to process.
                {
                    try
                    {
                        ProcessWizardEntries_InvokeThread();
                    }
                    catch (Exception ex)
                    {
                        Email.NotifyAdministrator("svcasfp@microsoft.com", "Error with processing wizard entries!", ex.Message + ":" + ex.InnerException);
                    }
                }

                if (DateTime.UtcNow.Minute % 5 == 0) // every 5 minutes, look for new invoices to move.
                {
                    try
                    {
                        InvoiceSending_InvokeThread();
                    }
                    catch (Exception ex)
                    {
                        Email.NotifyAdministrator("svcasfp@microsoft.com", "Error with Invoice sending!", ex.Message + ":" + ex.InnerException);
                    }
                }
                if (DateTime.UtcNow.Minute % 3 == 0) // every 5 minutes, look for new activation emails to send.
                {
                    try
                    {
                        ActivationCompleted_InvokeThread();
                    }
                    catch (Exception ex)
                    {
                        Email.NotifyAdministrator("svcasfp@microsoft.com", "Error with Activation email processing!", ex.Message + ":" + ex.InnerException);
                    }
                }
            }
        }
    }
}