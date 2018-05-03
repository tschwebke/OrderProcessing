using System.ComponentModel;
using System.Configuration.Install;
using System.Reflection;
using System.ServiceProcess;

namespace Microsoft.Operations.Infrastructure
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();

            serviceInstaller1.ServiceName = string.Format("RegSys TFS Connector", Assembly.GetCallingAssembly().GetVersion());
            serviceInstaller1.Description = "ASfP RegSys to TFS Connector";
            serviceProcessInstaller1.Account = ServiceAccount.LocalSystem;
            serviceInstaller1.StartType = ServiceStartMode.Automatic;
        }

        private void serviceInstaller1_AfterInstall(object sender, InstallEventArgs e)
        {
        }
    }
}