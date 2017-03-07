using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;

namespace Vertex.Omiga.Supervisor.Service.Host
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();

            ServiceInstaller si = new ServiceInstaller();
            ServiceProcessInstaller spi = new ServiceProcessInstaller();

            si.ServiceName = "VertexOmigaConfigurationServiceHost";
            si.DisplayName = "Vertex Omiga Configuration Service Host";
            si.Description = "Vertex Omiga Configuration Service Host";
            this.Installers.Add(si);

            spi.Account = ServiceAccount.LocalSystem;
            spi.Password = null;
            spi.Username = null;
 
            this.Installers.Add(spi);
        }
    }
}