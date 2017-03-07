using System;
using System.ServiceModel;
using System.ServiceModel.Description;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.ServiceProcess;
using System.Text;

using Vertex.Omiga.Supervisor.Service;
using Vertex.Omiga.Supervisor.Service.Interface;

namespace Vertex.Omiga.Supervisor.Service.Host
{
    /// <summary>
    /// Windows service host class
    /// </summary>
    public partial class VertexOmigaSupervisorServiceHost : ServiceBase
    {
        private static ServiceHost _host = null;

        /// <summary>
        /// Initializes a new instance of the <see cref="VertexOmigaSupervisorServiceHost"/> class.
        /// </summary>
        public VertexOmigaSupervisorServiceHost()
        {
            InitializeComponent();
        }

        /// <summary>
        /// When implemented in a derived class, executes when a Start command is sent to the service by the Service Control Manager (SCM) or when the operating system starts (for a service that starts automatically). Specifies actions to take when the service starts.
        /// </summary>
        /// <param name="args">Data passed by the start command.</param>
        protected override void OnStart(string[] args)
        {
            _host = new ServiceHost(typeof(OmigaSupervisorService));
            _host.Open();
        }

        /// <summary>
        /// When implemented in a derived class, executes when a Stop command is sent to the service by the Service Control Manager (SCM). Specifies actions to take when a service stops running.
        /// </summary>
        protected override void OnStop()
        {
            if (_host.State != CommunicationState.Closed)
                _host.Close();
        }

    }

}
