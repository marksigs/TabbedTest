using System;
using System.ServiceProcess;

namespace Vertex.Omiga.Supervisor.Service.Host
{
    /// <summary>
    /// Entry point
    /// </summary>
    public static class Program
    {
        /// <summary>
        /// Entry point
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun = new ServiceBase[] { 
                new VertexOmigaSupervisorServiceHost() 
            };

            ServiceBase.Run(ServicesToRun);
        }

    }

}
