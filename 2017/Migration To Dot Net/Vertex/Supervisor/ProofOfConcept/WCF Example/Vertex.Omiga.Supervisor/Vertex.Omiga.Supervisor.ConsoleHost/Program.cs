using System;
using System.ServiceModel;

using Vertex.Omiga.Supervisor.Service;
using Vertex.Omiga.Supervisor.Service.Interface;

namespace Vertex.Omiga.Supervisor.ConsoleHost
{
    /// <summary>
    /// Example console host for WCF service
    /// </summary>
    public class Program
    {
        /// <summary>
        /// Entry point
        /// </summary>
        /// <param name="args">Arguments</param>
        static void Main(string[] args)
        {
            using (ServiceHost host = new ServiceHost(typeof(OmigaSupervisorService)))
            {
                host.Open();

                Console.WriteLine("Service running, press ENTER to end service");
                Console.ReadLine();
            }
        }

    }

}
