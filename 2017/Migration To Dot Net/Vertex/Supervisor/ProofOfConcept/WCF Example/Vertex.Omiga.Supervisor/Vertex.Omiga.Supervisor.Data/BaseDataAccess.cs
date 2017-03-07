using System;
using System.Data.Common;
using System.Configuration;

using Vertex.Omiga.Supervisor.Data.Properties;

namespace Vertex.Omiga.Supervisor.Data
{
    /// <summary>
    /// Base class for data access
    /// </summary>
    public abstract class BaseDataAccess
    {
        #region Helpers
        /// <summary>
        /// Returns a database connection that is named as the DefaultConnection
        /// within the config file
        /// </summary>
        /// <returns>Database connection</returns>
        protected static DbConnection GetConnection()
        { 
            return GetConnection(Settings.Default.DefaultConnection);
        }

        /// <summary>
        /// Returns a named database connection. The connection with the specified
        /// name must exist within the config file
        /// </summary>
        /// <param name="connectionName">Name of connection as in config file</param>
        /// <returns>Database connection</returns>
        protected static DbConnection GetConnection(string connectionName)
        {
            ConnectionStringSettings connectionString = ConfigurationManager.ConnectionStrings[connectionName];
            DbProviderFactory factory = DbProviderFactories.GetFactory(connectionString.ProviderName);

            DbConnection connection = factory.CreateConnection();

            connection.ConnectionString = connectionString.ConnectionString;

            return connection;
        }

        #endregion

    }

}
