using System;
using System.Data.Common;
using System.Data;

namespace Vertex.Omiga.Supervisor.Data
{
    /// <summary>
    /// Helper class for data access
    /// </summary>
    internal static class DataAccessHelper
    {
        #region Command helpers
        /// <summary>
        /// Creates a stored procedure command object
        /// </summary>
        /// <param name="connection">Connection</param>
        /// <param name="procedureName">Name of the procedure</param>
        /// <returns>New command object</returns>
        internal static DbCommand CreateStoredProcedureCommand(DbConnection connection, string procedureName)
        {
            DbCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = procedureName;

            return command;
        }

        #endregion

        #region Parameter helpers
        /// <summary>
        /// Adds a new string parameter to the command object
        /// </summary>
        /// <param name="command">Command</param>
        /// <param name="parameterName">Name of the parameter</param>
        /// <param name="value">Parameter value</param>
        internal static void AddParameter(DbCommand command, string parameterName, string value)
        {
            AddParameter(command, parameterName, DbType.String, value);
        }

        /// <summary>
        /// Adds a new datetime parameter to the command object
        /// </summary>
        /// <param name="command">Command</param>
        /// <param name="parameterName">Name of the parameter</param>
        /// <param name="value">Parameter value</param>
        internal static void AddParameter(DbCommand command, string parameterName, DateTime value)
        {
            AddParameter(command, parameterName, DbType.DateTime, value);
        }

        /// <summary>
        /// Adds a new double? parameter to the command object
        /// </summary>
        /// <param name="command">Command</param>
        /// <param name="parameterName">Name of the parameter</param>
        /// <param name="value">Parameter value</param>
        internal static void AddParameter(DbCommand command, string parameterName, double? value)
        {
            AddParameter(command, parameterName, DbType.Double, value);
        }

        /// <summary>
        /// Adds a new bool? parameter to the command object
        /// </summary>
        /// <param name="command">Command</param>
        /// <param name="parameterName">Name of the parameter</param>
        /// <param name="value">Parameter value</param>
        internal static void AddParameter(DbCommand command, string parameterName, bool? value)
        {
            AddParameter(command, parameterName, DbType.Boolean, value);
        }

        /// <summary>
        /// Adds a new byte[] (e.g. a SQL timestamp) parameter to the command object
        /// </summary>
        /// <param name="command">Command</param>
        /// <param name="parameterName">Name of the parameter</param>
        /// <param name="value">Parameter value</param>
        internal static void AddParameter(DbCommand command, string parameterName, byte[] value)
        {
            AddParameter(command, parameterName, DbType.Binary, value);
        }

        #endregion

        #region Helpers
        /// <summary>
        /// Adds a new parameter to the command object
        /// </summary>
        /// <param name="command">Command</param>
        /// <param name="parameterName">Name of the parameter</param>
        /// <param name="dbType">Parameter type</param>
        /// <param name="value">Parameter value</param>
        private static void AddParameter(DbCommand command, string parameterName, DbType dbType, object value)
        {
            DbParameter parameter = command.CreateParameter();
            parameter.ParameterName = parameterName;
            parameter.DbType = dbType;
            parameter.Value = value;
            command.Parameters.Add(parameter);
        }

        #endregion

    }

}
