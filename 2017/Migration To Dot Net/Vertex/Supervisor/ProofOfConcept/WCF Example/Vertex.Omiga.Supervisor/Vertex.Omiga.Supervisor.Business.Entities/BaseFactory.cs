using System;
using System.Data;

namespace Vertex.Omiga.Supervisor.Business.Entities
{
    /// <summary>
    /// Base class for entity factories.
    /// Contains helper methods etc.
    /// </summary>
    public abstract class BaseFactory
    {
        /// <summary>
        /// Returns the timestamp column value
        /// </summary>
        /// <param name="reader">Data reader instance</param>
        /// <returns>Byte array representing the time stamp value</returns>
        protected static byte[] GetTimeStamp(IDataReader reader)
        {
            byte[] timestamp = (byte[])reader["TimeStamp"];

            return timestamp;
        }

    }

}
