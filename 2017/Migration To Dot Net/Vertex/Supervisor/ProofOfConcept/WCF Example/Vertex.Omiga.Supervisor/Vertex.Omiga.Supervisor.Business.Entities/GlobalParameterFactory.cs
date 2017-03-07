using System;
using System.Data;
using System.Data.Common;

using Vertex.Omiga.Supervisor.Business.Entities.Utilities;

namespace Vertex.Omiga.Supervisor.Business.Entities
{
    /// <summary>
    /// Entity factory
    /// </summary>
    public class GlobalParameterFactory : BaseFactory
    {
        /// <summary>
        /// Create a <see cref="GlobalParameter"/> entity from
        /// a data reader
        /// </summary>
        /// <param name="reader">
        /// <see cref="NullableDataReader"/> or <see cref="NullableDataRowReader"/> instance
        /// </param>
        /// <returns>New entity instance</returns>
        public static GlobalParameter CreateFromReader(INullableReader reader)
        {
            GlobalParameter entity = new GlobalParameter();

            entity.Name = reader.GetString("Name");
            entity.Description = reader.GetNullableString("Description");
            entity.StartDate = reader.GetDateTime("GlobalParameterStartDate");
            entity.ValueString = reader.GetNullableString("String");
            entity.ValueAmount = reader.GetNullableDouble("Amount");
            entity.ValueMaximumAmount = reader.GetNullableDouble("MaximumAmount");
            entity.ValuePercentage = reader.GetNullableDouble("Percentage");
            entity.ValueBoolean = Convert.ToBoolean(reader.GetNullableByte("Boolean"));
            byte? value = reader.GetNullableByte("Boolean");
            entity.ValueBoolean = (value.HasValue) ? (bool?)Convert.ToBoolean(value) : null;

            entity.TimeStamp = GetTimeStamp((IDataReader)reader);

            return entity;
        }

    }

}
