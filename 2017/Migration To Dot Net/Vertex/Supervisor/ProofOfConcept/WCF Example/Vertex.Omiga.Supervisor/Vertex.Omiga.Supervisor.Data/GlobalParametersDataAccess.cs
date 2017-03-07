using System;
using System.Configuration;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;

using Vertex.Omiga.Supervisor.Business.Entities;
using Vertex.Omiga.Supervisor.Business.Entities.Utilities;

namespace Vertex.Omiga.Supervisor.Data
{
    /// <summary>
    /// Data access layer (DAL) for database interaction with 
    /// <see cref="GlobalParameter"/> entities
    /// </summary>
    public sealed class GlobalParametersDataAccess : BaseDataAccess
    {
        /// <summary>
        /// Data access method to return an array of all
        /// <see cref="GlobalParameter"/> entities
        /// </summary>
        /// <returns>Array of entities</returns>
        /// <exception cref="DataException">
        /// Thrown if a database related error occurs
        /// </exception>
        public static GlobalParameter[] GetGlobalParameters()
        {
            try
            {
                List<GlobalParameter> entities = new List<GlobalParameter>();

                using (DbConnection connection = GetConnection())
                {
                    using (DbCommand command = DataAccessHelper.CreateStoredProcedureCommand(connection, "dbo.USP_GetGlobalParameters"))
                    {
                        connection.Open();

                        using (DbDataReader reader = command.ExecuteReader())
                        {
                            if (reader != null && reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    GlobalParameter entity = GlobalParameterFactory.CreateFromReader(new NullableDataReader(reader));
                                    entities.Add(entity);
                                }
                            }
                        }
                    }
                }

                return entities.ToArray();
            }

            catch (Exception ex)
            {
                throw new DataException(string.Format("A database error occurred ({0})", ex.Message), ex);
            }
        }

        /// <summary>
        /// Data access method to persist a new 
        /// <see cref="GlobalParameter"/> entity
        /// </summary>
        /// <param name="newGlobalParameter">New entity to persist</param>
        /// <returns>New entity read-back from database</returns>
        /// <exception cref="DataException">
        /// Thrown if a database related error occurs
        /// </exception>
        public static GlobalParameter CreateGlobalParameter(GlobalParameter newGlobalParameter)
        {
            try
            {
                GlobalParameter persistedGlobalParameter = null;

                using (DbConnection connection = GetConnection())
                {
                    using (DbCommand command = DataAccessHelper.CreateStoredProcedureCommand(connection, "dbo.USP_CreateGlobalParameter"))
                    {
                        DataAccessHelper.AddParameter(command, "Name", newGlobalParameter.Name);
                        DataAccessHelper.AddParameter(command, "GlobalParameterStartDate", newGlobalParameter.StartDate);
                        DataAccessHelper.AddParameter(command, "Description", newGlobalParameter.Description);
                        DataAccessHelper.AddParameter(command, "Amount", newGlobalParameter.ValueAmount);
                        DataAccessHelper.AddParameter(command, "MaximumAmount", newGlobalParameter.ValueMaximumAmount);
                        DataAccessHelper.AddParameter(command, "Percentage", newGlobalParameter.ValuePercentage);
                        DataAccessHelper.AddParameter(command, "Boolean", newGlobalParameter.ValueBoolean);
                        DataAccessHelper.AddParameter(command, "String", newGlobalParameter.ValueString);

                        connection.Open();

                        using (DbDataReader reader = command.ExecuteReader())
                        {
                            if (reader != null && reader.HasRows)
                            {
                                reader.Read();
                                persistedGlobalParameter = GlobalParameterFactory.CreateFromReader(new NullableDataReader(reader));
                            }
                            else
                            {
                                // SP did not return expected row containing new entity
                                throw new Exception("Expected data was not returned from procedure " + command.CommandText);
                            }
                        }
                    }
                }

                return persistedGlobalParameter;
            }

            catch (Exception ex)
            {
                throw new DataException(string.Format("A database error occurred ({0})", ex.Message), ex);
            }
        }

        /// <summary>
        /// Data access method to persist an updated 
        /// <see cref="GlobalParameter"/> entity
        /// </summary>
        /// <param name="newGlobalParameter">Entity to persist</param>
        /// <returns>Updated entity read-back from database</returns>
        /// <exception cref="DataException">
        /// Thrown if a database related error occurs
        /// </exception>
        public static GlobalParameter UpdateGlobalParameter(GlobalParameter globalParameter)
        {
            try
            {
                GlobalParameter updatedGlobalParameter = null;

                using (DbConnection connection = GetConnection())
                {
                    using (DbCommand command = DataAccessHelper.CreateStoredProcedureCommand(connection, "dbo.USP_UpdateGlobalParameter"))
                    {
                        DataAccessHelper.AddParameter(command, "Name", globalParameter.Name);
                        DataAccessHelper.AddParameter(command, "GlobalParameterStartDate", globalParameter.StartDate);
                        DataAccessHelper.AddParameter(command, "Description", globalParameter.Description);
                        DataAccessHelper.AddParameter(command, "Amount", globalParameter.ValueAmount);
                        DataAccessHelper.AddParameter(command, "MaximumAmount", globalParameter.ValueMaximumAmount);
                        DataAccessHelper.AddParameter(command, "Percentage", globalParameter.ValuePercentage);
                        DataAccessHelper.AddParameter(command, "Boolean", globalParameter.ValueBoolean);
                        DataAccessHelper.AddParameter(command, "String", globalParameter.ValueString);
                        DataAccessHelper.AddParameter(command, "TimeStamp", globalParameter.TimeStamp);

                        connection.Open();

                        using (DbDataReader reader = command.ExecuteReader())
                        {
                            if (reader != null && reader.HasRows)
                            {
                                reader.Read();
                                updatedGlobalParameter = GlobalParameterFactory.CreateFromReader(new NullableDataReader(reader));
                            }
                            else
                            {
                                // SP did not return expected row containing new entity
                                throw new Exception("Expected data was not returned from procedure " + command.CommandText);
                            }
                        }
                    }
                }

                return updatedGlobalParameter;
            }

            catch (Exception ex)
            {
                throw new DataException(string.Format("A database error occurred ({0})", ex.Message), ex);
            }
        }

        /// <summary>
        /// Data access method to delete a 
        /// <see cref="GlobalParameter"/> entity
        /// </summary>
        /// <param name="newGlobalParameter">Entity to be deleted</param>
        /// <exception cref="DataException">
        /// Thrown if a database related error occurs
        /// </exception>
        public static void DeleteGlobalParameter(GlobalParameterIdentity globalParameter)
        {
            try
            {
                using (DbConnection connection = GetConnection())
                {
                    using (DbCommand command = DataAccessHelper.CreateStoredProcedureCommand(connection, "dbo.USP_DeleteGlobalParameter"))
                    {
                        DataAccessHelper.AddParameter(command, "Name", globalParameter.Name);
                        DataAccessHelper.AddParameter(command, "GlobalParameterStartDate", globalParameter.StartDate);
                        DataAccessHelper.AddParameter(command, "TimeStamp", globalParameter.TimeStamp);

                        connection.Open();

                        command.ExecuteNonQuery();
                    }
                }
            }

            catch (Exception ex)
            {
                throw new DataException(string.Format("A database error occurred ({0})", ex.Message), ex);
            }
        }

    }

}
