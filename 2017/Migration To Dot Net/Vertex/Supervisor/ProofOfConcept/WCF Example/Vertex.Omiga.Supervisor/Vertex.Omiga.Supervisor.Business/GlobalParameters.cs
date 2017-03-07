using System;

using Vertex.Omiga.Supervisor.Data;
using Vertex.Omiga.Supervisor.Business.Entities;

namespace Vertex.Omiga.Supervisor.Business
{
    /// <summary>
    /// Omiga Supervisor business logic relating to 
    /// <see cref="GlobalParameter"/> entities
    /// </summary>
    public sealed class GlobalParameters
    {
        /// <summary>
        /// Retrieve all <see cref="GlobalParameter"/> entities
        /// </summary>
        /// <returns>Array of entities</returns>
        public static GlobalParameter[] GetGlobalParameters()
        {
            //todo: validate entities on retrieval

            // retrieve from database
            return GlobalParametersDataAccess.GetGlobalParameters();
        }

        /// <summary>
        /// Create a <see cref="GlobalParameter"/> entity
        /// </summary>
        /// <param name="globalParameter">Entity to be created</param>
        /// <returns>New entity</returns>
        public static GlobalParameter CreateGlobalParameter(GlobalParameter globalParameter)
        {
            //todo: validate entity info prior to persistence

            // persist to database
            return GlobalParametersDataAccess.CreateGlobalParameter(globalParameter);
        }

        /// <summary>
        /// Update a <see cref="GlobalParameter"/> entity
        /// </summary>
        /// <param name="globalParameter">Entity to be updated</param>
        /// <returns>Updated entity</returns>
        public static GlobalParameter UpdateGlobalParameter(GlobalParameter globalParameter)
        {
            //todo: validate entity info prior to persistence

            // persist to database
            return GlobalParametersDataAccess.UpdateGlobalParameter(globalParameter);
        }

        /// <summary>
        /// Delete a <see cref="GlobalParameterIdentity"/> entity identity
        /// </summary>
        /// <param name="globalParameter">Entity to be deleted</param>
        public static void DeleteGlobalParameter(GlobalParameterIdentity globalParameter)
        {
            // delete from database
            GlobalParametersDataAccess.DeleteGlobalParameter(globalParameter);
        }

    }

}
