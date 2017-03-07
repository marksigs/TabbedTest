using System;
using System.ServiceModel;

using Vertex.Omiga.Supervisor.Service.Types.Requests;
using Vertex.Omiga.Supervisor.Service.Types.Responses;
using Vertex.Omiga.Supervisor.Service.Types.Faults;

namespace Vertex.Omiga.Supervisor.Service.Interface
{
    /// <summary>
    /// Omiga supervisor service interface
    /// (Service Contract)
    /// </summary>
    [ServiceContract]
    public interface IOmigaSupervisorService
    {
        #region Global Parameters operations
        /// <summary>
        /// Retrieve all <see cref="GlobalParameter"/> entities
        /// </summary>
        /// <param name="request">Request</param>
        /// <returns>Response</returns>
        [OperationContract]
        [FaultContract(typeof(OmigaSupervisorFault))]
        GetGlobalParametersResponse GetGlobalParameters(GetGlobalParametersRequest request);

        /// <summary>
        /// Create a new <see cref="GlobalParameter"/> entity
        /// </summary>
        /// <param name="request">Request</param>
        /// <returns>Response</returns>
        [OperationContract]
        [FaultContract(typeof(OmigaSupervisorFault))]
        CreateGlobalParameterResponse CreateGlobalParameter(CreateGlobalParameterRequest request);

        /// <summary>
        /// Update an existing <see cref="GlobalParameter"/> entity
        /// </summary>
        /// <param name="request">Request</param>
        /// <returns>Response</returns>
        [OperationContract]
        [FaultContract(typeof(OmigaSupervisorFault))]
        UpdateGlobalParameterResponse UpdateGlobalParameter(UpdateGlobalParameterRequest request);

        /// <summary>
        /// Delete an existing <see cref="GlobalParameter"/> entity
        /// </summary>
        /// <param name="request">Request</param>
        /// <returns>Response</returns>
        [OperationContract]
        [FaultContract(typeof(OmigaSupervisorFault))]
        DeleteGlobalParameterResponse DeleteGlobalParameter(DeleteGlobalParameterRequest request);

        #endregion

    }

}
