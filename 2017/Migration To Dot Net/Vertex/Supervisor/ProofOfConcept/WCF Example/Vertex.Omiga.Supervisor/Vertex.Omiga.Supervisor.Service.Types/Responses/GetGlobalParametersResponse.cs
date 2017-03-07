using System;
using System.Runtime.Serialization;

using Vertex.Omiga.Supervisor.Service.Types.Contexts;
using Vertex.Omiga.Supervisor.Service.Types.Entities;

namespace Vertex.Omiga.Supervisor.Service.Types.Responses
{
    /// <summary>
    /// Data Transfer Object (DTO).
    /// Service operation response class
    /// </summary>
    [DataContract]
    public sealed class GetGlobalParametersResponse : BaseOperationResponse
    {
        /// <summary>
        /// Array of <see cref="GlobalParameter"/> entities
        /// </summary>
        [DataMember]
        public GlobalParameter[] GlobalParameters;

        /// <summary>
        /// Initialise response
        /// </summary>
        /// <param name="requestContext">
        /// Corresponding request context that produced this service response
        /// </param>
        public GetGlobalParametersResponse(RequestContext requestContext) : base(requestContext) { }

    }

}
