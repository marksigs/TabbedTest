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
    public sealed class DeleteGlobalParameterResponse : BaseOperationResponse
    {
        /// <summary>
        /// Initialise response
        /// </summary>
        /// <param name="requestContext">
        /// Corresponding request context that produced this service response
        /// </param>
        public DeleteGlobalParameterResponse(RequestContext requestContext) : base(requestContext) { }

    }

}
