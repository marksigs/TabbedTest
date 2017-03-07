using System;
using System.Runtime.Serialization;

namespace Vertex.Omiga.Supervisor.Service.Types.Requests
{
    /// <summary>
    /// Data Transfer Object (DTO).
    /// Service operation request class
    /// </summary>
    [DataContract]
    public sealed class GetGlobalParametersRequest : BasePagedOperationRequest
    {
        /// <summary>
        /// Initialise a new instance
        /// </summary>
        public GetGlobalParametersRequest() : base() { }

    }

}
