using System;
using System.Runtime.Serialization;

using Vertex.Omiga.Supervisor.Service.Types.Entities;

namespace Vertex.Omiga.Supervisor.Service.Types.Requests
{
    /// <summary>
    /// Data Transfer Object (DTO).
    /// Service operation request class
    /// </summary>
    [DataContract]
    public sealed class DeleteGlobalParameterRequest : BaseOperationRequest
    {
        /// <summary>
        /// Entity to be deleted
        /// </summary>
        [DataMember]
        public GlobalParameterIdentity EntityToDelete;

        /// <summary>
        /// Initialise a new instance
        /// </summary>
        public DeleteGlobalParameterRequest() : base() { }

    }

}
