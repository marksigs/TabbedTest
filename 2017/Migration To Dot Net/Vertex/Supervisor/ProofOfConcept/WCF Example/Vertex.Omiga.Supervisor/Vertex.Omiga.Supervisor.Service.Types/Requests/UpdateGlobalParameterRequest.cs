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
    public sealed class UpdateGlobalParameterRequest : BaseOperationRequest
    {
        /// <summary>
        /// Entity to be updated
        /// </summary>
        [DataMember]
        public GlobalParameter EntityToUpdate;

        /// <summary>
        /// Initialise a new instance
        /// </summary>
        public UpdateGlobalParameterRequest() : base() { }

    }

}
