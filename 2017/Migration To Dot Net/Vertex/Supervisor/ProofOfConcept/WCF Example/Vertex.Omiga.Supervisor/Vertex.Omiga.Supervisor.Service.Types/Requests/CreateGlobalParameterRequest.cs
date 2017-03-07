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
    public sealed class CreateGlobalParameterRequest : BaseOperationRequest
    {
        /// <summary>
        /// New entity to be created
        /// </summary>
        [DataMember]
        public GlobalParameter EntityToCreate;

        /// <summary>
        /// Initialise a new instance
        /// </summary>
        public CreateGlobalParameterRequest() : base() { }

    }

}
