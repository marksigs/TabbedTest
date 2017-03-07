using System;
using System.Runtime.Serialization;

namespace Vertex.Omiga.Supervisor.Service.Types.Contexts
{
    /// <summary>
    /// Data Transfer Object (DTO).
    /// Request context class.
    /// Used to transport contextual information to service
    /// operation requests
    /// </summary>
    [DataContract]
    public sealed class RequestContext : BaseContext
    {
        /// <summary>
        /// Authentication token
        /// </summary>
        [DataMember]
        public Guid AuthenticationToken;

        /// <summary>
        /// DateTime request received at service
        /// </summary>
        [DataMember]
        public DateTime ReceivedAt;

    }

}
