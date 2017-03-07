using System;
using System.Runtime.Serialization;

namespace Vertex.Omiga.Supervisor.Service.Types.Contexts
{
    /// <summary>
    /// Data Transfer Object (DTO).
    /// Response context class.
    /// Used to transport contextual information to service
    /// operation responses
    /// </summary>
    [DataContract]
    public sealed class ResponseContext : BaseContext
    {
        /// <summary>
        /// Corresponding request context that produced this 
        /// service response context
        /// </summary>
        [DataMember]
        public RequestContext RequestContext;

        /// <summary>
        /// DateTime response sent from service
        /// </summary>
        [DataMember]
        public DateTime SentAt;

        /// <summary>
        /// Initialise response context
        /// </summary>
        /// <param name="request">
        /// Corresponding request context that produced 
        /// this service response context
        /// </param>
        internal ResponseContext(RequestContext request) 
        {
            this.RequestContext = request;
        }

    }

}
