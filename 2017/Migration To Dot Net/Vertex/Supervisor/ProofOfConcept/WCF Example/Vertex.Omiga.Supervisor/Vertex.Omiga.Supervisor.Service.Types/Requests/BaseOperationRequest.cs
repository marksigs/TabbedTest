using System;
using System.Runtime.Serialization;

using Vertex.Omiga.Supervisor.Service.Types.Contexts;

namespace Vertex.Omiga.Supervisor.Service.Types.Requests
{
    /// <summary>
    /// Data Transfer Object (DTO).
    /// Base class for all service request classes
    /// </summary>
    [DataContract]
    public abstract class BaseOperationRequest
    {
        /// <summary>
        /// Request context. 
        /// Used to transport contextual information to service
        /// operation requests
        /// </summary>
        [DataMember]
        public RequestContext Context;

        /// <summary>
        /// Initialise a new request context
        /// </summary>
        public BaseOperationRequest()
        {
            this.Context = new RequestContext();
            this.Context.ContextGUID = Guid.NewGuid();
        }

        /// <summary>
        /// Validate the request and throw appropriate exceptions if invalid
        /// </summary>
        /// <exception cref="Exception">Thrown if fails validation</exception>
        public virtual void Validate()
        {
            // general context
            if (this.Context == null) 
                throw new Exception("Context not set");
        }

    }

}
