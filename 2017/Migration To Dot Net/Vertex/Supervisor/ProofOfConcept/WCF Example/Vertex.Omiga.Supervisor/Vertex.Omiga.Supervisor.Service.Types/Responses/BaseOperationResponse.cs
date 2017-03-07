using System;
using System.Runtime.Serialization;

using Vertex.Omiga.Supervisor.Service.Types.Contexts;

namespace Vertex.Omiga.Supervisor.Service.Types.Responses
{
    /// <summary>
    /// Data Transfer Object (DTO).
    /// Base class for all service response classes
    /// </summary>
    [DataContract]
    public abstract class BaseOperationResponse
    {
        /// <summary>
        /// Response context. 
        /// Used to transport contextual information to service
        /// operation reponses
        /// </summary>
        [DataMember]
        public ResponseContext Context;

        /// <summary>
        /// Initialise response
        /// </summary>
        /// <param name="requestContext">
        /// Corresponding request context that produced 
        /// this service response
        /// </param>
        public BaseOperationResponse(RequestContext requestContext)
        {
            Context = new ResponseContext(requestContext);
            Context.ContextGUID = Guid.NewGuid();
        }

        /// <summary>
        /// Validate the response and throw appropriate exceptions if invalid
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
