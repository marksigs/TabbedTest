using System;
using System.Runtime.Serialization;

using Vertex.Omiga.Supervisor.Service.Types.Contexts;

namespace Vertex.Omiga.Supervisor.Service.Types.Responses
{
    /// <summary>
    /// Data Transfer Object (DTO).
    /// Base class for paged service response classes
    /// </summary>
    [DataContract]
    public abstract class BasePagedOperationResponse : BaseOperationResponse
    {
        /// <summary>
        /// Number of data items in total 
        /// (i.e. not just for this page)
        /// </summary>
        [DataMember]
        public int TotalDataSize;

        /// <summary>
        /// Initialise response
        /// </summary>
        /// <param name="requestContext">Corresponding request context 
        /// that produced this service response</param>
        /// <param name="totalDataSize">Total size of the data</param>
        public BasePagedOperationResponse(RequestContext requestContext, int totalDataSize)
            : base(requestContext) 
        {
            this.TotalDataSize = totalDataSize;
        }

        /// <summary>
        /// Validate the response and throw appropriate exceptions if invalid
        /// </summary>
        /// <exception cref="Exception">Thrown if fails validation</exception>
        public override void Validate()
        {
            base.Validate();

            // add additional validation here
        }

    }

}
