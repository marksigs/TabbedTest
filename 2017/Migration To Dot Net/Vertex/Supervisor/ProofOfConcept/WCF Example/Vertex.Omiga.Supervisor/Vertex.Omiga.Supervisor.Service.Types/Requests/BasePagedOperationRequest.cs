using System;
using System.Runtime.Serialization;

using Vertex.Omiga.Supervisor.Service.Types.Contexts;
using Vertex.Omiga.Supervisor.Service.Types.Enums;

namespace Vertex.Omiga.Supervisor.Service.Types.Requests
{
    /// <summary>
    /// Data Transfer Object (DTO).
    /// Base class for paged service request classes
    /// </summary>
    [DataContract]
    public abstract class BasePagedOperationRequest : BaseOperationRequest
    {
        /// <summary>
        /// Data paging context
        /// </summary>
        [DataMember]
        public PagingContext PagingContext;

        /// <summary>
        /// Data sorting context
        /// </summary>
        [DataMember]
        public SortingContext<GetGlobalParametersSortBy> SortingContext;

        /// <summary>
        /// Initialise a new paged request context
        /// </summary>
        public BasePagedOperationRequest() : base()
        {
            PagingContext = new PagingContext();
            PagingContext.PageNumber = 1;
            PagingContext.PageSize = 50; //todo: default page size is hard-coded (should vary by service)

            SortingContext = new SortingContext<GetGlobalParametersSortBy>();
            SortingContext.SortOrder = SortOrder.Unspecified;
            SortingContext.SortBy = 0;
        }

        /// <summary>
        /// Validate the request and throw appropriate exceptions if invalid
        /// </summary>
        /// <exception cref="Exception">Thrown if fails validation</exception>
        public override void Validate()
        {
            base.Validate();

            // paging content
            if (this.PagingContext == null)
                throw new Exception("Paging context not set");
            if (this.PagingContext.PageNumber < 1)
                throw new Exception("Invalid page number");
            if (this.PagingContext.PageSize < 0)
                throw new Exception("Invalid page size");

            // sorting context
            if (this.SortingContext == null)
                throw new Exception("Sorting context not set");
        }

    }

}
