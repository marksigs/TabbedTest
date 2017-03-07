using System;
using System.Runtime.Serialization;

namespace Vertex.Omiga.Supervisor.Service.Types.Contexts
{
    /// <summary>
    /// Data Transfer Object (DTO).
    /// Data paging context class.
    /// </summary>
    [DataContract]
    public sealed class PagingContext
    {
        /// <summary>
        /// Page number.
        /// If used in a request context then this is the page required.
        /// If used in a response context then this is the page provided.
        /// </summary>
        [DataMember]
        public int PageNumber;

        /// <summary>
        /// Page size.
        /// Use zero for service operation default
        /// </summary>
        [DataMember]
        public int PageSize;

    }

}
