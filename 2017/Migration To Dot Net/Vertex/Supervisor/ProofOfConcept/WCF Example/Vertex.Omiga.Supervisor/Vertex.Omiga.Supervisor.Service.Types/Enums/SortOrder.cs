using System;
using System.Runtime.Serialization;

namespace Vertex.Omiga.Supervisor.Service.Types.Enums
{
    /// <summary>
    /// Sort orders
    /// </summary>
    [DataContract]
    public enum SortOrder
    {
        /// <summary>
        /// Default (unspecified) sort order
        /// </summary>
        [EnumMember]
        Unspecified = 0,

        /// <summary>
        /// Ascending sort order
        /// </summary>
        [EnumMember]
        Ascending,

        /// <summary>
        /// Descending sort order
        /// </summary>
        [EnumMember]
        Descending
    }

}
