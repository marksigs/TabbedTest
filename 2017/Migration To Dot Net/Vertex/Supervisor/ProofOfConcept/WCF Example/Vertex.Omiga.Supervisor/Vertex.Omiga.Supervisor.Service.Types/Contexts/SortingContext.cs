using System;
using System.Runtime.Serialization;

using Vertex.Omiga.Supervisor.Service.Types.Enums;

namespace Vertex.Omiga.Supervisor.Service.Types.Contexts
{
    /// <summary>
    /// Data Transfer Object (DTO).
    /// Data sorting context class.
    /// </summary>
    [DataContract]
    public sealed class SortingContext<T>
    {
        private T _sortBy;

        /// <summary>
        /// Order to sort data by
        /// </summary>
        [DataMember]
        public T SortBy
        {
            get { return _sortBy; }
            set
            {
                // workaround as cannot constrain generics to Enum
                if (value.GetType().BaseType != typeof(Enum))
                    throw new ArgumentException("T must be of type System.Enum");
                else
                    _sortBy = value;
            }
        }

        /// <summary>
        /// Sort order
        /// </summary>
        [DataMember]
        public SortOrder SortOrder;

    }

}
