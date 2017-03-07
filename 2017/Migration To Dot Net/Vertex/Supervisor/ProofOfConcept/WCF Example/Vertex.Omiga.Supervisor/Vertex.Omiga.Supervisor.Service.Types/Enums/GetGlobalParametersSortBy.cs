using System;
using System.Runtime.Serialization;

using Vertex.Omiga.Supervisor.Service.Types.Requests;

namespace Vertex.Omiga.Supervisor.Service.Types.Enums
{
    /// <summary>
    /// Defines allowable data sorting values for use with the
    /// <see cref="GetGlobalParametersRequest"/>
    /// service request
    /// </summary>
    [DataContract]
    public enum GetGlobalParametersSortBy
    {
        /// <summary>
        /// Default sort (i.e. unspecified)
        /// </summary>
        [EnumMember]
        Default = 0,

        /// <summary>
        /// Just an example
        /// </summary>
        [EnumMember]
        ExampleSortBy //todo: example code
    }

}
