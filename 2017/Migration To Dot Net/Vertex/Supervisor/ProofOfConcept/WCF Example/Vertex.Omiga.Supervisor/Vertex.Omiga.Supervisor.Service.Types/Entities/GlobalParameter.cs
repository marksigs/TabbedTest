using System;
using System.Runtime.Serialization;

namespace Vertex.Omiga.Supervisor.Service.Types.Entities
{
    /// <summary>
    /// Global Parameter service entity (serializable).
    /// Entity identity
    /// </summary>
    [DataContract]
    public class GlobalParameterIdentity
    {
        /// <summary>
        /// Timestamp (for concurrency)
        /// </summary>
        [DataMember]
        public byte[] TimeStamp;

        /// <summary>
        /// GlobalParameter name
        /// </summary>
        [DataMember]
        public string Name;

        /// <summary>
        /// Start date
        /// </summary>
        [DataMember]
        public DateTime StartDate;

    }

    /// <summary>
    /// Global Parameter service entity (serializable).
    /// Entity summary
    /// </summary>
    [DataContract]
    public class GlobalParameterSummary : GlobalParameterIdentity
    {
        /// <summary>
        /// Description
        /// </summary>
        [DataMember]
        public string Description;

    }

    /// <summary>
    /// Global Parameter service entity (serializable).
    /// Entity detail
    /// </summary>
    [DataContract]
    public sealed class GlobalParameter : GlobalParameterSummary
    {
        /// <summary>
        /// Amount value
        /// </summary>
        /// <remarks>May be null</remarks>
        [DataMember]
        public double? ValueAmount;

        /// <summary>
        /// Maximum amount value.
        /// If specified, then <see cref="ValuePercentage"/> must 
        /// also be specified
        /// </summary>
        /// <remarks>May be null</remarks>
        [DataMember]
        public double? ValueMaximumAmount;

        /// <summary>
        /// Percentage value
        /// </summary>
        /// <remarks>May be null</remarks>
        [DataMember]
        public double? ValuePercentage;

        /// <summary>
        /// Boolean value
        /// </summary>
        /// <remarks>May be null</remarks>
        [DataMember]
        public bool? ValueBoolean;

        /// <summary>
        /// String value
        /// </summary>
        /// <remarks>May be null</remarks>
        [DataMember]
        public string ValueString;

    }

}
