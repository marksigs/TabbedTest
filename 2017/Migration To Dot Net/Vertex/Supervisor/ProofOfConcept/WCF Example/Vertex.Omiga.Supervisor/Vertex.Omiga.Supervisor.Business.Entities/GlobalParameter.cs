using System;

namespace Vertex.Omiga.Supervisor.Business.Entities
{
    /// <summary>
    /// Global Parameter business entity.
    /// Entity identity
    /// Entity summary
    /// Entity detail
    /// </summary>
    public class GlobalParameterIdentity
    {
        /// <summary>
        /// Timestamp (for concurrency)
        /// </summary>
        public byte[] TimeStamp;

        /// <summary>
        /// GlobalParameter name
        /// </summary>
        public string Name;

        /// <summary>
        /// Start date
        /// </summary>
        public DateTime StartDate;

    }

    /// <summary>
    /// Global Parameter business entity.
    /// Entity identity
    /// Entity summary
    /// Entity detail
    /// </summary>
    public class GlobalParameterSummary : GlobalParameterIdentity
    {
        /// <summary>
        /// Description
        /// </summary>
        public string Description;

    }

    /// <summary>
    /// Global Parameter business entity.
    /// Entity identity
    /// Entity summary
    /// Entity detail
    /// </summary>
    public sealed class GlobalParameter : GlobalParameterSummary
    {
        /// <summary>
        /// Amount value
        /// </summary>
        public double? ValueAmount;

        /// <summary>
        /// Maximum amount value
        /// </summary>
        public double? ValueMaximumAmount;

        /// <summary>
        /// Percentage value
        /// </summary>
        public double? ValuePercentage;

        /// <summary>
        /// Boolean value
        /// </summary>
        public bool? ValueBoolean;

        /// <summary>
        /// String value
        /// </summary>
        public string ValueString;

    }

}
