using System;
using System.Runtime.Serialization;

namespace Vertex.Omiga.Supervisor.Service.Types.Contexts
{
    /// <summary>
    /// Data Transfer Object (DTO).
    /// Base context class
    /// </summary>
    [DataContract]
    public abstract class BaseContext
    {
        /// <summary>
        /// Calling application identifier
        /// </summary>
        [DataMember]
        public string RequestSource;

        /// <summary>
        /// Context GUID
        /// </summary>
        [DataMember]
        public Guid ContextGUID;

        /// <summary>
        /// Initialise new instance
        /// </summary>
        public BaseContext()
        {
            this.ContextGUID = Guid.NewGuid();
        }

    }

}
