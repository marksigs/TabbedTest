using System;
using System.Runtime.Serialization;

namespace Vertex.Omiga.Supervisor.Service.Types.Faults
{
    /// <summary>
    /// Omiga supervisor service fault class
    /// </summary>
    [DataContract]
    public class OmigaSupervisorFault
    {
        private Guid _guid;
        private string _message;

        /// <summary>
        /// Initialise a new instance
        /// </summary>
        /// <param name="faultMessage">Fault message</param>
        /// <param name="faultGuid">Fault GUID</param>
        public OmigaSupervisorFault(string faultMessage, Guid faultGuid)
        {
            _message = faultMessage;
            _guid = faultGuid;
        }

        /// <summary>
        /// Fault message
        /// </summary>
        [DataMember]
        public string FaultMessage
        {
            get { return _message; }
            set { _message = value; }
        }

        /// <summary>
        /// Fault GUID
        /// </summary>
        [DataMember]
        public Guid FaultGuid
        {
            get { return _guid; }
            set { _guid = value; }
        }

    }

}
