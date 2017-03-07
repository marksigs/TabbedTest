using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Runtime.Serialization;
using System.Xml.Serialization.Configuration;

namespace MSC.IntegrationService.ConfigService
{
    [Serializable]
    public class Recipient : ConfigurationElement, ISerializable
    {
        #region Constrcutors
        public Recipient() 
        {
            sName = new ConfigurationProperty("name", typeof(string), null, ConfigurationPropertyOptions.None);
            sUrl = new ConfigurationProperty("url", typeof(string), null, ConfigurationPropertyOptions.None);
            sUserId = new ConfigurationProperty("userId", typeof(string), null, ConfigurationPropertyOptions.None);
            sPassword = new ConfigurationProperty("password", typeof(string), null, ConfigurationPropertyOptions.None);
            sProtocol = new ConfigurationProperty("protocol", typeof(string), null, ConfigurationPropertyOptions.None);
            sAssembly = new ConfigurationProperty("assembly", typeof(string), null, ConfigurationPropertyOptions.None);
            sOperation = new ConfigurationProperty("operation", typeof(string), null, ConfigurationPropertyOptions.None);
            sOperationType = new ConfigurationProperty("operationType", typeof(string), null, ConfigurationPropertyOptions.None);
            sRecipientDomain = new ConfigurationProperty("recipientDomain", typeof(string), null, ConfigurationPropertyOptions.None);
            sRequestMapType = new ConfigurationProperty("requestMapType", typeof(string), null, ConfigurationPropertyOptions.None);
            sResponseMapType = new ConfigurationProperty("responseMapType", typeof(string), null, ConfigurationPropertyOptions.None);
        }
        protected Recipient(SerializationInfo info, StreamingContext context)
        {
        }

        #endregion
        #region ISerializable Implementation
        void ISerializable.GetObjectData(SerializationInfo info, StreamingContext context)
        {
        }
        #endregion

        #region Static Fields
        private static ConfigurationProperty sName;
        private static ConfigurationProperty sUrl;
        private static ConfigurationProperty sUserId;
        private static ConfigurationProperty sPassword;
        private static ConfigurationProperty sProtocol;
        private static ConfigurationProperty sAssembly;
        private static ConfigurationProperty sOperation;
        private static ConfigurationProperty sOperationType;
        private static ConfigurationProperty sRecipientDomain;
        private static ConfigurationProperty sRequestMapType;
        private static ConfigurationProperty sResponseMapType;
        #endregion

        #region Properties
        [ConfigurationProperty("name")]
        public string Name
        {
            get
            {
                return (string)base[sName];
            }
        }
        [ConfigurationProperty("url")]
        public string Url
        {
            get
            {
                return (string)base[sUrl];
            }
        }
        [ConfigurationProperty("userId")]
        public string UserId
        {
            get
            {
                return (string)base[sUserId];
            }
        }
        [ConfigurationProperty("password")]
        public string Password
        {
            get
            {
                return (string)base[sPassword];
            }
        }
        [ConfigurationProperty("protocol")]
        public string Protocol
        {
            get
            {
                return (string)base[sProtocol];
            }
        }
        [ConfigurationProperty("assembly")]
        public string Assembly
        {
            get
            {
                return (string)base[sAssembly];
            }
        }
        [ConfigurationProperty("operation")]
        public string Operation
        {
            get
            {
                return (string)base[sOperation];
            }
        }
        [ConfigurationProperty("operationType")]
        public string OperationType
        {
            get
            {
                return (string)base[sOperationType];
            }
        }
        [ConfigurationProperty("recipientDomain")]
        public string RecipientDomain
        {
            get
            {
                return (string)base[sRecipientDomain];
            }
        }
        [ConfigurationProperty("requestMapType")]
        public string RequestMapType
        {
            get
            {
                return (string)base[sRequestMapType];
            }
        }
        [ConfigurationProperty("responseMapType")]
        public string ResponseMapType
        {
            get
            {
                return (string)base[sResponseMapType];
            }
        }
        
        #endregion
    }
}
