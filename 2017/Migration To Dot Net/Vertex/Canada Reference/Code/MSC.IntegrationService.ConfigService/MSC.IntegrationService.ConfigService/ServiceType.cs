using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Runtime.Serialization;
using System.Xml.Serialization.Configuration;

namespace MSC.IntegrationService.ConfigService
{
    [Serializable]
    public class ServiceType : ConfigurationElement, ISerializable
    {
        #region Constrcutors
        public ServiceType() 
        {
            sName = new ConfigurationProperty("name", typeof(string), null, ConfigurationPropertyOptions.None);
            sDirection = new ConfigurationProperty("direction", typeof(string), null, ConfigurationPropertyOptions.None);
            bAggregateResponse = new ConfigurationProperty("aggregateResponse", typeof(bool), null, ConfigurationPropertyOptions.None);
            collRecipientList = new ConfigurationProperty("recipientList", typeof(RecipientList), null, ConfigurationPropertyOptions.None);
        }
        protected ServiceType(SerializationInfo info, StreamingContext context)
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
        private static ConfigurationProperty sDirection;
        private static ConfigurationProperty bAggregateResponse;
        private static ConfigurationProperty collRecipientList;
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
        [ConfigurationProperty("direction")]
        public string Direction
        {
            get
            {
                return (string)base[sDirection];
            }
        }
        [ConfigurationProperty("aggregateResponse")]
        public bool AggregateResponse
        {
            get
            {
                return (bool)base[bAggregateResponse];
            }
        }
        [ConfigurationProperty("recipientList")]
        public RecipientList RecipientList
        {
            get
            {
                return (RecipientList)base[collRecipientList];
            }
        }
        
        #endregion
    }
}
