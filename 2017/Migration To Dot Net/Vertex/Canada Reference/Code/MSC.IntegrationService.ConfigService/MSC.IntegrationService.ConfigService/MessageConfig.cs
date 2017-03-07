using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Runtime.Serialization;
using System.Xml.Serialization.Configuration;


namespace MSC.IntegrationService.ConfigService
{
    [Serializable]
    public class MessageConfig : ConfigurationSection, ISerializable
    {
        #region constructor
        /// <summary>
        /// Predefines valida properties and prepares property colelction
        /// </summary>
        public MessageConfig()
        {
            
            sDescription = new ConfigurationProperty("description", typeof(string), null, ConfigurationPropertyOptions.None);
            sType = new ConfigurationProperty("type", typeof(string), null, ConfigurationPropertyOptions.None);
            bDebug = new ConfigurationProperty("debug", typeof(bool), null, ConfigurationPropertyOptions.None);
            collServiceTypeList = new ConfigurationProperty("serviceTypeList", typeof(ServiceTypeList), null, ConfigurationPropertyOptions.None);
           /* ConfigurationPropertyCollection messageConfigProperties = new ConfigurationPropertyCollection();
            messageConfigProperties.Add(sDescription);
            messageConfigProperties.Add(sSource);
            messageConfigProperties.Add(sDirection);
            messageConfigProperties.Add(sMessageType);
            messageConfigProperties.Add(sAggregateResponse);
            */
        }
        
        //Deserialization constructor.
        protected MessageConfig(SerializationInfo info, StreamingContext context) 
        {
            /*
            sDescription = (ConfigurationProperty)info.GetValue("sDescription", typeof(ConfigurationProperty));
            sSource = (ConfigurationProperty)info.GetValue("sSource", typeof(ConfigurationProperty));
            sDirection = (ConfigurationProperty)info.GetValue("sDirection", typeof(ConfigurationProperty));
            sMessageType = (ConfigurationProperty)info.GetValue("sMessageType", typeof(ConfigurationProperty));
            sAggregateResponse = (ConfigurationProperty)info.GetValue("sAggregateResponse", typeof(ConfigurationProperty));
            */
        }
        //Serialization function.
        void ISerializable.GetObjectData(SerializationInfo info, StreamingContext context) 
        {
            /*
            info.AddValue("sDescription", sDescription);
            info.AddValue("sSource", sSource);
            info.AddValue("sDirection", sDirection);
            info.AddValue("sMessageType", sMessageType);
            info.AddValue("sAggregateResponse", sAggregateResponse);
              */
        }
         
        #endregion
        #region static fields
        private static ConfigurationProperty sDescription;
        private static ConfigurationProperty sType;
        private static ConfigurationProperty bDebug;
        private static ConfigurationProperty collServiceTypeList;
        #endregion
        
        #region Properties
        
        [ConfigurationProperty ("description")]
        public string Description
        {
            get 
            {
                return (string) base[sDescription]; 
            }
        }
        [ConfigurationProperty("type")]
        public string Source
        {
            get 
            {
                return (string)base[sType]; 
            }
        }
        [ConfigurationProperty("debug")]
        public bool Direction
        {
            get 
            {
                return (bool)base[bDebug]; 
            }
        }
        [ConfigurationProperty("serviceTypeList")]
        public ServiceTypeList ServiceTypes
        {
            get
            {
                return (ServiceTypeList)base[collServiceTypeList];
            }
        }
        #endregion
    }
    
    
}
