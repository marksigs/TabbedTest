using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Runtime.Serialization;
using System.Xml.Serialization.Configuration;

namespace MSC.IntegrationService.ConfigService
{
    [Serializable]
    [ConfigurationCollection(typeof(ServiceType), AddItemName = "serviceType",
      CollectionType = ConfigurationElementCollectionType.BasicMap)]
    public class ServiceTypeList : ConfigurationElementCollection, ISerializable
    {
        #region Constructors
        public ServiceTypeList() { }
        protected ServiceTypeList(SerializationInfo info, StreamingContext context)
        {
        }
        #endregion

        #region ISerializable Implementation
        void ISerializable.GetObjectData(SerializationInfo info, StreamingContext context)
        {
        }
        #endregion

        #region Properties
        public override ConfigurationElementCollectionType CollectionType
        {
            get
            {
                return ConfigurationElementCollectionType.BasicMap;
            }
        }
        protected override string ElementName
        {
            get { return "serviceType"; }
        }
        #endregion

        #region Indexers
        public ServiceType this[string name]
        {
            get
            {
                return (ServiceType)base.BaseGet(name);
            }
        }
        public ServiceType this[int index]
        {
            get { return (ServiceType)base.BaseGet(index); }
            set
            {
                if (base.BaseGet(index) != null)
                {
                    base.BaseRemoveAt(index);
                    
                }
                base.BaseAdd(index, value);
            }
        }
        #endregion
        
        #region Methods
        public ServiceType GetServicType(string p_name)
        {
            return (ServiceType)base.BaseGet(p_name);
        }
        public void Add(ServiceType p_serviceType)
        {
            base.BaseAdd(p_serviceType);
        }

        public void Remove(string p_name)
        {
            base.BaseRemove(p_name);
        }

        public void Remove(ServiceType p_serviceType)
        {
            base.BaseRemove(GetElementKey(p_serviceType));
        }

        public void Clear()
        {
            base.BaseClear();
        }

        public void RemoveAt(int p_index)
        {
            base.BaseRemoveAt(p_index);
        }

        public string GetKey(int p_index)
        {
            return (string)base.BaseGetKey(p_index);
        }
        #endregion

        #region Overrides
        protected override ConfigurationElement CreateNewElement()
        {
            return new ServiceType();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return (element as ServiceType).Name;
        }
        #endregion
    }
}
