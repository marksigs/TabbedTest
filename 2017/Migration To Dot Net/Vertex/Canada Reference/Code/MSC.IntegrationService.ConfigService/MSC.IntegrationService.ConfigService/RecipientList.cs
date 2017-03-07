using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Runtime.Serialization;
using System.Xml.Serialization.Configuration;

namespace MSC.IntegrationService.ConfigService
{
    [Serializable]
    [ConfigurationCollection(typeof(Recipient), AddItemName = "recipient",
      CollectionType = ConfigurationElementCollectionType.BasicMap)]
    public class RecipientList : ConfigurationElementCollection, ISerializable
    {
        #region Constructors
        public RecipientList() { }
        protected RecipientList(SerializationInfo info, StreamingContext context)
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
            get { return "recipient"; }
        }
        #endregion

        #region Indexers
        public Recipient this[string name]
        {
            get
            {
                return (Recipient)base.BaseGet(name);
            }
        }
        public Recipient this[int index]
        {
            get { return (Recipient)base.BaseGet(index); }
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
        public Recipient GetRecipient(string p_name)
        {
            return (Recipient)base.BaseGet(p_name);
        }
        public void Add(Recipient recipient)
        {
            base.BaseAdd(recipient);
        }

        public void Remove(string name)
        {
            base.BaseRemove(name);
        }

        public void Remove(Recipient recipient)
        {
            base.BaseRemove(GetElementKey(recipient));
        }

        public void Clear()
        {
            base.BaseClear();
        }

        public void RemoveAt(int index)
        {
            base.BaseRemoveAt(index);
        }

        public string GetKey(int index)
        {
            return (string)base.BaseGetKey(index);
        }
        #endregion

        #region Overrides

        protected override ConfigurationElement CreateNewElement()
        {
            return new Recipient();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return (element as Recipient).Name;
        }
        #endregion
    }
}
