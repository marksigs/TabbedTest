using System;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using System.Xml;
using System.IO;

using Vertex.Omiga.Supervisor.Service.Interface;
using Vertex.Omiga.Supervisor.Service.Types.Entities;

namespace Vertex.Omiga.Supervisor.Service.Tests
{
    /// <summary>
    /// Base unit test class for Omiga Supervisor.
    /// </summary>
    public abstract class BaseTestFixture
    {
        /// <summary>
        /// Service proxy
        /// </summary>
        protected IOmigaSupervisorService _proxy;

        #region Setups & Teardowns
        /// <summary>
        /// Fixture setup
        /// </summary>
        public virtual void FixtureSetup()
        {
            ChannelFactory<IOmigaSupervisorService> factory = new ChannelFactory<IOmigaSupervisorService>("NetTcpBinding_IOmigaSupervisorService");
            _proxy = factory.CreateChannel();
        }

        /// <summary>
        /// Fixture teardown
        /// </summary>
        public virtual void FixtureTearDown()
        {
            //
        }

        #endregion

        #region XML Serialization / Deserialization helpers
        /// <summary>
        /// Returns a string containing the specified object
        /// serialized to XML using the WCF serialization engine
        /// </summary>
        /// <typeparam name="T">Type of object</typeparam>
        /// <param name="instance">Object to serialize</param>
        /// <returns>XML string</returns>
        /// <remarks>
        /// Object must be serializable.
        /// Contains no exception handling
        /// </remarks>
        protected static string SerializeToString<T>(T instance)
        {
            StringBuilder builder = new StringBuilder();

            using (XmlWriter xmlWriter = XmlTextWriter.Create(builder))
            using (XmlDictionaryWriter writer = XmlDictionaryWriter.CreateDictionaryWriter(xmlWriter))
            {
                DataContractSerializer serializer = new DataContractSerializer(typeof(T));
                serializer.WriteObject(writer, instance);
            }

            return builder.ToString();
        }

        /// <summary>
        /// Returns an instance of the specified type from
        /// the XML string provided using the WCF serialization engine
        /// </summary>
        /// <typeparam name="T">Type of object</typeparam>
        /// <param name="xml">Serialized object XML</param>
        /// <returns>Instance of deserialized object</returns>
        /// <remarks>Contains no exception handling</remarks>
        protected static T DeserializeFromString<T>(string xml)
        {
            T instance;

            using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
            using (XmlDictionaryReader reader = XmlDictionaryReader.CreateDictionaryReader(xmlReader))
            {
                DataContractSerializer serializer = new DataContractSerializer(typeof(T));
                instance = (T)serializer.ReadObject(reader);
            }

            return instance;
        }

        #endregion

    }

}
