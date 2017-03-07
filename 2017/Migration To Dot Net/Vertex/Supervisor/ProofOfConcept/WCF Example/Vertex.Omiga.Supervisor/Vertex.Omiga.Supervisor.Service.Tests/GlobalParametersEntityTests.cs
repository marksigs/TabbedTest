using System;

using NUnit.Framework;

using Vertex.Omiga.Supervisor.Service.Interface;
using Vertex.Omiga.Supervisor.Service.Types.Entities;

namespace Vertex.Omiga.Supervisor.Service.Tests
{
    /// <summary>
    /// Unit tests for Omiga Supervisor services.
    /// Test fixture to test Global Parameter services
    /// </summary>
    [TestFixture]
    public class GlobalParametersServiceTests : BaseTestFixtureForEntity
    {
        #region Setups & Teardowns
        /// <summary>
        /// Fixture setup
        /// </summary>
        [TestFixtureSetUp]
        public override void FixtureSetup()
        {
            base.FixtureSetup();
        }

        /// <summary>
        /// Fixture teardown
        /// </summary>
        [TestFixtureTearDown]
        public override void FixtureTearDown()
        {
            base.FixtureTearDown();
        }

        #endregion

        #region Positive tests
        /// <summary>
        /// Positive test.
        /// Ensure some global parameters are returned 
        /// from the service
        /// </summary>
        [Test]
        public void GetGlobalParametersTest()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Positive test.
        /// Round-trip serializtion test for service entity type
        /// </summary>
        [Test]
        public override void RoundTripSerializationTest()
        {
            string entityDescription = "EntityDescription";
            string entityName = "EntityName";
            DateTime entityStartDate = DateTime.Now;
            double entityValueAmount = 1;
            double entityValuePercentage = 50.0;
            double entityValueMaximumAmount = 100.0;
            bool entityValueBoolean = true;
            string entityValueString = "EntityValueString";

            GlobalParameter entity = new GlobalParameter();
            entity.Description = entityDescription;
            entity.Name = entityName;
            entity.StartDate = entityStartDate;
            entity.ValueAmount = entityValueAmount;
            entity.ValueBoolean = entityValueBoolean;
            entity.ValueMaximumAmount = entityValueMaximumAmount;
            entity.ValuePercentage = entityValuePercentage;
            entity.ValueString = entityValueString;

            string xml = SerializeToString<GlobalParameter>(entity);

            //Console.WriteLine(xml);

            GlobalParameter readback = DeserializeFromString<GlobalParameter>(xml);

            Assert.AreEqual(entity.Description, readback.Description);
            Assert.AreEqual(entity.Name, readback.Name);
            Assert.AreEqual(entity.StartDate, readback.StartDate);
            Assert.AreEqual(entity.ValueAmount, readback.ValueAmount);
            Assert.AreEqual(entity.ValueBoolean, readback.ValueBoolean);
            Assert.AreEqual(entity.ValueMaximumAmount, readback.ValueMaximumAmount);
            Assert.AreEqual(entity.ValuePercentage, readback.ValuePercentage);
            Assert.AreEqual(entity.ValueString, readback.ValueString);
        }

        #endregion

        #region Negative tests
        //todo: Negative tests for GlobalParametersServiceTests

        #endregion

    }

}
