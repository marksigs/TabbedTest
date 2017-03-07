using System;

namespace Vertex.Omiga.Supervisor.Service.Tests
{
    /// <summary>
    /// Base unit test class for Omiga Supervisor 
    /// entity tests.
    /// </summary>
    public abstract class BaseTestFixtureForEntity : BaseTestFixture
    {
        /// <summary>
        /// Positive test.
        /// Round-trip serializtion test for service entity type
        /// </summary>
        public abstract void RoundTripSerializationTest();

    }

}
