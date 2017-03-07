using System;

using NUnit.Framework;

using Microsoft.Practices.EnterpriseLibrary.Validation;
using Microsoft.Practices.EnterpriseLibrary.Validation.Validators;

using Vertex.Omiga.Supervisor.Service.Types.Entities;

namespace Vertex.Omiga.Supervisor.Service.Tests
{
    /// <summary>
    /// Test fixture to test validation logic for the
    /// GlobalParameter service entity
    /// </summary>
    [TestFixture]
    public class GlobalParametersValidationTests : BaseTestFixtureForValidation
    {
        #region Positive tests
        /// <summary>
        /// Simple validation test
        /// </summary>
        [Test]
        public void GlobalParameterValidationSuccessTest()
        {
            //todo: write validation test(s)

            GlobalParameter globalParameter = new GlobalParameter();

            globalParameter.Name = "Name";
            globalParameter.StartDate = DateTime.Now;
            globalParameter.Description = "Description";
            globalParameter.ValueAmount = 0;
            globalParameter.ValueBoolean = true;
            globalParameter.ValuePercentage = 50;
            globalParameter.ValueMaximumAmount = 75;
            globalParameter.ValueString = string.Empty;

            // validate entity
            Validator validator = ValidationFactory.CreateValidatorFromConfiguration(typeof(GlobalParameter), "GlobalParameter");
            ValidationResults results = validator.Validate(globalParameter);

            /*
            // report results
            Console.WriteLine("Entity valid? {0}", results.IsValid);

            foreach (ValidationResult result in results)
                Console.WriteLine("...error: {0}, {1} ", result.Key, result.Message);
            */
        }

        #endregion

        #region Positive tests

        #endregion

    }

}
