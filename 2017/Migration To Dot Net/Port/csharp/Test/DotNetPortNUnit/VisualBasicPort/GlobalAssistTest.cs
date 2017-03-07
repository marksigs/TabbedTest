using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class GlobalAssistTest
	{
		public GlobalAssistTest()
		{
		}

		[Test]
		public void TestGetGlobalParamBoolean()
		{
			bool value = false;		
			value = GlobalAssist.GetGlobalParamBoolean("AddressTargetingEnabled");
			Assert.AreEqual(value, true);
			value = GlobalAssist.GetGlobalParamBoolean("ThisParameterDoesNotExist");
			Assert.AreEqual(value, false);
			value = GlobalAssist.GetMandatoryGlobalParamBoolean("AddressTargetingEnabled");
			Assert.AreEqual(value, true);
			try
			{
				GlobalAssist.GetMandatoryGlobalParamBoolean("ThisParameterDoesNotExist");
			}
			catch (ErrAssistException exception)
			{
				Assert.AreEqual(exception.OmigaError, OMIGAERROR.MissingParameter);
			}
		}

		[Test]
		public void TestGetGlobalParamString()
		{
			string value = String.Empty;
			value = GlobalAssist.GetGlobalParamString("AdminSystemName");
			Assert.AreEqual(value, "Optimus");
			value = GlobalAssist.GetGlobalParamString("ThisParameterDoesNotExist");
			Assert.AreEqual(value, String.Empty);
			value = GlobalAssist.GetMandatoryGlobalParamString("AdminSystemName");
			Assert.AreEqual(value, "Optimus");
			try
			{
				GlobalAssist.GetMandatoryGlobalParamString("ThisParameterDoesNotExist");
			}
			catch (ErrAssistException exception)
			{
				Assert.AreEqual(exception.OmigaError, OMIGAERROR.MissingParameter);
			}
		}

		[Test]
		public void TestGetGlobalParamAmount()
		{
			int value = 0;
			value = GlobalAssist.GetGlobalParamAmount("DocumentMangementSystemType");
			Assert.AreEqual(value, 3);
			value = GlobalAssist.GetGlobalParamAmount("ThisParameterDoesNotExist");
			Assert.AreEqual(value, 0);
			value = GlobalAssist.GetMandatoryGlobalParamAmount("DocumentMangementSystemType");
			Assert.AreEqual(value, 3);
			try
			{
				GlobalAssist.GetMandatoryGlobalParamAmount("ThisParameterDoesNotExist");
			}
			catch (ErrAssistException exception)
			{
				Assert.AreEqual(exception.OmigaError, OMIGAERROR.MissingParameter);
			}
		}

		[Test]
		public void TestGetGlobalParamAmountAsDouble()
		{
			double value = 0.0;
			value = GlobalAssist.GetGlobalParamAmountAsDouble("AdditionalIncomeRatio");
			Assert.AreEqual(value, 0.25);
			value = GlobalAssist.GetGlobalParamAmountAsDouble("ThisParameterDoesNotExist");
			Assert.AreEqual(value, 0.0);
			value = GlobalAssist.GetMandatoryGlobalParamAmountAsDouble("AdditionalIncomeRatio");
			Assert.AreEqual(value, 0.25);
			try
			{
				GlobalAssist.GetMandatoryGlobalParamAmountAsDouble("ThisParameterDoesNotExist");
			}
			catch (ErrAssistException exception)
			{
				Assert.AreEqual(exception.OmigaError, OMIGAERROR.MissingParameter);
			}
		}

		[Test]
		public void TestGetGlobalParamMaximumAmount()
		{
			int value = 0;
			value = GlobalAssist.GetGlobalParamMaximumAmount("LPMinValProportion");
			Assert.AreEqual(value, 70);
			value = GlobalAssist.GetGlobalParamMaximumAmount("ThisParameterDoesNotExist");
			Assert.AreEqual(value, 0);
			value = GlobalAssist.GetMandatoryGlobalParamMaximumAmount("LPMinValProportion");
			Assert.AreEqual(value, 70);
			try
			{
				GlobalAssist.GetMandatoryGlobalParamMaximumAmount("ThisParameterDoesNotExist");
			}
			catch (ErrAssistException exception)
			{
				Assert.AreEqual(exception.OmigaError, OMIGAERROR.MissingParameter);
			}
		}

		[Test]
		public void TestGetGlobalParamMaximumAmountAsDouble()
		{
			double value = 0.0;
			value = GlobalAssist.GetGlobalParamMaximumAmountAsDouble("LPMinValProportion");
			Assert.AreEqual(value, 70.0);
			value = GlobalAssist.GetGlobalParamMaximumAmountAsDouble("ThisParameterDoesNotExist");
			Assert.AreEqual(value, 0.0);
			value = GlobalAssist.GetMandatoryGlobalParamMaximumAmountAsDouble("LPMinValProportion");
			Assert.AreEqual(value, 70.0);
			try
			{
				GlobalAssist.GetMandatoryGlobalParamMaximumAmountAsDouble("ThisParameterDoesNotExist");
			}
			catch (ErrAssistException exception)
			{
				Assert.AreEqual(exception.OmigaError, OMIGAERROR.MissingParameter);
			}
		}

		[Test]
		public void TestGetGlobalParamPercentage()
		{
			double value = 70.0;
			value = GlobalAssist.GetGlobalParamPercentage("LPMinValProportion");
			Assert.AreEqual(value, 70.0);
			value = GlobalAssist.GetGlobalParamPercentage("ThisParameterDoesNotExist");
			Assert.AreEqual(value, 0.0);
			value = GlobalAssist.GetMandatoryGlobalParamPercentage("LPMinValProportion");
			Assert.AreEqual(value, 70.0);
			try
			{
				GlobalAssist.GetMandatoryGlobalParamPercentage("ThisParameterDoesNotExist");
			}
			catch (ErrAssistException exception)
			{
				Assert.AreEqual(exception.OmigaError, OMIGAERROR.MissingParameter);
			}
		}

		[Test]
		public void TestGetAllGlobalBandedParamValuesAsXml()
		{
			XmlNode xmlResponseNode = null;
			GlobalAssist.GetAllGlobalBandedParamValuesAsXml("TotalScore", "AMOUNT", xmlResponseNode);
			Assert.IsTrue(xmlResponseNode != null);
			Assert.IsTrue(xmlResponseNode.OuterXml.Length > 0);
		}
	}
}
