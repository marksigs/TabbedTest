using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;
using NUnit.Framework;
using Vertex.Fsd.Omiga.omBase;
using Vertex.Fsd.Omiga.VisualBasicPort;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class GlobalParameterTest
	{
		/*
		[Test]
		public void TestGlobalParameterDOCreate()
		{
			GlobalParameterDO globalParameterDO = new GlobalParameterDO();
			globalParameterDO.Create(
				"<GLOBALPARAMETER>" +
					"<NAME>Adrian</NAME>" +
					"<GLOBALPARAMETERSTARTDATE>25/07/2007 00:00:00</GLOBALPARAMETERSTARTDATE>" +
					"<DESCRIPTION>My global parameter</DESCRIPTION>" +
					"<STRING>This must be the place</STRING>" +
				"</GLOBALPARAMETER>");
		}

		[Test]
		public void TestGlobalParameterDOUpdate()
		{
			GlobalParameterDO globalParameterDO = new GlobalParameterDO();
			globalParameterDO.Update(
				"<GLOBALPARAMETER>" +
					"<NAME>Adrian</NAME>" +
					"<GLOBALPARAMETERSTARTDATE>25/07/2007 00:00:00</GLOBALPARAMETERSTARTDATE>" +
					"<DESCRIPTION>My global parameter</DESCRIPTION>" +
					"<STRING>Stop making sense</STRING>" +
				"</GLOBALPARAMETER>");
		}

		[Test]
		public void TestGlobalParameterDODelete()
		{
			GlobalParameterDO globalParameterDO = new GlobalParameterDO();
			globalParameterDO.Delete(
				"<GLOBALPARAMETER>" +
					"<NAME>Adrian</NAME>" +
					"<GLOBALPARAMETERSTARTDATE>25/07/2007 00:00:00</GLOBALPARAMETERSTARTDATE>" +
				"</GLOBALPARAMETER>");
		}
		 */

		[Test]
		public void TestGlobalParameterDOGetCurrentParameter()
		{
			GlobalParameterDO globalParameterDO = new GlobalParameterDO();
			string response = globalParameterDO.GetCurrentParameter("AccountNumber");		// STRING
			response = globalParameterDO.GetCurrentParameter("AdditionalIncomeRatio");		// AMOUNT 0.25
			response = globalParameterDO.GetCurrentParameter("AddressNos");					// AMOUNT 3
			response = globalParameterDO.GetCurrentParameter("AddressTargetingEnabled");	// BOOLEAN 1
			response = globalParameterDO.GetCurrentParameter("AFAdditionalIncomeCheck");	// BOOLEAN 0
		}

		[Test]
		public void TestGlobalParameterDOGetCurrentParameterByType()
		{
			GlobalParameterDO globalParameterDO = new GlobalParameterDO();
			string response = globalParameterDO.GetCurrentParameterByType("AccountNumber", "STRING");
		}

		[Test]
		public void TestGlobalParameterDOGetCurrentParameterListEx()
		{
			GlobalParameterDO globalParameterDO = new GlobalParameterDO();
			string response = globalParameterDO.GetCurrentParameterListEx(
				"<REQUEST>" + 
					"<GLOBALPARAMETER>" +
						"<NAME>AccountNumber</NAME>" +
					"</GLOBALPARAMETER>" +
					"<GLOBALPARAMETER>" +
						"<NAME>AdminSystemName</NAME>" +
					"</GLOBALPARAMETER>" +
				"</REQUEST>");
		}

		[Test]
		public void TestGlobalParameterBOGetCurrentParameter()
		{
			try
			{
				GlobalParameterBO globalParameterBO = new GlobalParameterBO();
				string response = globalParameterBO.GetCurrentParameter("AccountNumber1");
			}
			catch (Exception)
			{
			}
		}

		[Test]
		public void TestGlobalParameterBOIsTaskManager()
		{
			GlobalParameterBO globalParameterBO = new GlobalParameterBO();
			bool response = globalParameterBO.IsTaskManager();
			Assert.IsTrue(response);
		}

		[Test]
		public void TestGlobalParameterBOIsMultipleLender()
		{
			GlobalParameterBO globalParameterBO = new GlobalParameterBO();
			string response = globalParameterBO.IsMultipleLender();
		}

		[Test]
		public void TestGlobalParameterBOFindCurrentParameterList()
		{
			GlobalParameterBO globalParameterBO = new GlobalParameterBO();
			string response = globalParameterBO.FindCurrentParameterList(
				"<REQUEST>" + 
					"<GLOBALPARAMETER>" +
						"<NAME>AccountNumber</NAME>" +
					"</GLOBALPARAMETER>" +
					"<GLOBALPARAMETER>" +
						"<NAME>AdminSystemName</NAME>" +
					"</GLOBALPARAMETER>" +
				"</REQUEST>");
		}

		[Test]
		public void TestGlobalParameterBOGetCurrentParameterListEx()
		{
			GlobalParameterBO globalParameterBO = new GlobalParameterBO();
			string response = globalParameterBO.GetCurrentParameterListEx(
				"<REQUEST>" +
					"<GLOBALPARAMETER>" +
						"<NAME>AccountNumber</NAME>" +
					"</GLOBALPARAMETER>" +
					"<GLOBALPARAMETER>" +
						"<NAME>AdminSystemName</NAME>" +
					"</GLOBALPARAMETER>" +
				"</REQUEST>");
		}

		[Test]
		public void TestGlobalBandedParameterDOGetCurrentParameter()
		{
			GlobalBandedParameterDO globalBandedParameterDO = new GlobalBandedParameterDO();
			string response = globalBandedParameterDO.GetCurrentParameter("DependantsIncomeFactor", "2.0");
		}

		[Test]
		public void TestGlobalBandedParameterBOGetCurrentParameter()
		{
			GlobalBandedParameterBO globalBandedParameterBO = new GlobalBandedParameterBO();
			string response = globalBandedParameterBO.GetCurrentParameter("DependantsIncomeFactor", "2.0");
		}

	}
}
