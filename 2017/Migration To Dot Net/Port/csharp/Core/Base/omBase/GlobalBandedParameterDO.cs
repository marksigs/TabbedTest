/*
--------------------------------------------------------------------------------------------
Workfile:			GlobalBandedParameterDO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Global banded parameters data object
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
MCS		17/08/99	Created
LD		07/11/00	Explicitly close recordsets
------------------------------------------------------------------------------------------
BBG Specific History:

Prog	Date		Description
TK		22/11/04	BBG1821 - Performance related fixes.
------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		27/07/2007	First .Net version. Ported from GlobalParameterDO.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using System.Data.SqlClient;
using System.EnterpriseServices;
using System.Runtime.InteropServices;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.omBase
{
	[ComVisible(true)]
	[ProgId("omBase.GlobalBandedParameterDO")]
	[Guid("CDA181F5-F12F-40D1-A866-6624A107AF6F")]
	[Transaction(TransactionOption.Supported)]
	public class GlobalBandedParameterDO : ServicedComponent
	{
		// ------------------------------------------------------------------------------------------
		// Table format:
		// e.g.
		// VALUATIONFEESET				NOT NULL FLOAT(63)
		// STARTDATE					  NOT NULL DATE
		// TYPEOFVALUATION				NOT NULL NUMBER(5)
		// MAXIMUMVALUE				   NOT NULL NUMBER(10)
		// LOCATION					   NOT NULL NUMBER(5)
		// AMOUNT						 Number(10)
		// ------------------------------------------------------------------------------------------


		// header ----------------------------------------------------------------------------------
		// description:
		// Get the data for a single instance of the persistant data associated with
		// this data object
		// pass:
		// name  Name of the parameter
		// minimumHighBandText	  Qualifier to select appropriate band
		// 
		// return:
		// GetCurrentParameter	 string containing XML data stream representation of
		// data retrieved
		// Raise Errors: if record not found, raise omiga4RecordNotFound
		// ------------------------------------------------------------------------------------------
		public string GetCurrentParameter(string name, string minimumHighBandText) 
		{
			string response = "";

			try 
			{
				string text = new GlobalBandedParameterBooleanText(name, Convert.ToDouble(minimumHighBandText)).Description;
				double? d = new GlobalBandedParameterAmount(name, Convert.ToDouble(minimumHighBandText)).Value;
				text = new GlobalBandedParameterString(name, Convert.ToDouble(minimumHighBandText)).Value;
				text = new GlobalBandedParameterBooleanText(name, Convert.ToDouble(minimumHighBandText)).Value;
				response = new GlobalBandedParameter(name, Convert.ToDouble(minimumHighBandText)).ToString();
			}
			catch (ErrAssistException)
			{
				throw;
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception);
			}
			finally
			{
				 ContextUtility.SetComplete();
			}

			return response;
		}	

		protected override bool CanBePooled() 
		{
			return true;
		}

		#region Not implemented

		/*
		// These methods were defined in the VB6 code as raising a "not implemented" error.
		// Comment out here so it is impossible to call them.
		public string Create(string request) 
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}

		public string Update(string request) 
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}

		public string Delete(string request) 
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}

		public string DeleteAll(string request) 
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}

		public string GetData(string request) 
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}

		public string FindList(string request) 
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}
		 * 
		public string AddDerivedData(string request) 
		{
			throw new Vertex.Fsd.Omiga.VisualBasicPort.NotImplementedException();
		}
		*/

		#endregion

	}

}
