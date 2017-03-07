/*
--------------------------------------------------------------------------------------------
Workfile:			TemplateHandlerDO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Handles print template database.
--------------------------------------------------------------------------------------------
History:
 
Prog	Date		Description
RF		02/11/2000	Created
IK		17/02/2003	BM0200 - add TraceAssist support (remove most of OOSS inheritance)
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		21/06/2007	First .Net version. Ported from TemplateHandlerDO.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Printing
{
	/// <summary>
	/// Handles print template database.
	/// </summary>
	public class TemplateHandlerDO : ITraceAssist
	{
		private const string _typeName = "TemplateHandlerDO";
		private const string _loggerName = "TemplateHandler";
		private TraceAssist _traceAssist = new TraceAssist(_loggerName);

		/// <summary>
		/// Gets a single instance of the persistant data associated with this business object.
		/// </summary>
		/// <param name="xmlRequest">Xml request data stream.</param>
		/// <returns>Xml response node.</returns>
		public XmlNode GetTemplate(XmlElement xmlRequest)
		{
			XmlNode xmlTemplate = null;
			const string functionName = "GetTemplate";

			// ik_20030210
			try
			{
				ContextUtility.SetComplete();

				_traceAssist.TraceMethodEntry(_typeName, functionName);

				XmlDocument xmlClassDefinition = null;

				// TODO call .Net omDPSClassDef
				//objIClassDef = new omDPSClassDef();
				//xmlClassDefinition = objIClassDef.LoadTemplateData;

				_traceAssist.TraceMessage(_typeName, functionName, "[1.1]", "DOAssist.GetData call");
				_traceAssist.TraceMessage(_typeName, functionName, "[1.2]", "DOAssist.GetData return");
				_traceAssist.TraceMethodExit(_typeName, functionName);

				xmlTemplate = DOAssist.GetDataEx(xmlRequest, xmlClassDefinition);
			}
			catch (Exception exception)
			{
				_traceAssist.TraceMethodError(_typeName, functionName, exception);
				throw new ErrAssistException(exception);
			}

			return xmlTemplate;
		}

		/// <summary>
		/// Get the data for all instances of the persistant data associated with
		/// this data object for the values supplied.
		/// </summary>
		/// <param name="xmlRequest">Xml request data stream.</param>
		/// <returns>Xml response node.</returns>
		public XmlNode FindAvailableTemplates(XmlElement xmlRequest)
		{
			XmlElement xmlListElement = null;
			const string functionName = "FindAvailableTemplates";
			
			try
			{
				// ik_20030210
				ContextUtility.SetComplete();

				_traceAssist.TraceMethodEntry(_typeName, functionName);

				// HasValue the correct keys have been passed in
				string securityLevelText = XmlAssist.GetMandatoryNodeText(xmlRequest, ".//SECURITYLEVEL");
				string stageNumberText = XmlAssist.GetMandatoryNodeText(xmlRequest, ".//STAGENUMBER");
				string languageText = XmlAssist.GetMandatoryNodeText(xmlRequest, ".//LANGUAGE");

				DataSet dataSet = new DataSet();
				using (SqlConnection sqlConnection = new SqlConnection(AdoAssist.GetDbConnectString()))
				{
					string sqlString = "SELECT * FROM TEMPLATE WITH(NOLOCK) WHERE SECURITYLEVEL <= @securityLevel AND STAGENUMBER >= @stageNumber AND LANGUAGE = @language";
					using (SqlCommand sqlCommand = new SqlCommand(sqlString, sqlConnection))
					{
						sqlCommand.Parameters.Add("securityLevel", SqlDbType.Int).Value = Convert.ToInt32(securityLevelText);
						sqlCommand.Parameters.Add("stageNumber", SqlDbType.Int).Value = Convert.ToInt32(stageNumberText);
						sqlCommand.Parameters.Add("language", SqlDbType.Int).Value = Convert.ToInt32(languageText);

						using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand))
						{
							sqlDataAdapter.Fill(dataSet);
						}
					}
				}

				if (dataSet.Tables == null || dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows == null || dataSet.Tables[0].Rows.Count == 0)
				{
					throw new RecordNotFoundException();
				}

				XmlDocument xmlClassDefinition = null;
				// TODO call .Net omDPSClassDef
				//objIClassDef = new omDPSClassDef();
				//xmlClassDefinition = objIClassDef.LoadTemplateData;
				XmlDocument xmlOut = new XmlDocument();
				xmlListElement = xmlOut.CreateElement("TEMPLATELIST");
				xmlOut.AppendChild(xmlListElement);
				// Convert recordset to XML
				DOAssist.GetXMLFromWholeRecordset(dataSet.Tables[0], xmlClassDefinition, xmlListElement);

				_traceAssist.TraceMethodExit(_typeName, functionName);
			}
			catch (Exception exception)
			{
				_traceAssist.TraceMethodError(_typeName, functionName, exception);
				throw new ErrAssistException(exception);
			}

			return xmlListElement;
		}


		#region ITraceAssist Members

		/// <summary>
		/// Initialise tracing for this object.
		/// </summary>
		/// <param name="isTracingEnabled">Indicates whether tracing is enabled.</param>
		/// <param name="traceId">The identifier for this trace.</param>
		/// <param name="startTime">The date and time the trace started.</param>
		public void InitialiseTraceInterface(bool isTracingEnabled, string traceId, DateTime startTime)
		{
			if (isTracingEnabled)
			{
				_traceAssist.TraceInitialiseFromParent(isTracingEnabled, traceId, startTime);
			}
		}

		#endregion
	}

}
