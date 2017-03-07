/*
--------------------------------------------------------------------------------------------
Workfile:			TemplateHandlerBO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Handles print template database.
--------------------------------------------------------------------------------------------
History:
 
Prog	Date		Description
RF		02/11/2000	Created
IK		17/02/2003	BM0200 - add TraceAssist support
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		21/06/2007	First .Net version. Ported from TemplateHandlerBO.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Printing
{
	/// <summary>
	/// Handles print template database.
	/// </summary>
	public class TemplateHandlerBO : ITraceAssist
	{
		private const string _typeName = "TemplateHandlerBO";
		private const string _loggerName = "TemplateHandler";
		private const string _rootNodeName = "TEMPLATE";
		private TraceAssist _traceAssist = new TraceAssist(_loggerName);
		private TemplateHandlerDO _templateHandlerDO = new TemplateHandlerDO();

		/// <summary>
		/// Gets a single instance of the persistant data associated with this business object.
		/// </summary>
		/// <param name="xmlRequest">Xml request data stream.</param>
		/// <returns>Xml response node.</returns>
		public XmlElement GetTemplate(XmlElement xmlRequest)
		{
			const string functionName = "GetTemplate";
			XmlElement xmlResponseElement = null;

			try
			{
				_traceAssist.TraceRequest(xmlRequest.OuterXml);
				_traceAssist.TraceMethodEntry(_typeName, functionName);

				XmlDocument xmlOut = new XmlDocument();
				xmlResponseElement = xmlOut.CreateElement("RESPONSE");
				xmlOut.AppendChild(xmlResponseElement);
				xmlResponseElement.SetAttribute("TYPE", "SUCCESS");

				XmlNode xmlRequestNode = null;
				if (xmlRequest.Name == _rootNodeName)
				{
					xmlRequestNode = xmlRequest;
				}
				else
				{
					xmlRequestNode = xmlRequest.GetElementsByTagName(_rootNodeName)[0];
				}

				if (xmlRequestNode == null)
				{
					throw new MissingPrimaryTagException(_rootNodeName + " tag not found");
				}

				_traceAssist.TraceInitialiseOffspring(_templateHandlerDO);
				xmlResponseElement.AppendChild(_templateHandlerDO.GetTemplate((XmlElement)xmlRequestNode));

				_traceAssist.TraceResponse(xmlOut.OuterXml);
				_traceAssist.TraceMethodExit(_typeName, functionName);
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception);
			}

			return xmlResponseElement;
		}

		/// <summary>
		/// Gets a single instance of the persistant data associated with this business object.
		/// </summary>
		/// <param name="request">Xml request data stream containing data to be persisted.</param>
		/// <returns>Xml response.</returns>
		public string GetTemplate(string request)
		{
			string response = "";
			const string functionName = "GetTemplate";

			try
			{
				// ik_20030210
				ContextUtility.SetComplete();

				_traceAssist.TraceMethodEntry(_typeName, functionName);

				// Create default response block
				XmlDocument xmlOut = new XmlDocument();
				XmlElement xmlResponseElement = xmlOut.CreateElement("RESPONSE");
				xmlOut.AppendChild(xmlResponseElement);
				xmlResponseElement.SetAttribute("TYPE", "SUCCESS");
				XmlDocument xmlIn = XmlAssist.Load(request);

				// Delegate to FreeThreadedDomDocument40 based method and attach returned data to our response
				XmlNode xmlTempResponseNode = GetTemplate(xmlIn.DocumentElement);
				ErrAssistException.CheckXmlResponseNode(xmlTempResponseNode, xmlResponseElement, true);
				XmlAssist.AttachResponseData((XmlNode)xmlResponseElement, (XmlElement)xmlTempResponseNode);
				response = xmlResponseElement.OuterXml;

				_traceAssist.TraceMethodExit(_typeName, functionName);
			}
			catch (Exception exception)
			{
				response = new ErrAssistException(exception).CreateErrorResponse();
			}

			return response;
		}

		/// <summary>
		/// Finds all the templates with: a security level &lt;= to the supplied security level, 
		/// a stage &gt;= to the supplied stage, a language matching the supplied language.
		/// </summary>
		/// <param name="xmlRequest">Xml request data stream containing security level, appication stage.</param>
		/// <returns>
		/// Xml containing list of template ids, names and descriptions.
		/// </returns>
		public XmlNode FindAvailableTemplates(XmlElement xmlRequest) 
		{
			const string functionName = "FindAvailableTemplates";
			XmlElement xmlResponseElement = null;

			try 
			{
				_traceAssist.TraceRequest(xmlRequest.OuterXml);
				_traceAssist.TraceMethodEntry(_typeName, functionName);

				XmlDocument xmlOut = new XmlDocument();
				xmlResponseElement = xmlOut.CreateElement("RESPONSE");
				xmlOut.AppendChild(xmlResponseElement);
				xmlResponseElement.SetAttribute("TYPE", "SUCCESS");

				XmlNode xmlRequestNode = null;
				if (xmlRequest.Name == _rootNodeName)
				{
					xmlRequestNode = xmlRequest;
				}
				else
				{
					xmlRequestNode = xmlRequest.GetElementsByTagName(_rootNodeName)[0];
				}

				if (xmlRequestNode == null)
				{
					throw new MissingPrimaryTagException(_rootNodeName + " tag not found");
				}

				xmlResponseElement.AppendChild(_templateHandlerDO.FindAvailableTemplates((XmlElement)xmlRequestNode));

				_traceAssist.TraceResponse(xmlOut.OuterXml);
				_traceAssist.TraceMethodExit(_typeName, functionName);
			}
			catch (Exception exception)
			{
				throw new ErrAssistException(exception);
			}

			return xmlResponseElement;
		}

		/// <summary>
		/// Finds all the templates with: a security level &lt;= to the supplied security level, 
		/// a stage &gt;= to the supplied stage, a language matching the supplied language.
		/// </summary>
		/// <param name="request">Xml request data stream containing security level, appication stage.</param>
		/// <returns>
		/// Xml containing list of template ids, names and descriptions.
		/// </returns>
		/// <remarks>
		/// The <paramref name="request"/> parameter has the form:
		/// <code>
		/// &lt;REQUEST&gt;
		///	 &lt;TEMPLATE&gt;
		///		 &lt;SECURITYLEVEL/&gt;
		///		 &lt;STAGE/&gt;
		///		 &lt;LANGUAGE/&gt;
		///	 &lt;/TEMPLATE&gt;
		/// &lt;/REQUEST&gt;
		/// </code>
		/// </remarks>
		public string FindAvailableTemplates(string request)
		{
			string response = "";
			const string functionName = "FindAvailableTemplates";

			try
			{
				// ik_20030210
				ContextUtility.SetComplete();

				_traceAssist.TraceMethodEntry(_typeName, functionName);

				// Create default response block
				XmlDocument xmlOut = new XmlDocument();
				XmlElement xmlResponseElement = xmlOut.CreateElement("RESPONSE");
				xmlOut.AppendChild(xmlResponseElement);
				xmlResponseElement.SetAttribute("TYPE", "SUCCESS");
				XmlDocument xmlIn = XmlAssist.Load(request);

				// Delegate to FreeThreadedDomDocument40 based method and attach returned data to our response
				XmlNode xmlTempResponseNode = FindAvailableTemplates(xmlIn.DocumentElement);
				ErrAssistException.CheckXmlResponseNode(xmlTempResponseNode, xmlResponseElement, true);
				XmlAssist.AttachResponseData((XmlNode)xmlResponseElement, (XmlElement)xmlTempResponseNode);
				response = xmlResponseElement.OuterXml;

				_traceAssist.TraceMethodExit(_typeName, functionName);
			}
			catch (Exception exception)
			{
				response = new ErrAssistException(exception).CreateErrorResponse();
				_traceAssist.TraceIdentErrorResponse(ref response);
			}

			return response;
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
