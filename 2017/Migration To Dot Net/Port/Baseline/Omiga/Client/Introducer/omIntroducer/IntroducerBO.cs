/*
--------------------------------------------------------------------------------------------
Workfile:			IntroducerBO.cs
Copyright:			Copyright ©2006 Vertex Financial Services

Description:		COM XML Wrapper encapsulating introducer
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PSC		23/10/2006	EP2_22   Created
PSC		21/11/2006	EP2_138	 Add GetIntroducerPipeline
SR		26/01/2007	EP2_987
SR		31/01/2007	EP2_103	 log the request string in functions 'GetIntroducerUserFirms' and 'AuthenticateFirm'
SR		31/01/2007	EP2_103	 log the request string in functions 'ValidateBroker' 		
--------------------------------------------------------------------------------------------
*/
using System;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Xml;
using System.IO;
using Vertex.Fsd.Omiga.Core;
using Vertex.Fsd.Omiga.omLogging;

namespace Vertex.Fsd.Omiga.omIntroducer
{
	/// <summary>
	/// COM callable wrapper for Introducer
	/// </summary>
	[ComVisible(true)]
	[ProgId("omIntroducer.IntroducerBO")]
	[Guid("4386F797-5BAA-4b8a-A09A-E54EFAA3380D")]
	public class IntroducerBO
	{
		private static readonly Logger _logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		public IntroducerBO()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		/// <summary>
		/// External XML interface to the component
		/// </summary>
		/// <param name="request">XML string containing the request</param>
		/// <returns>XML string containing the response</returns>
		public string omRequest(string request)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;
			string response = null;
			XmlTextReader reader = null;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				// Find the OPERATION and TYPE attributes - these determine how the message is processed.
				reader = new XmlTextReader(new StringReader(request));

				if (reader.Read() && reader.NodeType == XmlNodeType.Element && reader.Name == "REQUEST")
				{
					string operation = null;

					if (reader.MoveToAttribute("OPERATION")) 
					{
						operation = reader.Value;
					}
					else 
					{
						throw new OmigaException("Invalid request: missing OPERATION attribute.");
					}
					
					response = ProcessRequest(operation, request);
					
					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Response: {1}.", functionName, response);
					}
					
					return response;
				}
				else
				{
					throw new OmigaException("Invalid request: missing REQUEST root element.");
				}
			}
			catch (OmigaException omEx)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(functionName + ": Error processing request.", omEx);
				}

				return omEx.ToOmiga4Response();
			}
			catch (Exception ex)
			{
				OmigaException omEx = new OmigaException("Error processing request.", ex);
				
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(omEx.Message, omEx);
				}

				return omEx.ToOmiga4Response();
			}
			finally
			{
				if (reader != null)
				{
					reader.Close();
				}
				
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Passes the request onto the correct method given the operation
		/// </summary>
		/// <param name="operation">The operation to be performed</param>
		/// <param name="request">XML string containing the request</param>
		/// <returns>XML string containing the response</returns>
		private string ProcessRequest(string operation, string request)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;
			
			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting. Operation = {1}", functionName, operation);
			}

			try
			{
				string response = null;
			
				switch (operation.ToUpper())
				{
					case "GETINTRODUCERUSERFIRMS":
						response = GetIntroducerUserFirms(request);
						break;
					case "AUTHENTICATEFIRM":
						response = AuthenticateFirm(request);
						break;
					case "VALIDATEBROKER":
						response = ValidateBroker(request);
						break;
					// PSC 21/11/2006 EP2_138 - Start
					case "GETINTRODUCERPIPELINE":
						response = GetIntroducerPipeline(request);
						break;
					// PSC 21/11/2006 EP2_138 - End
					default:
					{
						OmigaException omEx = new OmigaException("Invalid IntroducerBO.ProcessRequest operation: " + operation + ".");
					
						if (_logger.IsErrorEnabled)
						{
							_logger.Error(functionName + ": " + omEx.Message, omEx);
						}

						throw omEx;
					}
				}

				return response;			
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Validates an introducer's credentials and returns their linked firms
		/// </summary>
		/// <param name="request">XML string containing the request</param>
		/// <returns>XML string containing the response</returns>
		private string GetIntroducerUserFirms(string request)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
				_logger.DebugFormat("{0}: Request:{1}", functionName, request);
			}

			try
			{
				XmlDocumentEx requestDoc = new XmlDocumentEx();
				requestDoc.LoadXml(request);

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Retrieving parameters from request.", functionName);
				}

				XmlElementEx xmlRequest = (XmlElementEx)requestDoc.SelectMandatorySingleNode("REQUEST");
				XmlElementEx xmlCredentials = (XmlElementEx)xmlRequest.SelectMandatorySingleNode("INTRODUCERCREDENTIALS");
				string userName = xmlCredentials.GetMandatoryAttribute("INTRODUCERUSERNAME");
				string password = xmlCredentials.GetMandatoryAttribute("INTRODUCERUSERPASSWORD");
				int loginAttempts = xmlCredentials.GetMandatoryAttributeAsInteger("LOGINATTEMPTS");

				Introducer introducer = new Introducer();

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Calling Introducer.ValidateIntroducerAndGetFirms.", functionName);
				}
				
				string introducerFirms = introducer.ValidateIntroducerAndGetFirms(userName, password, loginAttempts); 

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Returned from Introducer.ValidateIntroducerAndGetFirms. Response: {1}", functionName, introducerFirms);
				}


				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Returning: {1}", functionName, introducerFirms);
				}

				return introducerFirms;
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Authenticates an introducer against a firm 
		/// </summary>
		/// <param name="request">XML string containing the request</param>
		/// <returns>XML string containing the response</returns>
		private string AuthenticateFirm(string request)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
				_logger.DebugFormat("{0}: Request:{1}", functionName, request);
			}

			try
			{
				XmlDocumentEx requestDoc = new XmlDocumentEx();
				requestDoc.LoadXml(request);
			
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Retrieving parameters from request.", functionName);
				}

				FirmType firmType = FirmType.PrincipalFirm;

				XmlElementEx xmlRequest = (XmlElementEx)requestDoc.SelectMandatorySingleNode("REQUEST");
				XmlElementEx xmlFirm = (XmlElementEx)xmlRequest.SelectSingleNode("ARFIRM");

				// No ARFIRM node found so look for PRINCIPALFIRM
				if (xmlFirm == null)
				{
					xmlFirm = (XmlElementEx)xmlRequest.SelectSingleNode("PRINCIPALFIRM");
				}
				else
				{
					firmType = FirmType.ARFirm;
				}

				// Neither ARFIRM or PRINCIPALFIRM found
				if (xmlFirm == null)
				{
					throw new OmigaException("Invalid request: missing ARFIRM or PRINCIPALFIRM node.");
				}

				string introducerId = xmlFirm.GetMandatoryAttribute("INTRODUCERID"); 
				string firmId = xmlFirm.GetMandatoryAttribute("FIRMID");

				Introducer introducer = new Introducer();
				
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Calling Introducer.AuthenticateFirm.", functionName);
				}

				string firmPermissions = introducer.AuthenticateFirm(introducerId, firmId, firmType);
 
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Returned from Introducer.AuthenticateFirm. Response: {1}", functionName, firmPermissions);
				}

				firmPermissions = "<RESPONSE TYPE=\"SUCCESS\">" + firmPermissions + "</RESPONSE>";

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Returning: {1}", functionName, firmPermissions);
				}

				return firmPermissions;
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Validates a brokers creduentials
		/// </summary>
		/// <param name="request">XML string containing the request</param>
		/// <returns>XML string containing the response</returns>
		private string ValidateBroker(string request)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
				_logger.DebugFormat("{0}: Request:{1}", functionName, request);
			}

			try
			{
				XmlDocumentEx requestDoc = new XmlDocumentEx();
				requestDoc.LoadXml(request);

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Retrieving parameters from request.", functionName);
				}

				XmlElementEx xmlRequest = (XmlElementEx)requestDoc.SelectMandatorySingleNode("REQUEST");
				XmlElementEx xmlCredentials = (XmlElementEx)xmlRequest.SelectMandatorySingleNode("BROKERCREDENTIALS");
				string brokerIdentification = xmlCredentials.GetAttribute("BROKERIDENTIFICATION");
				string brokerFsaRef = xmlCredentials.GetAttribute("BROKERFSAREF");

				if (brokerIdentification.Length == 0 && brokerFsaRef.Length == 0)
				{
					throw new OmigaException("Invalid request: missing BROKERIDENTIFICATION or BROKERFSAREF attribute.");
				}

				Introducer introducer = new Introducer();
				
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Calling Introducer.ValidateBroker.", functionName);
				}

				string firmPermissions = introducer.ValidateBroker(brokerIdentification, brokerFsaRef);
		
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Returned from Introducer.ValidateBroker. Response: {1}", functionName, firmPermissions);
				}

				firmPermissions = "<RESPONSE TYPE=\"SUCCESS\">" + firmPermissions + "</RESPONSE>";

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Returning: {1}", functionName, firmPermissions);
				}

				return firmPermissions;
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		// PSC 21/11/2006 EP2_138 - Start
		/// <summary>
		/// Retrieves a list of pipeline cases linked to an introducer
		/// </summary>
		/// <param name="request">XML string containing the request</param>
		/// <returns>XML string containing the response</returns>
		private string GetIntroducerPipeline(string request)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
				_logger.DebugFormat("{0}: Request:{1}", functionName, request);
			}

			try
			{
				XmlDocumentEx requestDoc = new XmlDocumentEx();
				requestDoc.LoadXml(request);

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Retrieving parameters from request.", functionName);
				}

				XmlElementEx requestElement = (XmlElementEx)requestDoc.SelectMandatorySingleNode("REQUEST");
				XmlElementEx pipelineSearchElement  = (XmlElementEx)requestElement.SelectMandatorySingleNode("PIPELINESEARCH");
				string userId = requestElement.GetMandatoryAttribute("USERID");
				string unitId = requestElement.GetMandatoryAttribute("UNITID");
				string introducerId = pipelineSearchElement.GetMandatoryAttribute("INTRODUCERID");
				bool includePrevious = pipelineSearchElement.GetAttributeAsBool("INCLUDEPREVIOUS");
				bool includeCancelled = pipelineSearchElement.GetAttributeAsBool("INCLUDECANCELLED");
				bool includeFirm = pipelineSearchElement.GetAttributeAsBool("INCLUDEFIRM");

				Introducer introducer = new Introducer();
				
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Calling Introducer.GetIntroducerPipeline.", functionName);
				}

				string introducerPipeline = introducer.GetIntroducerPipeline(userId, 
																			 unitId,
																			 introducerId,
																			 includePrevious,
																			 includeCancelled,
																			 includeFirm);
		
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Returned from Introducer.GetIntroducerPipeline. Response: {1}", functionName, introducerPipeline);
				}

				introducerPipeline = "<RESPONSE TYPE=\"SUCCESS\">" + introducerPipeline + "</RESPONSE>";

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Returning: {1}", functionName, introducerPipeline);
				}

				return introducerPipeline;
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}
		// PSC 21/11/2006 EP2_138 - End
	}
}
