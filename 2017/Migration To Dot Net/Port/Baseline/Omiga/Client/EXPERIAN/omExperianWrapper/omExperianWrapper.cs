/*
--------------------------------------------------------------------------------------------
Workfile:			omExperianWrapper.cs
Copyright:			Copyright © Vertex Financial Services Ltd 2005

Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
LDM		31/01/2006  EP6 Epsom New wrapper object to call experian and provide a stub option
SAB		07/03/2006  EP6 Added code to make real call and restructured stub code
SAB		21/03/2006	EP258 Incorporates Experian errors in the Omiga error message text.
LDM		22/03/2006  EP6 Put in the new "upgrade credit check" call to Experian
LDM		22/05/2006  EP597 Put in Adrians performance monitoring, for load testing.
--------------------------------------------------------------------------------------------
*/

using System;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Serialization; // for XmlSerializer
using System.IO;				// for reading file
using System.Net;
using System.Text;
using System.Threading;
using System.Reflection;
using System.Web;
using Vertex.Fsd.Omiga.Core;
using omLogging = Vertex.Fsd.Omiga.omLogging;

namespace Vertex.Fsd.Omiga.omExperianWrapper
{
	public class ExperianWrapperBO
	{
		public ExperianWrapperBO()
		{
		}

		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
		private const string _kyc = "KYC";
		private const string _cc = "CREDITCHECK";
		private const string _hunter = "HUNTER";
		private const string _target = "ADDRESSTARGET";		
		private const string _upgrade = "UPGRADECC";		

		public void CallExperian(string callType, XmlNode requestNode, ref XmlDocument responseXML, string applicationNumber)
		{
			Microsoft.Win32.RegistryKey connectionKey;
			bool sendAStub = false;
			try
			{
				// Set the call type to upper case for ease of validation
				callType = callType.ToUpper();
				
				// Log the request and applicationnumber
				if(_logger.IsDebugEnabled) 
				{
					
					string logAppNo = "Unknown";
					if (applicationNumber != null && applicationNumber.Length > 0)
					{
						logAppNo = applicationNumber;
					}
					omLogging.ThreadContext.Properties["ApplicationNumber"] = logAppNo;	
					
					_logger.Debug("Starting " + MethodInfo.GetCurrentMethod().Name + 
						"- Call Type: " + callType + " Call data: " + requestNode.OuterXml);
				}

				// Find the Experian registry key
				connectionKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey ("Software\\Marlborough Stirling\\Experian", false);

				// If the registry includes the key check if a stubbed response should be returned
				if (connectionKey != null)
				{
					string useCCStub = (string)connectionKey.GetValue("UseCCStub Y/N");
					string useKYCStub = (string)connectionKey.GetValue("UseKYCStub Y/N");					
					string useHunterStub = (string)connectionKey.GetValue("UseHunterStub Y/N");
					if (((callType == _cc || callType == _target) && useCCStub != null && useCCStub == "Y") 
						|| (callType == _kyc && useCCStub != null && useKYCStub == "Y")
						|| (callType == _hunter && useHunterStub != null && useHunterStub == "Y"))
					{
						sendAStub = true;
					}
				}

				if (sendAStub == true)
				{
					sendStubbedResponse(callType, requestNode, ref responseXML, applicationNumber);
				}
				else
				{
					//LDM 22/05/2006 EP597
					try
					{
						sendRequestToExperian(callType, requestNode, ref responseXML, applicationNumber);
						Performance.ExperianSuccessful.Increment();
						Performance.ExperianSuccessfulPercentageBase.Increment();
						Performance.ExperianSuccessfulPercentage.Increment();
					}
					catch
					{
						Performance.ExperianUnsuccessful.Increment();
						Performance.ExperianSuccessfulPercentageBase.Increment();
						throw;
					}
				}
			}
			catch(Exception ex)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": exception occurred: " + ex.Message.ToString() );
				}
				throw new OmigaException(ex);
			}
			finally
			{
			}
		}
		
		private void sendStubbedResponse(string callType, XmlNode requestNode, ref XmlDocument responseXML, string applicationNumber)
		{
			Microsoft.Win32.RegistryKey connectionKey;
			XmlDocument responseDocStubXML = new XmlDocument();

			try
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug("Starting " + MethodInfo.GetCurrentMethod().Name);
				}
				
				// Check to see if we have a resgistry key containing the stub settings
				connectionKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey ("Software\\Marlborough Stirling\\Experian", false);
				
				if (connectionKey != null)
				{
					int experianWaitTime = 0;
					string stubPath = "";
					string stubFileName = "";
					
					experianWaitTime = Convert.ToInt16(connectionKey.GetValue("StubWaitTimeMs"));
					stubPath = (string)connectionKey.GetValue("StubPath");
					stubFileName = (string)connectionKey.GetValue(callType + "FileName");
					
					if (_logger.IsDebugEnabled)
					{
						_logger.Debug("Registry settings : " + 
							" Experian Stub Wait Time(Ms) : " + experianWaitTime +
							" Experian Stub Path : " + stubPath +
							" Experian Stub File Name : " + stubFileName);
					}

					// Simulate a call to Experian					
					if(experianWaitTime > 0)
					{
						Thread.Sleep(experianWaitTime);
					}
						
					// Load the stub					
					if(	stubPath != null && stubPath.Length > 0 && stubFileName != null && stubFileName.Length > 0 )
					{
						responseXML.Load(stubPath + "\\" + stubFileName + ".xml");
					}
					else 
					{
						throw new OmigaException("Stub path not found");
					}					
					
					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug("Stub Response: " + responseXML.OuterXml);
					}
				}
			}
			catch(Exception ex)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": exception occurred: " + ex.Message.ToString() );
				}
				
				throw new OmigaException(ex);
			}
			finally
			{
				responseDocStubXML = null;
			}
		}

		private void sendRequestToExperian(string callType, XmlNode requestNode, ref XmlDocument responseXML, string applicationNumber)
		{
			XmlNode experianRequestNode;
			XmlNode experianErrNode;
			HttpWebRequest HTTPRequest;
			HttpWebResponse HTTPResponse;
			StreamReader responseStream;
			try 
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug("Starting " + MethodInfo.GetCurrentMethod().Name); 
					_logger.Debug("AS1!");//LDM 22/05/2006 EP597
				}

				// Identify the Servlet to call
				string uriRegKey;
				string experianServlet;
				
				switch(callType)       
				{         
					case _cc:   
						uriRegKey = "ExperianCCServletUri";
						break;                  
					case _target:   
						uriRegKey = "ExperianCCServletUri";
						break;
					case _upgrade:   
						uriRegKey = "ExperianCCServletUri";
						break;
					case _hunter:
						uriRegKey = "ExperianCCServletUri";
						break;
					case _kyc:            
						uriRegKey = "ExperianKYCServletUri";
						break;
					default:            
						throw new OmigaException("Invalid Experian call type.");
				}
				experianServlet = new GlobalParameter(uriRegKey).String;

				// Set connection propertties
				Uri servletUri = new Uri(experianServlet, false);
				HTTPRequest = (HttpWebRequest) WebRequest.Create(servletUri);

				// Attach the data - we need to use HTTP Post
				string postData = requestNode.OuterXml; // No need to URLEncode as XML contentType
				HTTPRequest.Method="POST";
				byte [] postBuffer = System.Text.Encoding.UTF8.GetBytes(postData);
				HTTPRequest.ContentLength = postBuffer.Length;
				HTTPRequest.ContentType = "application/xml";
		
				double experianTimeout = new GlobalParameter("ExperianCallTimeout").Amount;
				
				HTTPRequest.Timeout = Convert.ToInt16(experianTimeout);

				bool useProxyServer = new GlobalParameter("ExperianUseProxyServer").Boolean;
				if (useProxyServer)
				{
					string proxyServer = new GlobalParameter("ExperianProxyServer").String;
					WebProxy proxy = new WebProxy(proxyServer, true);

					string user = new GlobalParameter("ExperianProxyUser").String;
					string domain = new GlobalParameter("ExperianProxyDomain").String;
					string password = new GlobalParameter("ExperianProxyPassword").String;

					CredentialCache cc = new CredentialCache();
					cc.Add(proxy.Address, "Negotiate", new NetworkCredential(user, password, domain));
					cc.Add(proxy.Address, "Basic", new NetworkCredential(user, password, domain));
					proxy.Credentials = cc;
					HTTPRequest.Proxy = proxy;
				}

				// Make the Experian call 

				//LDM 22/05/2006 EP597
				DateTime start = DateTime.Now;
				Performance.ExperianActive.Increment();

				try
				{
				// Write the post data to the Servlet server ...
				Stream postDataStream = HTTPRequest.GetRequestStream();
				postDataStream.Write(postBuffer,0,postBuffer.Length);
				postDataStream.Close();

				// Retrieve the response and extract it ...
				HTTPResponse = (HttpWebResponse) HTTPRequest.GetResponse();
				Encoding enc = System.Text.Encoding.UTF8;
				responseStream = new StreamReader(HTTPResponse.GetResponseStream(),enc);
				}
				finally
				{
					Performance.ExperianActive.Decrement();
					TimeSpan duration = DateTime.Now - start;
					if (_logger.IsDebugEnabled) _logger.Debug("Duration: " + duration.Seconds.ToString()); 
					Performance.ExperianRequestAverageTime.IncrementBy(duration.Seconds * PerformanceCounterEx.Frequency);
					Performance.ExperianRequestAverageTimeBase.Increment();
				}

				// Load the response XML into the responseXML document
				string experianXMLResponse = responseStream.ReadToEnd();

				if (_logger.IsDebugEnabled) 
				{
					_logger.Debug("Experian response: " + experianXMLResponse);
				}

				// Has any response been returned ?
				if (experianXMLResponse == "")
				{
					throw new OmigaException("Experian did not return a response.");
				} 
				else
				{
					// If an error is returned it can contain XML embedded in the error message which does not
					// remain encoded following HTTPRequest.GetResponse().  As this causes an error when it 
					// is initially parsed, I URLEncode it.
					string tagName = "ERR1";
					string startTag = "<" + tagName + ">";
					string endTag = "</" + tagName + ">";
					int startTagPos = experianXMLResponse.IndexOf(startTag);
					int endTagPos = experianXMLResponse.IndexOf(endTag);
					if (startTagPos >= 0 && endTagPos >= 0 && (startTagPos + startTag.Length) < endTagPos)
					{
						// Search for the message tag
						tagName = "MESSAGE";
						startTag = "<" + tagName + ">";
						endTag = "</" + tagName + ">";
						startTagPos = experianXMLResponse.IndexOf(startTag);
						endTagPos = experianXMLResponse.IndexOf(endTag);
						if (startTagPos >= 0 && endTagPos >= 0 && (startTagPos + startTag.Length) < endTagPos)
						{
							// We are not interested in encoding the start tag so ignore it
							int messageStringInd = startTagPos + startTag.Length;
							string messageText = HttpUtility.UrlEncode(experianXMLResponse.Substring(messageStringInd,endTagPos-messageStringInd)); 
							experianXMLResponse = experianXMLResponse.Substring(0, messageStringInd) 
								+ messageText 
								+ experianXMLResponse.Substring(endTagPos, experianXMLResponse.Length-endTagPos);
						}
					}
					
					responseXML.LoadXml(experianXMLResponse);

					// Does the response contain a request node ?
					experianRequestNode = responseXML.SelectSingleNode("//GEODS/REQUEST");
					if (experianRequestNode == null)
					{
						throw new OmigaException("The Experian response did not contain a RESPONSE node");
					}
					else
					{
						// Has Experian returned any errors
						experianRequestNode = responseXML.SelectSingleNode("//GEODS/REQUEST[@success='N']");
						
						// 21/03/2006 - EP258 - Incorporates Experian error text in the Omiga error message
						if (experianRequestNode != null)
						{
							string experianErrorMessage = "No error message reported";
							experianErrNode = responseXML.SelectSingleNode("//ERR1/MESSAGE");
							
							if (experianErrNode != null)
							{
								experianErrorMessage = HttpUtility.UrlDecode(experianErrNode.InnerText);
							}								
													
							throw new OmigaException("The Experian software failed with the following error: " + experianErrorMessage);
						}
					}
				}
				HTTPResponse.Close();
				responseStream.Close();
			}
			catch (Exception ex)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error(MethodInfo.GetCurrentMethod().Name + 
						": exception occurred: " + ex.Message.ToString() );
				}
				throw new OmigaException(ex);
			}				
			finally
			{
				HTTPRequest = null;
				HTTPResponse = null;				
				responseStream = null;
				experianRequestNode = null;
				experianErrNode = null;
			}			
		}
	}
}