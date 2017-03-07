/* History
 * 
 * Prog		Date		Defect		Description
 * MV		17/10/2005	MAR182		Created
 * GHun		30/11/2005	MAR609		Fixed GuidStringToByteArray, and added functions shared between inbound and outbound
 * GHun		15/12/2005	MAR855		Moved AddMessageToQueue processing to FirstTitleNTxBO
 * PSC		10/01/2006	MAR961		Use omLogging wrapper
 * RF		09/03/2006  MAR1392		Writing MOF message to ApplicationFirstTitle table is timing out 
 * MHeys	10/04/2006	MAR1603		Terminate COM+ objects cleanly
 * MHeys	21/04/206	MAR1655		Terminate COM+ objects cleanly pass2
 * HMA      05/05/2006  MAR1713		Correct error test in Unlock Application
*/
using omApp;
using System;
using System.Text;
using System.Xml;
//mcs using omApp;
using Vertex.Fsd.Omiga.Core;
//mcs using omMQ;
using Vertex.Fsd.Omiga.omFirstTitle.XML;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 10/01/2006 MAR961

namespace Vertex.Fsd.Omiga.omFirstTitle.General
{
	public class GeneralManager
	{
		
		public byte[] GuidStringToByteArray(string Guid)
		{
			//string s1 =    "0x1234567890abcd567890123445677889";

			//byte[] bytearray = new byte[strGuid.Length];
			byte[] bytearray = new byte[16];

			for (int i = 0; i < Guid.Length-1; i += 2)	//MAR609 GHun
			{
				string hex = Guid.Substring(i,2);
				bytearray[(i/2)] = byte.Parse(hex, System.Globalization.NumberStyles.HexNumber);	//MAR609 GHun
			}
			return bytearray;
		}

		//MAR609 GHun
		// PSC 10/01/2006 MAR961
		public bool CreateApplicationLock(string ApplicationNumber, string userId, string unitId, string userAuthorityLevel, string channelId, omLogging.Logger _logger)
		{
			const string cFunctionName = "CreateApplicationLock";
		
			try
			{
				//Build Request for Application Lock
				XmlDocument xdCreateLock = new XmlDocument();
				XmlDocument xdResponse = new XmlDocument();

				XmlElement xeRequest = (XmlElement) xdCreateLock.CreateElement("REQUEST");
				xdCreateLock.AppendChild(xeRequest);
				
				xeRequest.SetAttribute("UNITID", unitId);
				xeRequest.SetAttribute("USERID", userId);
				xeRequest.SetAttribute("USERAUTHORITYLEVEL", userAuthorityLevel);
				xeRequest.SetAttribute("CHANNELID", channelId);

				XmlNode xnAppLockNode = xdCreateLock.CreateElement("APPLICATIONLOCK");
				xeRequest.AppendChild(xnAppLockNode);
			
				//MAR609 GHun Fix CreateLock request XML
				XmlElement xElement = xdCreateLock.CreateElement("APPLICATIONNUMBER");
				xElement.InnerText = ApplicationNumber;
				xnAppLockNode.AppendChild(xElement);
				xElement = xdCreateLock.CreateElement("UNITID");
				xElement.InnerText = unitId;
				xnAppLockNode.AppendChild(xElement);
				xElement = xdCreateLock.CreateElement("USERID");
				xElement.InnerText = userId;
				xnAppLockNode.AppendChild(xElement);
				xElement = xdCreateLock.CreateElement("MACHINEID");
				xElement.InnerText = "";
				xnAppLockNode.AppendChild(xElement);
				xElement = xdCreateLock.CreateElement("CHANNELID");
				xElement.InnerText = channelId;
				xnAppLockNode.AppendChild(xElement);
				//MAR609 End

				//Call ApplicationManangerBO.CreateLock()
				//mcs ApplicationManagerBOClass AppManagerBO = new ApplicationManagerBOClass();
                ApplicationManagerBO AppManagerBO = new ApplicationManagerBO();
                // MAR1603 M Heys 10/04/2006 start
                try
				{
					xdResponse.LoadXml(AppManagerBO.CreateLock(xdCreateLock.OuterXml));
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(AppManagerBO);
				}
				// MAR1603 M Heys 10/04/2006 end

				string ResponseType = xdResponse.DocumentElement.GetAttribute("TYPE");
				if (ResponseType != "SUCCESS")
				{
					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug(cFunctionName + " error: " + xdResponse.OuterXml);
					}
					return false;	//MAR609 GHun return false on failure
				}
			}
			catch (Exception ex)
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(cFunctionName + " error: " + ex.Message);
				}
				return false;	//MAR609 GHun return false on failure
			}
			
			return true;
		}

		// PSC 10/01/2006 MAR961
		// MAR1713  Add extra information to Unlock Application request
		public bool UnlockApplication(string ApplicationNumber, string userId, string machineId, omLogging.Logger _logger)
		{
			const string cFunctionName  = "UnlockApplication";
			try
			{
				//Build Request for Application unLock
				XmlDocument xdUnLock = new XmlDocument();
				XmlDocument xdResponse = new XmlDocument();
				XmlNode xnAppLockNode;

				XmlElement xeRequest = (XmlElement) xdUnLock.CreateElement("REQUEST");
				
				xdUnLock.AppendChild(xeRequest);
				
				xeRequest.SetAttribute("USERID", userId);
				xeRequest.SetAttribute("MACHINEID", machineId);

				xnAppLockNode = xdUnLock.CreateElement("APPLICATION");
				xeRequest.AppendChild(xnAppLockNode);

				XmlNode xnAppNum =  xdUnLock.CreateElement("APPLICATIONNUMBER");
				xnAppNum.InnerText = ApplicationNumber;
				xnAppLockNode.AppendChild(xnAppNum);

                //mcs ApplicationManagerBOClass AppManagerBO = new ApplicationManagerBOClass();
                ApplicationManagerBO AppManagerBO = new ApplicationManagerBO();
				
				//MAR1655 M Heys 21/04/2006 start
				try
				{
					xdResponse.LoadXml(AppManagerBO.UnlockApplicationAndCustomers(xdUnLock.OuterXml));
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(AppManagerBO);
				}
				//MAR1655 M Heys 21/04/2006 end
				string responseType = xdResponse.DocumentElement.GetAttribute("TYPE");
				if (responseType != "SUCCESS")     // MAR1713
				{
					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug(cFunctionName + " error: " + xdResponse.OuterXml);
					}
					return false;	//MAR609 GHun return false on failure
				}
			}
			catch(Exception ex)
			{
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(cFunctionName + " error:" + ex.Message);
				}
				return false;	//MAR609 GHun return false on failure
			}

			return true;
		}
		//MAR609 End

		//MAR855 GHun
		// PSC 10/01/2006 MAR961
		// RF 09/02/2006 MAR1392 Pass process context not logger
		//public bool AddMessageToQueue(string strData, bool bAddToOutbound, bool bAddDelay, string interfaceName, string applicationNumber, string userId, string unitId, string userAuthorityLevel, string channelId, omLogging.Logger logger)
		public bool AddMessageToQueue(string strData, bool bAddToOutbound, 
			bool bAddDelay, string interfaceName, string applicationNumber, 
			string userId, string unitId, string userAuthorityLevel, 
			string channelId, string processContext)
		{
			FirstTitleNTxBO ntxBO = new FirstTitleNTxBO();
			//return ntxBO.AddMessageToQueue(strData, bAddToOutbound, bAddDelay, interfaceName, applicationNumber, userId, unitId, userAuthorityLevel, channelId, logger);
			return ntxBO.AddMessageToQueue(strData, bAddToOutbound, bAddDelay, 
				interfaceName, applicationNumber, userId, unitId, 
				userAuthorityLevel, channelId, processContext);
		}
		//MAR855 End

	}
}
