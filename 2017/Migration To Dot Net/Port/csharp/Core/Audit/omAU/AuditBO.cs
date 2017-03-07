using System;
using System.Xml;
using System.EnterpriseServices;
using System.Runtime.InteropServices; 
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.omAU
{
	[ComVisible(true)]
    [ProgId("omAU.AuditBO")]
    [Guid("250F7B45-BFD0-4837-B17A-413C92001DC2")]
    [Transaction(TransactionOption.Supported)]
    public class AuditBO : ServicedComponent, IAuditBO
    {
        // Workfile:      AuditBO.cls
        // Copyright:     Copyright © 1999 Marlborough Stirling
        // Description:   Business objects class for omAU
        // 
        // Dependencies:  AuditTxBO, AuditDO
        // Issues:        Instancing:         MultiUse
        // MTSTransactionMode: UsesTransaction
        // ------------------------------------------------------------------------------------------
        // History:
        // 
        // Prog   Date        Description
        // RF     30/09/99    Created based on template version 30/09/99.
        // RF     27/01/00    Pick up ScriptInterface object from omBase.
        // MC     16/05/00    SYS0210 - Synchronise the password change date/time with corresponding
        // access audit record
        // CL     18/10/00    Core00004 Modifications made to conform to coding templates
        // MV     06/03/01    SYS2001: changed  the Return Data from xmlTempResponseNode.xml to xmlResponseElem.xml
        // ------------------------------------------------------------------------------------------
        // BBG Specific History:
        // 
        // Prog   Date        Description
        // TK     30/11/2004  E2EM00002504 - Performance related fixes.
        // ------------------------------------------------------------------------------------------
        
        private const string cstrROOT_NODE_NAME = "ACCESSAUDIT";
		private bool _isObjectContext = false;
       
        XmlNode IAuditBO.Validate(XmlElement vxmlRequest, IAuditBOMethod veboMethod) 
        {
            
            XmlNode IAuditBO_Validate = null;
            XmlDocument xmlOut = null;
            XmlElement xmlResponseElem = null;
            XmlElement xmlRequestElem = null;
            
            string strUserId = String.Empty;
            string strAUDITRECORDTYPE = String.Empty;
            string strSUCCESSINDICATOR = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:  performs hard coded validation for this object
            // pass:         vxmlRequest  xml Request data stream containing details of action to be
            // performed and data required
            // This is the full request block as received from the client
            // return:       xml Response Node
            // ------------------------------------------------------------------------------------------
            
            try 
            {
                // Create a new DOM document and add a response node "SUCCESS"
                // this will be used if we have a success from the DO
                
                xmlOut = new XmlDocument();
                
                xmlResponseElem = xmlOut.CreateElement("RESPONSE");
                xmlOut.AppendChild(xmlResponseElem);
                xmlResponseElem.SetAttribute("TYPE", "SUCCESS");

                // Find the node REQUEST within the imported xml
                if (vxmlRequest.Name == "REQUEST")
                {
                    xmlRequestElem = vxmlRequest;
                }
                else
                {
                    xmlRequestElem =(XmlElement) XmlAssist.GetNode(vxmlRequest, ".//REQUEST");
                }

                // Get the value of the attribute USERID withing the REQUEST node
                bool blnOutValue;
                strUserId = XmlAssist.GetAttributeValue(xmlRequestElem, "REQUEST", "USERID", out blnOutValue);
                // If no value is found in the above node then throw an error
                if (strUserId.Length == 0)
                {
                    throw new ErrAssistException(OMIGAERROR.InvalidParameter, "Expected USERID attribute in REQUEST tag"); 
                }
            

                // Get the text of the AUDITRECORDTYPE node
                strAUDITRECORDTYPE = XmlAssist.GetMandatoryNodeText((XmlNode)xmlRequestElem, ".//AUDITRECORDTYPE");
                
                // Get the text of the SUCCESSINDICATOR node
                strAUDITRECORDTYPE = XmlAssist.GetMandatoryNodeText((XmlNode)xmlRequestElem, ".//SUCCESSINDICATOR");
                if (_isObjectContext) ContextUtil.SetComplete();
                IAuditBO_Validate = xmlResponseElem;

                return IAuditBO_Validate;
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex); 
                if (exception.IsWarning())
                {
                    XmlElement tmpxmlResponseElem = xmlResponseElem;
                    exception.AddWarning(tmpxmlResponseElem);
                }
                
                if (exception.IsSystemError())
                {
                    exception.LogError(exception.Message, System.Diagnostics.EventLogEntryType.Error);
                }

				if (_isObjectContext) ContextUtil.SetAbort();
            }
            return IAuditBO_Validate;
        }

        public string CreateAccessAudit(string vstrRequest) 
        {
            string CreateAccessAudit = String.Empty;
           
            XmlDocument xmlIn = new XmlDocument();
            XmlDocument xmlOut = new XmlDocument();
            XmlElement xmlResponseElem = null;
            XmlNode xmlTempResponseNode = null;
            // header ----------------------------------------------------------------------------------
            // description:
            // Create an accessaudit record.
            // For a change to password, a change password record is also created.
            // For an application access or release, an applicationaccess record is also created.
            // pass:
            // vstrRequest
            // xml Request data stream containing data to be persisted.
            // USERID and MACHINEID from REQUEST element are passed down to DO.
            // Format:
            // <REQUEST USERID= MACHINEID=>
            // <ACCESSAUDIT>
            // <AUDITRECORDTYPE></AUDITRECORDTYPE>
            // <SUCCESSINDICATOR></SUCCESSINDICATOR>
            // <ONBEHALFOFUSERID>Optional</ONBEHALFOFUSERID>
            // <APPLICATIONNUMBER></APPLICATIONNUMBER>
            // <PASSWORDCREATIONDATE>Optional</PASSWORDCREATIONDATE>
            // </ACCESSAUDIT>
            // </REQUEST>
            // return:
            // ------------------------------------------------------------------------------------------

			if (_isObjectContext) ContextUtil.SetComplete();
			
			try 
            {
                xmlResponseElem = xmlOut.CreateElement("RESPONSE");
                xmlOut.AppendChild(xmlResponseElem);
                xmlResponseElem.SetAttribute("TYPE", "SUCCESS");

                xmlIn = XmlAssist.Load(vstrRequest);

                xmlTempResponseNode = ((IAuditBO)this).Validate(xmlIn.DocumentElement, IAuditBOMethod.bomCreateAccessAudit);

                ErrAssistException.CheckXmlResponseNode(xmlTempResponseNode, xmlResponseElem, true);

                // ------------------------------------------------------------------------------------------
                // create the record(s) using the updated xml request string
                // ------------------------------------------------------------------------------------------
                // Using the IAuditBO interface
				xmlTempResponseNode = ((IAuditBO)this).CreateAccessAudit(xmlIn.DocumentElement);
                ErrAssistException.CheckXmlResponseNode(xmlTempResponseNode, xmlResponseElem, true);

                CreateAccessAudit = xmlResponseElem.OuterXml;
				if (_isObjectContext) ContextUtil.SetComplete();
                
                return CreateAccessAudit; 
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex); 
                if (exception.IsWarning())
                {
                    XmlElement tmpxmlResponseElem = xmlResponseElem;
                    exception.AddWarning(tmpxmlResponseElem);
                }
               
                if (exception.IsSystemError())
                {       
                    exception.LogError(exception.Message, System.Diagnostics.EventLogEntryType.Error);
                }

				if (_isObjectContext) ContextUtil.SetAbort();
            }
            return CreateAccessAudit;
        }

        public string GetNumberOfFailedAttempts(string vstrXMLRequest) 
        {
            string GetNumberOfFailedAttempts = String.Empty;
            
            XmlDocument xmlIn = null;
            XmlDocument xmlOut = null;
            XmlElement xmlResponseElem = null;
            XmlNode xmlTempResponseNode = null;

            // header ----------------------------------------------------------------------------------
            // description:  Get a single instance of the persistant data associated with this
            // business object
            // pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
            // return:                       xml Response data stream containing results of operation
            // either: TYPE="SUCCESS"
            // or: TYPE="SYSERR" and <ERROR> element
            // ------------------------------------------------------------------------------------------

            try 
            {
                // Create default response block
                xmlOut = new XmlDocument();
                xmlResponseElem = xmlOut.CreateElement("RESPONSE");
                xmlOut.AppendChild(xmlResponseElem);
                xmlResponseElem.SetAttribute("TYPE", "SUCCESS");
                
                xmlIn = XmlAssist.Load(vstrXMLRequest);

                // Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
				xmlTempResponseNode = ((IAuditBO)this).GetNumberOfFailedAttempts(xmlIn.DocumentElement);
                
                ErrAssistException.CheckXmlResponseNode(xmlTempResponseNode, xmlResponseElem, true);

				XmlAssist.AttachResponseData(xmlTempResponseNode, xmlResponseElem);
                
                GetNumberOfFailedAttempts = xmlResponseElem.OuterXml;
				if (_isObjectContext) ContextUtil.SetComplete();
                
                return GetNumberOfFailedAttempts;
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex);
                if (exception.IsWarning())
                {
                    XmlElement tmpxmlResponseElem = xmlResponseElem;
                    exception.AddWarning(tmpxmlResponseElem);
                }

                if (exception.IsSystemError())
                {
                    exception.LogError(exception.Message , System.Diagnostics.EventLogEntryType.Error);
                }

				if (_isObjectContext) ContextUtil.SetComplete();
            }
            return GetNumberOfFailedAttempts;
        }

        XmlNode IAuditBO.CreateAccessAudit(XmlElement vxmlRequest) 
        {
            XmlNode IAuditBO_CreateAccessAudit = null;

            XmlDocument xmlOut = null;
            XmlElement xmlResponseElem = null;
            XmlNode xmlTempResponseNode = null;
			IAuditTxBO objIAuditTxBO;
            XmlNode xmlTableNode = null;
            XmlElement xmlElem = null;
            string strMachineId = String.Empty;
            string strUserId = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:
            // Create an accessaudit record.
            // For a change to password, a change password record is also created.
            // For an application access or release, an applicationaccess record is also created.
            // pass:
            // vxmlRequest
            // xml Request data stream containing data to be persisted.
            // USERID and MACHINEID from REQUEST element are passed down to DO.
            // Format:
            // <REQUEST USERID= MACHINEID=>
            // <ACCESSAUDIT>
            // <AUDITRECORDTYPE></AUDITRECORDTYPE>
            // <SUCCESSINDICATOR></SUCCESSINDICATOR>
            // <ONBEHALFOFUSERID>Optional</ONBEHALFOFUSERID>
            // <APPLICATIONNUMBER></APPLICATIONNUMBER>
            // <PASSWORDCREATIONDATE>Optional</PASSWORDCREATIONDATE>
            // </ACCESSAUDIT>
            // </REQUEST>
            // return:
            // ------------------------------------------------------------------------------------------

            try 
            {
                xmlOut = new XmlDocument();
                xmlResponseElem = xmlOut.CreateElement("RESPONSE");
                xmlOut.AppendChild(xmlResponseElem);
                xmlResponseElem.SetAttribute("TYPE", "SUCCESS");
                // ------------------------------------------------------------------------------------------
                // Add USERID and MACHINEID from REQUEST element to the body of the request.
                // MACHINEID is optional.
                // ------------------------------------------------------------------------------------------
                // Get the value of attribute MACHINEID in node REQUEST
                
                bool blnOutValue;
                strMachineId = XmlAssist.GetAttributeValue(vxmlRequest, "REQUEST", "MACHINEID", out blnOutValue);

                // Get the value of the attribute USERID withing the REQUEST node
                strUserId = XmlAssist.GetAttributeValue(vxmlRequest, "REQUEST", "USERID", out blnOutValue);

                xmlTableNode = XmlAssist.GetMandatoryNode(vxmlRequest, ".//ACCESSAUDIT");
                 
                xmlElem = vxmlRequest.OwnerDocument.CreateElement("USERID");
                // Make the text of the new element equal to the text of item
                xmlElem.InnerText = strUserId;
                xmlTableNode.AppendChild(xmlElem);
                // If there is a Machine ID then enter this into the MACHINEID node
                if (strMachineId.Length  > 0)
                {
                    xmlElem = vxmlRequest.OwnerDocument.CreateElement("MACHINEID");
                    xmlElem.InnerText = strMachineId;
                    xmlTableNode.AppendChild(xmlElem);
                }
                
                objIAuditTxBO = new AuditTxBO();
                //[DS]- Removed reference to m_objContext. TBD: Instantiation of AuditDO()
                //Type typAuditTxBO = Type.GetTypeFromProgID(App.Title + ".AuditTxBO", true);
                //objIAuditTxBO = (IAuditTxBO)Activator.CreateInstance(typAuditTxBO);
            
                // call Business Transaction Object Create function
                xmlTempResponseNode = objIAuditTxBO.CreateAccessAudit(vxmlRequest);
               
                ErrAssistException.CheckXmlResponseNode(xmlTempResponseNode, xmlResponseElem, true);
                IAuditBO_CreateAccessAudit = xmlResponseElem;
				if (_isObjectContext) ContextUtil.SetComplete();
                
                return IAuditBO_CreateAccessAudit;
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex); 
                if (exception.IsWarning())
                {
                    XmlElement tmpxmlResponseElem = xmlResponseElem;
                    exception.AddWarning(tmpxmlResponseElem);
                }
                
                if (exception.IsSystemError())
                {
                    exception.LogError(exception.Message , System.Diagnostics.EventLogEntryType.Error);
                }
                IAuditBO_CreateAccessAudit = exception.CreateErrorResponseNode();
				if (_isObjectContext) ContextUtil.SetAbort();
            }
            return IAuditBO_CreateAccessAudit;
        }

        XmlNode IAuditBO.GetNumberOfFailedAttempts(XmlElement vxmlRequest) 
        {
            XmlNode IAuditBO_GetNumberOfFailedAttempts = null;
            XmlNode xmlDataNode = null;
            XmlDocument xmlOut = null;
            XmlElement xmlResponseElem = null;
			IAuditDO objIAuditDO;
            int lngNoAttempts = 0;
            // header ----------------------------------------------------------------------------------
            // description:  Get a single instance of the persistant data associated with this
            // business object
            // pass:         vxmlRequest  xml Request data stream containing data to be persisted
            // return:       xml Response Node
            // ------------------------------------------------------------------------------------------
            try 
            {
                xmlOut = new XmlDocument();
                xmlResponseElem = xmlOut.CreateElement("RESPONSE");
                xmlOut.AppendChild(xmlResponseElem);
                xmlResponseElem.SetAttribute("TYPE", "SUCCESS");

				objIAuditDO = new AuditDO();
				//[DS]- Removed reference to m_objContext. TBD: Instantiation of AuditDO()
                //Type typAuditDO = Type.GetTypeFromProgID(App.Title + ".AuditDO", true);
                //objIAuditDO = (IAuditDO)Activator.CreateInstance(typAuditDO);

                lngNoAttempts = objIAuditDO.GetNumberOfFailedAttempts(vxmlRequest);

				xmlDataNode = xmlOut.CreateElement("NUMBEROFATTEMPTS");
                xmlDataNode.InnerText = Convert.ToString(lngNoAttempts);
                xmlResponseElem.AppendChild(xmlDataNode);

                IAuditBO_GetNumberOfFailedAttempts = xmlResponseElem;
				if (_isObjectContext) ContextUtil.SetComplete();
                
                return IAuditBO_GetNumberOfFailedAttempts;
            }
            catch(Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex); 
                if (exception.IsWarning())
                {
                    XmlElement tmpxmlResponseElem = xmlResponseElem;
                    exception.AddWarning(xmlResponseElem);
                }
                if (exception.IsSystemError())
                {
                    exception.LogError(exception.Message , System.Diagnostics.EventLogEntryType.Error);        
                }

				if (_isObjectContext) ContextUtil.SetComplete();
            }
            return IAuditBO_GetNumberOfFailedAttempts;
        }

		protected override void Activate()
		{
			_isObjectContext = ContextUtility.IsObjectContext;
			base.Activate();
		}

		protected override void Deactivate()
		{
			_isObjectContext = false;
		 	 base.Deactivate();
		}
    }
}
