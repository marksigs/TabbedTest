using System;
using System.Xml;
using System.EnterpriseServices;
using System.Runtime.InteropServices;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.omAU
{
    [ComVisible(true)]
    [ProgId("omAU.AuditTxBO")]
    [Guid("E9833F37-786F-45CB-8F1B-F107D42A343D")]
    [Transaction(TransactionOption.Required)]
    public class AuditTxBO : ServicedComponent , IAuditTxBO 
    {

        // Workfile:      AuditTxBO.cls
        // Copyright:     Copyright © 1999 Marlborough Stirling
        // Description:   Transaction Business objects class for omAU
        // 
        // Dependencies:  List any other dependent components
        // e.g. AuditDO
        // Issues:        not part of public interface
        // Instancing:         MultiUse
        // MTSTransactionMode: RequiresTransaction
        // ------------------------------------------------------------------------------------------
        // History:
        // 
        // Prog  Date         Description
        // RF     30/09/99    Created
        // RF     30/11/99    Change MTSTransactionMode from RequiresNewTransaction to RequiresTransaction
        // because of problem in locking applications
        // CL     18/10/00    core00004 Modifications made to conform to coding templates
        // MV     06/03/01    SYS2001: Commenting in CreateAccessAudit
        // ------------------------------------------------------------------------------------------
        
        private const string cstrROOT_NODE_NAME = "ACCESSAUDIT";
        
        XmlNode IAuditTxBO.CreateAccessAudit(XmlElement vxmlRequest)
        {
            XmlNode IAuditTxBO_CreateAccessAudit = null;
            XmlElement xmlElement = null;
            System.Xml.XmlDocument xmlOut = null;
            XmlElement  xmlResponseElem = null;
            IAuditDO objIAuditDO;
            // header ----------------------------------------------------------------------------------
            // description:
            // pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
            // return:       n/a
            // ------------------------------------------------------------------------------------------
            try 
            {
                xmlOut = new System.Xml.XmlDocument();
                xmlResponseElem = xmlOut.CreateElement("RESPONSE");
                xmlOut.AppendChild(xmlResponseElem);
                xmlResponseElem.SetAttribute("TYPE", "SUCCESS");

                // HasValue if the element passed in is the one with the table name else
                // look below the node that is passed in
                if (vxmlRequest.Name == cstrROOT_NODE_NAME)
                {
                    xmlElement = vxmlRequest;
                }
                else
                {   
                    xmlElement = (XmlElement)vxmlRequest.GetElementsByTagName(cstrROOT_NODE_NAME).Item(0);
                }
                if (xmlElement == null)
                {
                   throw new ErrAssistException(OMIGAERROR.MissingPrimaryTag, cstrROOT_NODE_NAME + " tag not found"); 
                }

                objIAuditDO = new AuditDO();

                //Type typAuditDO = Type.GetTypeFromProgID(App.Title + ".AuditDO", true);
                //objIAuditDO = (IAuditDO)Activator.CreateInstance(typAuditDO);
                
                objIAuditDO.CreateAccessAudit(xmlElement);
                IAuditTxBO_CreateAccessAudit = xmlResponseElem;
                ContextUtility.SetComplete();
                return IAuditTxBO_CreateAccessAudit;
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex); 
                if (exception.IsWarning())
                {
                    XmlElement tmpXmlResponseElem = xmlResponseElem ;
                    exception.AddWarning(tmpXmlResponseElem);
                }
                
                if (exception.IsSystemError())
                {
                    exception.LogError(exception.Message, System.Diagnostics.EventLogEntryType.Error);
                }
                
                ContextUtility.SetAbort();
            }
            return IAuditTxBO_CreateAccessAudit;
        }
    }
}
