using System;
using System.Xml;
using System.EnterpriseServices;
using System.Runtime.InteropServices;
using Vertex.Fsd.Omiga.VisualBasicPort;

namespace Vertex.Fsd.Omiga.omAU
{
    [ComVisible(true)]
    [ProgId("omAU.AuditNTxBO")]
    [Guid("53CFC442-B0DE-42C8-99BB-C2EBDEAEB40E")]
    [Transaction(TransactionOption.RequiresNew)]
    public class AuditNTxBO : ServicedComponent, IAuditNTxBO
    {

        // Workfile:      AuditNTxBO.cls
        // Copyright:     Copyright © 2005 Marlborough Stirling
        // Description:   New Transaction Business objects class for omAU
        // 
        // Dependencies:  List any other dependent components
        // e.g. AuditDO
        // Issues:        not part of public interface
        // Instancing:         MultiUse
        // MTSTransactionMode: RequiresNewTransaction
        // ------------------------------------------------------------------------------------------
        // History:
        // 
        // Prog  Date         Description
        // AS     11/04/05    Created
        // ------------------------------------------------------------------------------------------
        
        XmlNode IAuditNTxBO.CreateAccessAudit(XmlElement vxmlRequest) 
        {
			return ((IAuditTxBO)new AuditTxBO()).CreateAccessAudit(vxmlRequest);
        }
    }
}
