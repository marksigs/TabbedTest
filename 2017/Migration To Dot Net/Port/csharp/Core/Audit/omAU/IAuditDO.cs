using System;
using System.Xml;

namespace Vertex.Fsd.Omiga.omAU
{
    public interface IAuditDO
    {
        // Workfile:      AuditDO.cls
        // Copyright:     Copyright © 1999 Marlborough Stirling
        // Description:   Data objects class for omAU
        // 
        // Dependencies:  ADOAssist
        // Add any other dependent components
        // 
        // Issues:        Instancing:         MultiUse
        // MTSTransactionMode: UsesTransaction
        // ------------------------------------------------------------------------------------------
        // History:
        // 
        // Prog  Date         Description
        // MV     06/03/01    SYS2001: changed CreateAccessAudit Function to CreateAccessaudit Procedure (SUB)
        // ------------------------------------------------------------------------------------------

        void CreateAccessAudit(XmlElement vxmlTableElement);
        void AddDerivedData(XmlNode vxmlData);
        int GetNumberOfFailedAttempts(XmlElement vxmlTableElement);
        bool IsPasswordChange(string vstrAuditRecType);
        bool IsApplicationAccess(string vstrAuditRecType);
        bool IsApplicationRelease(string vstrAuditRecType);
        bool IsLogon(string vstrAuditRecType);
        string GetApplicationLockValueId();
        string GetApplicationReleaseValueId();
        string GetChangePasswordValueId();
        string GetLogonValueId();        
    }
}
