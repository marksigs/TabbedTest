using System;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.omAU
{
    public class omAUClassDef : IomAUClassDef
    {
        // Workfile:      ClassDef.cls
        // Copyright:     Copyright © 1999 Marlborough Stirling
        // 
        // Description:   Code template for omiga4 Class Definition
        // 
        // ------------------------------------------------------------------------------------------
        // History:
        // 
        // Prog   Date        Description
        // PSC    14/12/99    Created
        // PSC    28/02/00    Remove Dim As New
        // ------------------------------------------------------------------------------------------
        
		XmlDocument IomAUClassDef.LoadAccessAuditData() 
        {
            const string strXml = "<TABLENAME>ACCESSAUDIT" + 
                "<PRIMARYKEY>ACCESSAUDITGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" + 
                "<OTHERS>USERID<TYPE>dbdtString</TYPE></OTHERS>" + 
                "<OTHERS>ACCESSDATETIME<TYPE>dbdtDateTime</TYPE></OTHERS>" + 
                "<OTHERS>ATTEMPTNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" + 
                "<OTHERS>AUDITRECORDTYPE<TYPE>dbdtComboId</TYPE>" + 
                "<COMBO>AccessAuditType</COMBO></OTHERS>" + 
                "<OTHERS>MACHINEID<TYPE>dbdtString</TYPE></OTHERS>" + 
                "<OTHERS>SUCCESSINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>" + 
                "</TABLENAME>";
            return XmlAssist.Load(strXml);
        }

        XmlDocument IomAUClassDef.LoadApplicationAccessData() 
        {
           const string strXml = "<TABLENAME>APPLICATIONACCESS" + 
                "<PRIMARYKEY>ACCESSAUDITGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" + 
                "<OTHERS>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" + 
                "</TABLENAME>";
           return XmlAssist.Load(strXml);
        }

        XmlDocument IomAUClassDef.LoadChangePasswordData() 
        {
            const string strXml = "<TABLENAME>CHANGEPASSWORD" + 
                "<PRIMARYKEY>ACCESSAUDITGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" + 
                "<OTHERS>ONBEHALFOFUSERID<TYPE>dbdtString</TYPE></OTHERS>" + 
                "</TABLENAME>";
            return XmlAssist.Load(strXml);
        }
    }
}
