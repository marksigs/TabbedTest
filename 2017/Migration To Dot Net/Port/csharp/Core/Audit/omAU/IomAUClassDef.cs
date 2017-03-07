using System;
using System.Xml;

namespace Vertex.Fsd.Omiga.omAU
{
    public interface IomAUClassDef
    {
        XmlDocument LoadApplicationAccessData();
        XmlDocument LoadChangePasswordData(); 
        XmlDocument LoadAccessAuditData();
    }
}
