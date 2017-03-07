using System;
using System.Xml;

namespace Vertex.Fsd.Omiga.omAU
{
    public interface IAuditNTxBO
    {
        XmlNode CreateAccessAudit(XmlElement vxmlRequest);       
    }
}
