using System;
using System.Xml;

namespace Vertex.Fsd.Omiga.omAU
{
    public interface IAuditTxBO
    {
        XmlNode CreateAccessAudit(XmlElement vxmlRequest);
    }
}
