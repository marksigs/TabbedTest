using System;
using System.Xml;

namespace Vertex.Fsd.Omiga.omAU
{
    public enum IAuditBOMethod 
	{
        bomCreateAccessAudit,
    }

    public interface IAuditBO
    {
        XmlNode CreateAccessAudit(XmlElement vxmlRequest);
        XmlNode Validate(XmlElement vxmlRequest, IAuditBOMethod vbomMethod);
		XmlNode GetNumberOfFailedAttempts(XmlElement vxmlRequest);
	}
}
