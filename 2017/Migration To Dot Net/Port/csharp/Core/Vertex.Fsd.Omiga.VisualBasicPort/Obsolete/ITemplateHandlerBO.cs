/*
--------------------------------------------------------------------------------------------
Workfile:			ITemplateHandlerBO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Interface for template handling business objects.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		20/06/2007	First .Net version. Ported from ITemplateHandlerBO.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Printing;

// Missing XML comment for publicly visible type or member.
#pragma warning disable 1591

namespace Vertex.Fsd.Omiga.VisualBasicPort.Obsolete
{
	[Obsolete("Use TemplateHandlerBO instead")]
    public class ITemplateHandlerBO
    {
		private TemplateHandlerBO _templateHandlerBO = new TemplateHandlerBO();

		[Obsolete("Use TemplateHandlerBO.GetTemplate instead")]
		public XmlElement GetTemplate(XmlElement xmlRequest)
		{
			return _templateHandlerBO.GetTemplate(xmlRequest);
		}

		[Obsolete("Use TemplateHandlerBO.FindAvailableTemplates instead")]
		public XmlNode FindAvailableTemplates(XmlElement xmlRequest)
		{
			return _templateHandlerBO.FindAvailableTemplates(xmlRequest);
		}
    }
}
