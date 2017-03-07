/*
--------------------------------------------------------------------------------------------
Workfile:			ITemplateHandlerDO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Interface for template handling data objects.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		20/06/2007	First .Net version. Ported from ITemplateHandlerDO.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Printing;

// Missing XML comment for publicly visible type or member.
#pragma warning disable 1591

namespace Vertex.Fsd.Omiga.VisualBasicPort.Obsolete
{
	[Obsolete("Use TemplateHandlerDO instead")]
    public class ITemplateHandlerDO
    {
		private TemplateHandlerDO _templateHandlerDO = new TemplateHandlerDO();

		[Obsolete("Use TemplateHandlerDO.GetTemplate instead")]
        public XmlNode GetTemplate(XmlElement xmlTableElement)
		{
			return _templateHandlerDO.GetTemplate(xmlTableElement);
		}

		[Obsolete("Use TemplateHandlerDO.FindAvailableTemplates instead")]
		public XmlNode FindAvailableTemplates(XmlElement xmlTableElement)
		{
			return _templateHandlerDO.FindAvailableTemplates(xmlTableElement);
		}
    }
}
