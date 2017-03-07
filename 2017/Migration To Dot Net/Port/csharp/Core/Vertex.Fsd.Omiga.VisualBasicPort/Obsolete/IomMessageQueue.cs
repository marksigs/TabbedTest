/*
--------------------------------------------------------------------------------------------
Workfile:			IomMessageQueue.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Interface for Omiga message queueing.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		20/06/2007	First .Net version. Ported from IomMessageQueue.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Messaging;

// Missing XML comment for publicly visible type or member.
#pragma warning disable 1591

namespace Vertex.Fsd.Omiga.VisualBasicPort.Obsolete
{
	[Obsolete("Use OmMessageQueue instead")]
    public class IomMessageQueue
    {
		[Obsolete("Use OmMessageQueue.SendToQueue instead")]
		public XmlNode SendToQueue(XmlNode xmlRequest)
		{
			return OmMessageQueue.SendToQueue(xmlRequest);
		}
    }
}
