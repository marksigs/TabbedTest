/*
--------------------------------------------------------------------------------------------
Workfile:			IOmigaMessageQueue.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Interface for Omiga support for OMMQ.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		20/06/2007	First .Net version. Ported from OmigaToMessageQueueInterface.idl.
					For simplicity I have re-implemented this interface natively in C# rather
					than reuse the existing COM interface definition.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Messaging;

// Missing XML comment for publicly visible type or member.
#pragma warning disable 1591

namespace Vertex.Fsd.Omiga.VisualBasicPort.Obsolete
{
	[Obsolete("Use OmigaMessageQueueOMMQ instead")]
	public class IOmigaMessageQueue
	{
		[Obsolete("Use OmigaMessageQueueOMMQ.AsyncSend instead")]
		public string AsyncSend(string request, string message)
		{
			OmigaMessageQueueOMMQ omigaMessageQueueOMMQ = new OmigaMessageQueueOMMQ();
			return omigaMessageQueueOMMQ.AsyncSend(request, message);
		}
	}
}
