/*
--------------------------------------------------------------------------------------------
Workfile:			PrintManagerConstants.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		omPM constants.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
AS		04/04/2006	First version.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		09/05/2007	First .Net version. Ported from PrintManagerConstants.bas.
--------------------------------------------------------------------------------------------
*/
using System;

// Missing XML comment for publicly visible type or member.
#pragma warning disable 1591

namespace Vertex.Fsd.Omiga.VisualBasicPort.Printing
{
	/// <summary>
	/// Constants used by the omPM component.
	/// </summary>
	public class PrintManagerConstants
	{
		public const int DOCUMENTMANAGEMENTSYSTEMTYPE_DMS = 1;
		public const int DOCUMENTMANAGEMENTSYSTEMTYPE_FILENET = 2;
		public const int DOCUMENTMANAGEMENTSYSTEMTYPE_GEMINI = 3;
		public const int EVENTKEY_CREATED = 0;
		public const int EVENTKEY_EDITED = 1;
		public const int EVENTKEY_VIEWED = 2;
		public const int EVENTKEY_REPRINTED = 3;
		public const int EVENTKEY_FULFILLMENTSEND = 4;
		public const int EVENTKEY_FULFILLMENTRESEND = 5;
		public const int EVENTKEY_FULFILLMENTCANCEL = 6;
		public const int EVENTKEY_SMS = 7;
		public const int EVENTKEY_EMAIL = 8;
		public const int EVENTKEY_RECATEGORISATION = 9;
		public const int EVENTKEY_NOTAPPROVED = 10;
		public const int EVENTKEY_APPROVED = 11;
		public const int EVENTKEY_GEMINIPRINTED = 12;
		public const int GEMINIPRINTSTATUS_AWAITINGAPPROVAL = 10;
		public const int GEMINIPRINTSTATUS_NOTAPPROVED = 20;
		public const int GEMINIPRINTSTATUS_APPROVED = 30;
		public const int GEMINIPRINTSTATUS_GEMINIPRINTED = 40;
	}
}
