/*
--------------------------------------------------------------------------------------------
Workfile:			AxWordConstants.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Defines constants used in AxWord.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
AS		15/08/2007	First version. Ported from AxwordConstants.bas.
--------------------------------------------------------------------------------------------
*/
using System;

namespace Vertex.Fsd.Omiga.AxWord
{
	/*
	internal enum AxWordError
	{
		UnspecifiedError = 0x80040000 + 512 + 999,
		// XML related errors
		XmlParserError = 0x80040000 + 512 + 1,
		XmlZeroLengthString = 0x80040000 + 512 + 2,
		// No XML supplied
		XmlMissingMandatoryNode = 0x80040000 + 512 + 3,
		// Word Related Errors
		WinWordProcessError = 0x80040000 + 512 + 10,
		// Word is already running on the system
		WinWordOpenError = 0x80040000 + 512 + 11,
		// Word couldnt open a file
		// OLE Related Errors
		OleCreateEmbedError = 0x80040000 + 512 + 20,
		// Can't create the embedded object, permission error? registry error?
	}
	*/

	public enum DocumentDeliveryType
	{
		Unknown = 0,
		Word = 10,
		Pdf = 20,
		Rtf = 30,
		Xml = 40,
		Tif = 50,
	}

	public enum DocumentCompressionMethod
	{
		Unknown,
		None,
		Compapi,
		Zlib,
	}
}
