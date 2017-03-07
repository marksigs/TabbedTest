/*
--------------------------------------------------------------------------------------------
Workfile:			ScriptInterface.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
RF		20/10/99	RunScript now returns a string
--------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		27/07/2007	First .Net version. Ported from ScriptInterface.cls.
--------------------------------------------------------------------------------------------
*/
using System;

namespace Vertex.Fsd.Omiga.omBase
{
	public static class ScriptInterface
	{
		public static string RunScript(string componentName, string objectName, string functionName, string XmlRequest) 
		{
			return "<RESPONSE TYPE=\"SUCCESS\"></RESPONSE>";
		}
	}
}
