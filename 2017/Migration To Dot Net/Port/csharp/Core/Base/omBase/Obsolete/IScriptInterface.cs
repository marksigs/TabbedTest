/*
--------------------------------------------------------------------------------------------
Workfile:			IScriptInterface.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		
--------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		27/07/2007	First .Net version. Ported from IScriptInterface.cls.
--------------------------------------------------------------------------------------------
*/
using System;

namespace Vertex.Fsd.Omiga.omBase
{
	[Obsolete("Use ScriptInterface instead")]
    public class IScriptInterface
    {
		[Obsolete("Use ScriptInterface.RunScript instead")]
        public string RunScript(string componentName, string objectName, string functionName, string XmlRequest) 
        {
			return ScriptInterface.RunScript(componentName, objectName, functionName, XmlRequest);
        }
    }
}
