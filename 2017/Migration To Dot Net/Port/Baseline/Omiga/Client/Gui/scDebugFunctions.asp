<%@  Language=JScript %>
<%/*
Workfile:      scDebugFunctions.asp
Copyright:     Copyright © 2003 Marlborough Stirling

Description:   Helper functions for client side debug logging.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AS		17/02/03	First version
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<OBJECT ID="ClientDebugLog" STYLE="DISPLAY: none; VISIBILITY: hidden" CLASSID="CLSID:94772CA5-20F8-478E-8061-F17143075B0E" CODEBASE="omDebug.cab#version=3,4,3,0" VIEWASTEXT tabindex="-1"></OBJECT>
<HTML ID=DebugFunctions>
<script language="JavaScript">

// Comment out the following line for scDebugFunctions.asp.
// var log = new DebugLog("C:\\Omiga4Log\\Omiga4.log");
// Uncomment the following line for scDebugFunctions.asp.
var public_description = new DebugLog("C:\\Omiga4Log\\Omiga4.log");

function DebugLog(strFileName)
{
	try
	{
		this.m_objDebugLog	= new ActiveXObject("omDebug.ClientDebugLog");
		this.WriteLine		= DebugLogWriteLine;
		this.CreateGuid		= DebugLogCreateGuid;
		if (this.m_objDebugLog != null)
		{
			this.m_objDebugLog.FileName = <% Response.Write("\"" + Request.QueryString("FileName") + "\""); %>;
		}
	}
	catch(e)
	{
	}
}

function DebugLogWriteLine(strGuid, nServerElapsedTimeMS, nIndent, strType, strFunction, strXML)
{
	var bSuccess = false;
	
	if (this.m_objDebugLog != null)
	{
		bSuccess = this.m_objDebugLog.WriteFunctionLine(strGuid, nServerElapsedTimeMS, nIndent, strType, strFunction, strXML);
	}

	return bSuccess;
}

function DebugLogCreateGuid()
{
	var strGuid = "";

	if (this.m_objDebugLog != null)
	{
		strGuid = this.m_objDebugLog.CreateGuid();
	}

	return strGuid;
}


</script>
