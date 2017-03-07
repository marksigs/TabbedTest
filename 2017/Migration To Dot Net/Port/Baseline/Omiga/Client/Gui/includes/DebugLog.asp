<%
/*
Workfile:      DebugLog.asp
Copyright:     Copyright © 2003 Marlborough Stirling

Description:   Wrapper functions for client side debug log object.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AS		18/02/03	First version
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<script language="JavaScript">
function DebugLogWrite(strLabel, nIndent, nServerElapsedTimeMS, strText)
{
	try
	{
		if (window.dialogArguments != null)
		{
			<% /* Popup window. */ %>
			window.dialogArguments[5].top.frames[1].document.all.scDebugFunctions.Write(strLabel, nIndent, nServerElapsedTimeMS, strText);
		}
		else
		{
			top.frames[1].document.all.scDebugFunctions.Write(strLabel, nIndent, nServerElapsedTimeMS, strText);
		}
	}
	catch(e)
	{
		<% /* Uncomment to see errors. */ %>
		<% /* alert("Exception in DebugLog.asp:DebugLogWrite(): " + e.description + " (" + e.number + ")"); */ %>
	}
}
function DebugLogWriteLine(strLabel, nIndent, nServerElapsedTimeMS, strText)
{
	DebugLogWrite(strLabel, nIndent, nServerElapsedTimeMS, strText + "\r\n");
}
</script>
