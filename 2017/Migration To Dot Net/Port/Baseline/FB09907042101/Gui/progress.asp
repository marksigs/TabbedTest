<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%@ LANGUAGE="JSCRIPT" %>

<html>
<% /*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:			progress.asp
Copyright:			Copyright © 2006 Marlborough Stirling
Description:		progress dialog
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
BCan	17/08/2004	Created
PJO     07/03/2006  MAR1359 Put into a working state for release
GHun	20/03/2006	MAR1453
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4" />
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
	<title>Progress...</title>
	<link rel="stylesheet" type="text/css" href="stylesheet.css" />
</head>
<body onKeyPress="checkKeyPress()">
	<div id="divStatus" style="TOP: 10px; LEFT: 10px; WIDTH: 380px; HEIGHT: 80px; POSITION: absolute;" class="msgGroup">
		<table id="tblStatus" border="0" cellspacing="0" cellpadding="0" width="380">
			<tr>
				<td align="center"><img id="progressImg" alt="progress image" src="images/ing_rotate.gif" border="0" width="90" height="90"></td>
			</tr>
    		<tr align="center">
				<td id="colStatus" align="center" class="msgLabelWait"></td>
			</tr>
		</table>
	</div>
	
<script type="text/javascript">
function window.onload ()
{
	if (location.search.length > 0)
	{
		colStatus.innerText = unescape (location.search.substring(1));
	}
}

function checkKeyPress()
{
	if (window.event.keyCode == 27)
        window.close();
}
</script>
</body>
</html>
