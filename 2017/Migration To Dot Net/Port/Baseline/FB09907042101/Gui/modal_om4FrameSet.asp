<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<%@ LANGUAGE="JSCRIPT" %>
<% 
	/* PSC 16/10/2005 MAR57 - Start */
	var sStartScreen = Request.QueryString("Screen") + '';
	if (sStartScreen == '' || sStartScreen == 'undefined')
		sStartScreen = "MN010.asp";
	/* PSC 16/10/2005 MAR57 - End */
%>
<html>
<%/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMids History:

Prog	Date		Description
HMA     24/11/04    BMIDS948  Changed rows for navmenu.
GHun	23/12/2004	BMIDS960 Added noresize attribute to frames
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS History:

Prog	Date		Description
GHun	24/08/2005	MAR52 Added src="blank.htm" to all frames with no source
PSC		12/10/2005	MAR57 Increase size of frame for omigamenu
IK		05/06/2006	MAR1849 cache combos & global params for browser session
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
	<title><!--#include file="customise/modal_om4framesetcustomise.asp"--></title>
<script language="JScript" type="text/javascript">
<!--

	<% /* IK_05/06/2006_MAR1849 combo cache */ %>
	var xCombos  = new ActiveXObject("microsoft.xmldom");
	xCombos.async = false;
	xCombos.setProperty("SelectionLanguage","XPath");
	xCombos.loadXML("<RESPONSE><LIST/></RESPONSE>");

	<% /* IK_05/06/2006_MAR1849 globals cache */ %>
	var xGlobals  = new ActiveXObject("microsoft.xmldom");
	xGlobals.async = false;
	xGlobals.setProperty("SelectionLanguage","XPath");
	xGlobals.loadXML("<GLOBALS/>");


<%/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Function:		CheckForLocks()
	Description:	Called before this page is unloaded.
					It allows the closing down of the framework to
					be halted if there are locks outstanding and it
					is not the intention to exit.
	Args Passed:	N/a
	Returns:		N/a
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/%>
function CheckForLocks()
{
	<%// If the case is read only there will not be locks anyway
	// Otherwise cause a message to be displayed if there is
	// an application lock or customer lock
	
	/* This is used to fudge a fix because a msgbox cancel seems to
	   raise an onbeforeunloadevent needlessly SYS0735.
	   Any not-null value will disable the event.  */%>
	if(omigamenu.frmContext.idReadOnly.value != "1")
	{
		if(omigamenu.frmContext.idApplicationNumber.value != "")
		{
			<% /* SYS3311 Set context variable correctly. We do not need a 
			message to appear. We just need the application to unlock. 
			event.returnValue = "Exiting will unlock the application";
			*/ %>
			omigamenu.frmContext.IE5BugDisableUnloadEvent.value = "";
		}
		else
		{
			if(omigamenu.frmContext.idCustomerNumber.value != "")
			{
			<% /* SYS3311 Set context variable correctly. We do not need a 
			message to appear. We just need the customer to unlock. 
			event.returnValue = "Exiting will unlock the customer";
			*/ %>
				omigamenu.frmContext.IE5BugDisableUnloadEvent.value = "";
			}
		}
	}
}
-->
</script>
</head>
<% /* These frames are pixel-perfect on size. Change them at your peril!!!! 
	  BMIDS948  Change rows from 275 to 289                                */ %>
<frameset id="idFrames" cols="151,*" rows="648" frameborder="no" border="0" framespacing="0" onbeforeunload="CheckForLocks()"> 
	<frameset rows="335,*" cols="*"> 
		<frame id="omigamenu" frameborder="no" noresize name="omigamenu" marginwidth="0" scrolling="no" src="blank.htm">
		<% /* PSC 16/10/2005 MAR57 */ %>
		<frame id="navMenu" src="fw020.asp?Screen=<%=sStartScreen%>" frameborder="no" noresize name="navigation">
	</frameset>
	<frame id="mainFrame" name="main" frameborder="no" scrolling="yes" noresize src="blank.htm">

	<noframes>
		<body bgcolor="#FFFFFF">&nbsp;</body>
	</noframes>
</frameset>

</html>
