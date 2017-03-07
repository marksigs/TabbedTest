<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<!--
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      ContextDebug.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   A simple debug page to help the analyst's find
				out whats in the context menu
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
APS		03/07/00	Initial creation
JLD		02/02/01	Increased scroll size to 20 for fast forward buttons
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
GHun	22/07/2005	MAR14		Apply ING Style Sheet and GUI Images
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
-->

<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
	<title>Debug: Gui Context</title>
</head>

<body>
	<!-- Specify Screen Layout Here -->
	<form id="frmScreen" mark validate="onchange">
		<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 340px; WIDTH: 410px; POSITION: ABSOLUTE" class="msgGroup">
		
		<div id="spnTable" style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE">
		<table id="tblTable" width="400" border="0" cellspacing="0"cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="50%" class="TableHead">Context Variable Name</td><td width="50%" class="TableHead">Value</td></tr>
			<tr id="row01">	<td class="TableTopLeft">&nbsp;</td><td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row11">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row12">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row13">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row14">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row15">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row16">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row17">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row18">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row19">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row20">	<td class="TableBottomLeft">&nbsp;</td><td class="TableBottomRight">&nbsp;</td></tr>
		</table>
		</div>
		
		<span style="TOP: 310px; LEFT: 4px; POSITION: ABSOLUTE">
			<input id="btnFiveBack" value="<<" type="button" style="WIDTH: 30px; HEIGHT: 25px" class="msgButton">
			<input id="btnBack" value="<" type="button" style="WIDTH: 30px; HEIGHT: 25px" class="msgButton">
			<input id="btnForward" value=">" type="button" style="WIDTH: 30px; HEIGHT: 25px" class="msgButton">
			<input id="btnFiveForward" value=">>" type="button" style="WIDTH: 30px; HEIGHT: 25px" class="msgButton">
		</span>

		</div>
	</form>

	<!-- Main Buttons -->
	<div style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 355px; WIDTH: 410px"> 
	<!-- #include FILE="msgButtons.asp" --> 
	</div> 

	<!-- Specify Code Here -->
	<script language="JScript" type="text/javascript">
	<!--
		// JScript Code
		var m_iTableLength = 20;
		var m_iTableIndex = 0;
		var frmContext;
		var scScreenFunctions;
		var m_BaseNonPopupWindow = null;
		
		/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			Function:		window.onload()
			Description:	One-time initialization code is usually performed 
							in response to the window.onload procedure. 
							Waiting for this event to fire ensures that the 
							page is completely loaded.
			Args Passed:	N/a
			Returns:		N/a
		~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
		function window.onload()
		{
			RetrieveData();
			
			// Make the required buttons available on the bottom of the screen
			// (see msgButtons.asp for details)
			var sButtonList = new Array("Submit");
					
			ShowMainButtons(sButtonList);
			
			DisplayTable(0);

			Validation_Init();

			scScreenFunctions.SetFocusToFirstField(frmScreen);
			
			window.returnValue	= null;
		}
	
		function RetrieveData()
		{	
			var sArguments = window.dialogArguments;
			window.dialogTop	= sArguments[0];
			window.dialogLeft	= sArguments[1];
			window.dialogWidth	= sArguments[2];
			window.dialogHeight = sArguments[3];
			var sParameters		= sArguments[4];
			m_BaseNonPopupWindow = sArguments[5];
			scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
			frmContext = sParameters[0];
		}

		function DisplayTable(iTableIndex)
		{
			var sName, sValue;
			var nRowIndex = 1;
			//var frmContext = thisWindow.parent.frames("omigamenu").document.forms("frmContext");
			
			iBegin = iTableIndex;						
			for (var j = iBegin; j < (iBegin + m_iTableLength); j++)
			{
				sName = "";
				sValue = "";
				if (j < frmContext.length)
				{
					sName = frmContext(j).name;
					sValue = frmContext(j).value;										
				}
				if (sValue == "") sValue = " ";
				if (sName == "") sName = " ";	
				scScreenFunctions.SizeTextToField(tblTable.rows(nRowIndex).cells(0),sName);	
				scScreenFunctions.SizeTextToField(tblTable.rows(nRowIndex).cells(1),sValue);
				nRowIndex++;
			}	
		}
		
		function frmScreen.btnBack.onclick()
		{
			if (m_iTableIndex != 0)
			{
				m_iTableIndex--;
			}
			DisplayTable(m_iTableIndex);
		}
		
		function frmScreen.btnFiveBack.onclick()
		{
			m_iTableIndex -= 20;
			if (m_iTableIndex < 0) m_iTableIndex = 0;
			DisplayTable(m_iTableIndex);
		}
		
		function frmScreen.btnForward.onclick()
		{
			m_iTableIndex++;
			DisplayTable(m_iTableIndex);
		}
		function frmScreen.btnFiveForward.onclick()
		{
			m_iTableIndex += 20;			
			DisplayTable(m_iTableIndex);
		}
		
		function btnSubmit.onclick()
		{
			// go back to the calling page
			window.returnValue	= null;
			window.close();
		}
	-->
	</script>
	<script src="validation.js" type="text/javascript"></script>
</body>
</html>
