<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      cr060.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Areas of interest screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		04/02/00	Optimised version
AY		07/02/00	Forgot to put the new validation script in!
AY		29/03/00	scScreenFunctions change
IW		15/05/00	SYS0201 - Customer Name Populated
BG		17/05/00	SYS0752 Removed Tooltips
LD		23/05/02	SYS4727 Use cached versions of frame functions

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Areas of Interest</title>
</head>

<body>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" style="VISIBILITY: hidden">
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 320px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Customer Name
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<input id="txtCustomerName" maxlength="40" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>

	<span style="TOP: 40px; LEFT: 140px; POSITION: ABSOLUTE">
		<img id="MortPic" src="images/sml_mortg1.gif" style="TOP: 0px; LEFT: 0px; HEIGHT: 50px; WIDTH: 250px; POSITION: ABSOLUTE" WIDTH="254" HEIGHT="52">
		<img id="MortTick" src="images/chk_tick.gif" style="TOP: 12px; LEFT: 300px; HEIGHT: 29px; WIDTH: 29px; POSITION: ABSOLUTE" WIDTH="29" HEIGHT="29">
	</span>

	<span style="TOP: 100px; LEFT: 140px; POSITION: ABSOLUTE">
		<img id="LifePic" src="images/sml_life1.gif" style="TOP: 0px; LEFT: 0px; HEIGHT: 50px; WIDTH: 250px; POSITION: ABSOLUTE" WIDTH="254" HEIGHT="52">
		<img id="LifeTick" src="images/chk_tick.gif" style="TOP: 12px; LEFT: 300px; HEIGHT: 29px; WIDTH: 29px; VISIBILITY: VISIBLE; POSITION: ABSOLUTE" onclick="OnChangeTickStatus(LifeTick, LifeCross)" WIDTH="29" HEIGHT="29">
		<img id="LifeCross" src="images/chk_cross.gif" style="TOP: 12px; LEFT: 300px; HEIGHT: 29px; WIDTH: 29px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE" onclick="OnChangeTickStatus(LifeCross, LifeTick)" WIDTH="29" HEIGHT="29">
	</span>

	<span style="TOP: 160px; LEFT: 140px; POSITION: ABSOLUTE">
		<img id="PensionPic" src="images/sml_pen1.gif" style="TOP: 0px; LEFT: 0px; HEIGHT: 50px; WIDTH: 250px; POSITION: ABSOLUTE" WIDTH="254" HEIGHT="52">
		<img id="PensionTick" src="images/chk_tick.gif" style="TOP: 12px; LEFT: 300px; HEIGHT: 29px; WIDTH: 29px; VISIBILITY: VISIBLE; POSITION: ABSOLUTE" onclick="OnChangeTickStatus(PensionTick, PensionCross)" WIDTH="29" HEIGHT="29">
		<img id="PensionCross" src="images/chk_cross.gif" style="TOP: 12px; LEFT: 300px; HEIGHT: 29px; WIDTH: 29px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE" onclick="OnChangeTickStatus(PensionCross, PensionTick)" WIDTH="29" HEIGHT="29">
	</span>

	<span style="TOP: 220px; LEFT: 140px; POSITION: ABSOLUTE">
		<img id="InvestmentPic" src="images/sml_inv1.gif" style="TOP: 0px; LEFT: 0px; HEIGHT: 50px; WIDTH: 250px; POSITION: ABSOLUTE" WIDTH="254" HEIGHT="52">
		<img id="InvestmentTick" src="images/chk_tick.gif" style="TOP: 12px; LEFT: 300px; HEIGHT: 29px; WIDTH: 29px; VISIBILITY: VISIBLE; POSITION: ABSOLUTE" onclick="OnChangeTickStatus(InvestmentTick, InvestmentCross)" WIDTH="29" HEIGHT="29">
		<img id="InvestmentCross" src="images/chk_cross.gif" style="TOP: 12px; LEFT: 300px; HEIGHT: 29px; WIDTH: 29px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE" onclick="OnChangeTickStatus(InvestmentCross, InvestmentTick)" WIDTH="29" HEIGHT="29">
	</span>

	<span style="TOP: 290px; LEFT: 10px; POSITION: ABSOLUTE">
		<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
		<span style="TOP: 0px; LEFT: 64px; POSITION: ABSOLUTE">
			<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
	</span>
</div>
</form>

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var m_sCustomerNumber = null;
var m_sCustomerVersionNumber = null;
var m_sCustomerReadOnly = null;
var scScreenFunctions;
var m_sCustomerName = null
var m_BaseNonPopupWindow = null;

function window.onload()
{
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];
	
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions.ShowCollection(frmScreen);

<%	//APS UNIT TEST REF 11 & 12 - You cannot read the context from a popup window
%>	GetAreasOfInterest(sArgArray[0]);

<%	//Passing through the CustomerNumber and CustomerVersionNumber
	// APS UNIT TEST REF 72 - Read Only processing
%>	m_sCustomerNumber = sArgArray[1];
	m_sCustomerVersionNumber = sArgArray[2];
	m_sCustomerReadOnly = sArgArray[3];
	m_sCustomerName = sArgArray[4];
	
	//set the Customer Name
	frmScreen.txtCustomerName.value = m_sCustomerName;
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCustomerName");

	if (m_sCustomerReadOnly == "1") frmScreen.btnOK.disabled = true;

<%	// AY 15/09/99 - return null if OK is not pressed
%>	window.returnValue = null;
}

function OnChangeTickStatus(ImgToHide, ImgToShow)
{
<%	// AY 15/09/99 - this function to be run when the images are clicked
	// flag a change
%>	FlagChange(true);
	ChangeTickStatus(ImgToHide, ImgToShow);
}

function ChangeTickStatus(ImgToHide, ImgToShow)
{
	ImgToHide.style.visibility="hidden";
	ImgToShow.style.visibility="visible";
}

function frmScreen.btnOK.onclick()
{
<%	// AY 15/09/99 return a 2D array
%>	var sReturn = new Array();

	sReturn[0] = IsChanged();
	sReturn[1] = WriteAreasOfInterest();

	window.returnValue = sReturn;
	window.close();
}

function frmScreen.btnCancel.onclick()
{
	window.close();
}

function WriteAreasOfInterest()
{
<%	// <AREASOFINTERESTLIST>
			// <AREASOFINTEREST>
			// </AREASOFINTEREST>
	// </AREASOFINTERESTLIST>
%>	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTicks = new Array("LifeTick", "PensionTick", "InvestmentTick");
	var TagAREASOFINTERESTLIST = XML.CreateActiveTag("AREASOFINTERESTLIST");

	for(var nLoop = 0;nLoop < 3;nLoop++)
		if(frmScreen.all(sTicks[nLoop]).style.visibility == "visible")
		{
			XML.ActiveTag = TagAREASOFINTERESTLIST;
			XML.CreateActiveTag("AREASOFINTEREST");
			XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
			XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
			XML.CreateTag("INTERESTAREA", nLoop+1);
		}

	var sReturn = XML.XMLDocument.xml;
	XML = null;

	return sReturn;
}

function GetAreasOfInterest(sXML)
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.LoadXML(sXML);
	XML.CreateTagList("AREASOFINTERESTLIST");

<%	// APS UNIT TEST REF 12 - Moved the initialise of the buttons out to here
	// from the if statement because if the user has no areas of interest the
	// XML is effectively NULL and this means the buttons are defaulted to a tick
%>	ChangeTickStatus(frmScreen.LifeTick,frmScreen.LifeCross);
	ChangeTickStatus(frmScreen.PensionTick,frmScreen.PensionCross);
	ChangeTickStatus(frmScreen.InvestmentTick,frmScreen.InvestmentCross);

	if(XML.ActiveTagList.length > 0)
	{
		XML.CreateTagList("AREASOFINTEREST");

		for(var nLoop = 0;nLoop < XML.ActiveTagList.length;nLoop++)
		{
			XML.SelectTagListItem(nLoop);

			switch(XML.GetTagText("INTERESTAREA"))
			{
				case "1":
					ChangeTickStatus(frmScreen.LifeCross,frmScreen.LifeTick);
				break;
				case "2":
					ChangeTickStatus(frmScreen.PensionCross,frmScreen.PensionTick);
				break;
				case "3":
					ChangeTickStatus(frmScreen.InvestmentCross,frmScreen.InvestmentTick);
				break;
			}
		}
	}

	XML = null;
}
-->
</script>
</body>
</html>
