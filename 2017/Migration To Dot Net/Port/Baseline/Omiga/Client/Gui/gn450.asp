<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      gn450.asp
Copyright:     Copyright © 2001 Marlborough Stirling

Description:   Popup
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
CL		22/01/01	SYS1756 Screen first created
CL		09/01/01	SYS1936 Change for 'Other' fields
CL		05/04/01	SYS2239
DPF		20/06/02	Changes to file to bring in line with Core V7.0.2, changes are...
					SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		13/05/2002	BMIDS00004	Modified PopulateScreen()
MV		15/05/2002	BMIDS00004	Modified PopulateScreen()
GHun	22/05/2002	BMIDS00005	Added view option, and changed editting to also edit other fields 
DPF		03/07/2002	BMIDS00151	Added line to PopulateScreen to ensure Other Reason text box is correctly
								disabled / enabled.
GHun	07/08/2002	BMIDS00303	Other Reason text was cleared and enabled incorrectly
ASu		13/09/2002	BMIDS00153	Validation and error handling on overflow from the textarea added.
ASu		26/09/2002	BMIDS00153	Length restriction added to 'Other Reason' field.
MC		19/04/2004	CC057		Blank spaces padded to the title text.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Contact History Detail  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* removed as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077 
<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
*/ %>
<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: visible" year4 validate="onchange" mark>
<div id="divBackground" style="HEIGHT: 300px; LEFT: 4px; POSITION: absolute; TOP: 8px; WIDTH: 620px" class="msgGroup">
		
	<span style="LEFT: 5px; POSITION: absolute; TOP: 10px; WIDTH: 80px" class="msgLabel">
		Customer Name
		<span style="LEFT: 94px; POSITION: absolute; TOP: -3px">
			<input id="txtCustomer" style="WIDTH: 350px" class="msgTxt">
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 45px" class="msgLabel">
		Contact Reason
	</span>	
	<span style="LEFT: 100px; POSITION: absolute; TOP: 45px" class="msgLabel">	
		<select id="cboReason" style="WIDTH: 250px" class="msgCombo"></select>
	</span>
	
	<span style="LEFT: 364px; POSITION: absolute; TOP: 45px" class="msgLabel" >
		User Name
	</span>
	
	<span style="LEFT: 434px; POSITION: absolute; TOP: 45px">
		<input id="txtUserName" style="WIDTH: 170px" class="msgTxt">
	</span>
		
	<span id="spnOtherReason" style="LEFT: 4px; POSITION: absolute; TOP: 80px" class="msgLabel">
		Other Reason
	</span>	
	
	<span  style="LEFT: 100px; POSITION: absolute; TOP: 75px">
		<input id="txtOtherReason" style="WIDTH: 250px" maxlength="30" class="msgTxt">
	</span>
		
	<span style="LEFT: 363px; POSITION: absolute; TOP: 80px" class="msgLabel" >
		Unit Name
		<span style="LEFT: 70px; POSITION: absolute; TOP: -3px">
			<input id="txtUnitName" style="WIDTH: 170px" class="msgTxt">
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 110px" class="msgLabel">
		Details
		<span style="LEFT: 360px; POSITION: absolute; TOP: 3px" class="msgLabel" >
			Date/Time
		</span>
		<span style="LEFT: 430px; POSITION: absolute; TOP: -3px">
			<input id="txtDateTime" style="WIDTH: 170px" class="msgTxt">
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 30px"><TEXTAREA class=msgTxt id=txtDetails rows=6 style="WIDTH: 600px"></TEXTAREA>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 130px; WIDTH: 150px">
			<input id="chkCheckBox" type="checkbox" value="1">
			<label for="chkCheckBox" class="msgLabel">Delete this record</label>
		</span>
	</span>
	<span id="spnButtons" style="LEFT: 4px; POSITION: absolute; TOP: 270px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
			<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="LEFT: 64px; POSITION: absolute; TOP: 0px">
			<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
	</span>
	
</div>
</form><!-- #include FILE="attribs/gn450attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sReadOnly = "";
var m_sContactHistoryXML = "";
var m_sMetaAction = "";
var m_sCustomerName = "";
var m_sUserName = "";
var	m_sUnitName = "";
var m_sCreateRequest = null
var m_sCustomerID = "";	
var m_sUserID = "";
var m_sUnitID = "";		
var m_sUserRole = "";
var m_BaseNonPopupWindow = null;
		

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	//removed as per Core V7.0.2
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	
	RetrieveData();

	<%	//This function is contained in the field attributes file (remove if not required) %>
	SetMasks();
	Validation_Init();

	SetScreenOnReadOnly();
	GetComboList();
	PopulateScreen();
	window.returnValue = null;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function SetScreenOnReadOnly()
{
	scScreenFunctions.ShowCollection(frmScreen);

	if (m_sReadOnly == "1")
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
}

function RetrieveData()
{
	var sArguments = window.dialogArguments;
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters		= sArguments[4];
	
	//next 2 lines added as per Core V7.0.2
	m_BaseNonPopupWindow = sArguments[5];
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();

	m_sCreateRequest	= sParameters[0];
	m_sReadOnly			= sParameters[1];
	m_sContactHistoryXML= sParameters[2]; 
	m_sMetaAction		= sParameters[3]; 
	m_sCustomerName		= sParameters[4]; 		

	if (m_sMetaAction == "Add")
	{
		m_sUserName		= sParameters[5]; 
		m_sUnitName		= sParameters[6]; 
		m_sCustomerID	= sParameters[7]; 
	}
	
	m_sUserID			= sParameters[8]; 
	m_sUnitID			= sParameters[9]; 
	m_sUserRole			= sParameters[10]; 
}

function PopulateScreen()
{	
	frmScreen.txtCustomer.value = m_sCustomerName;

	<% /* BMIDS00005 View should also display the data passed in */ %>		
	if ((m_sMetaAction == "Edit") || (m_sMetaAction == "View"))
	{
		//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
		//var XML = new scXMLFunctions.XMLObject();
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.LoadXML(m_sContactHistoryXML);		
		if (XML.SelectTag(null,"CONTACTHISTORY") != null)
		{
			m_sCustomerID = XML.GetTagText("CUSTOMERNUMBER"); 
			frmScreen.txtOtherReason.value = XML.GetTagText("OTHERREASONTEXT"); 
			frmScreen.txtUserName.value = XML.GetTagText("USERNAME"); 
			frmScreen.txtUnitName.value = XML.GetTagText("UNITNAME"); 
			frmScreen.txtDateTime.value = XML.GetTagText("CONTACTHISTORYDATETIME"); 
			frmScreen.txtDetails.value = XML.GetTagText("CONTACTTEXT"); 
			frmScreen.cboReason.value = XML.GetTagText("CONTACTREASONCODE"); 
			frmScreen.txtOtherReason.value = XML.GetTagText("OTHERREASONTEXT"); 
			
			scScreenFunctions.SetCheckBoxValue(frmScreen, "chkCheckBox", XML.GetTagText("STATUSINDICATOR"));

			<% /* BMIDS00005 If vieweing set all controls to read only */ %>
			if (m_sMetaAction == "View")
				scScreenFunctions.SetScreenToReadOnly(frmScreen);			
			else
			{
				var sContactHistoryDeleteAuthority = XML.GetGlobalParameterAmount(document,"ContactHistoryDeleteAuthority");
				if (parseInt(m_sUserRole) < parseInt(sContactHistoryDeleteAuthority))
					scScreenFunctions.SetFieldState(frmScreen, "chkCheckBox", "R");	
				
				<% /* BMIDS00005 Editing should allow the reason, other reason and details to be changed		
				scScreenFunctions.SetFieldState(frmScreen, "cboReason", "R");
				scScreenFunctions.SetFieldState(frmScreen, "txtDetails", "R");
				*/ %>
			}
			
		}
		else
		{
			alert("Contact History Record not found");
		}
	
		//frmScreen.chkCheckBox.focus();
	}
	else
	{
		scScreenFunctions.SetFieldState(frmScreen, "chkCheckBox", "R");	
		frmScreen.txtUserName.value = m_sUserName; 
		frmScreen.txtUnitName.value = m_sUnitName; 
		frmScreen.txtDetails.focus();
		
	}
	scScreenFunctions.SetFieldState(frmScreen, "txtCustomer", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtUserName", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtUnitName", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtDateTime", "R");	
	<% /* BMIDS00303 Calling SetRequired here incorrectly clears existing data in OtherReason
	SetRequired();  //BMIDS00151 - DPF 3/7/02 - added to ensure Other Reason text box is only enabled if it should be
	*/ %>
	
	<% /* BMIDS00303 Other Reason should only be enabled on entry when editing and the Reason
	validation type is either O or C. For adding it will be enabled when cboReason changes appropriately */ %>
	if ((m_sMetaAction == "Edit") && (scScreenFunctions.IsValidationType(frmScreen.cboReason,"O") || scScreenFunctions.IsValidationType(frmScreen.cboReason,"C")))
	{
		scScreenFunctions.SetFieldState(frmScreen, "txtOtherReason", "W");
		spnOtherReason.style.color = "red";
	}
	else
		scScreenFunctions.SetFieldState(frmScreen, "txtOtherReason", "R");
	<% /* BMIDS00303 End */ %>
}

function GetComboList()
{
	//next line replaced by line below as per COre V7.0.2 - DPF 20/06/02 - BMIDS00077
	//var XML = new scXMLFunctions.XMLObject();
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sCombo = new Array("ContactHistoryReason");
	var bSuccess = false;

	if(XML.GetComboLists(document,sCombo))
	{
		bSuccess = true;
		bSuccess = XML.PopulateCombo(document,frmScreen.cboReason,"ContactHistoryReason",true);
		
		<% /* BMIDS00005 Remove CRS contact types (validation type = 'C') from the combo when adding and editing*/ %>
		if ((m_sMetaAction == "Add") || (m_sMetaAction == "Edit"))
		{
			for (var iIndex = 0; iIndex < frmScreen.cboReason.length; iIndex++)
			{	
				if (scScreenFunctions.IsOptionValidationType(frmScreen.cboReason, iIndex , "C"))
					frmScreen.cboReason.remove(iIndex);
			}
		}
	}

	if(!bSuccess)
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnOK.disabled = true;
	}
}

function SetRequired()
{
	if (scScreenFunctions.IsValidationType(frmScreen.cboReason,"O"))
	{  
		scScreenFunctions.SetFieldState(frmScreen, "txtOtherReason", "W");
		frmScreen.txtOtherReason.value = "";	
		spnOtherReason.style.color = "red";
	}
	else
	{		
		scScreenFunctions.SetFieldState(frmScreen, "txtOtherReason", "D");		
		spnOtherReason.style.color = "darkblue";
	}
}

function frmScreen.cboReason.onchange()
{
	SetRequired();
}

function frmScreen.btnCancel.onclick()
{
	window.returnValue	= null;
	window.close();
}

function frmScreen.btnOK.onclick()
{
	<% /* BMIDS000005 OK button just closes the window when viewing */ %>
	if (m_sMetaAction == "View")
	{
		window.returnValue	= null;
		window.close();
	}
	else
	{
		if (frmScreen.onsubmit()== false)
			return;
	
		//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
		//var XML = new scXMLFunctions.XMLObject();
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTagFromArray(m_sCreateRequest, m_sMetaAction);
		XML.CreateActiveTag("CONTACTHISTORY");
		XML.CreateTag("CUSTOMERNUMBER", m_sCustomerID);
		XML.CreateTag("CONTACTHISTORYDATETIME", frmScreen.txtDateTime.value);
		
		<% /* ASu BMIDS00153 - Start*/ %>
		if (frmScreen.onsubmit() == true)
		{
			bSuccess = true;
			if(RestrictLength(frmScreen, "txtDetails", 2000, true))
				bSuccess = false;
		}
		<% /* ASu BMIDS00153 - End*/ %>
				
		if (m_sMetaAction == "Edit")
		{	
			<% /* BMIDS00005 Edit should also save reason, other reason and details */ %>
			XML.CreateTag("CONTACTREASONCODE",frmScreen.cboReason.value);
			XML.CreateTag("OTHERREASONTEXT",frmScreen.txtOtherReason.value);
			XML.CreateTag("CONTACTTEXT",frmScreen.txtDetails.value);
			
			if (scScreenFunctions.GetCheckBoxValue(frmScreen, "chkCheckBox") == true)
				XML.CreateTag("STATUSINDICATOR","1");
			else
				XML.CreateTag("STATUSINDICATOR","0");
				
			// 			XML.RunASP(document, "EditContactHistory.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document, "EditContactHistory.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}

		}
		else
		{
			if (scScreenFunctions.IsValidationType(frmScreen.cboReason,"O") && frmScreen.txtOtherReason.value == "" )
			{
				alert("Other Reason needs completing");
				return;
			}
			else
			{
				XML.CreateTag("CONTACTREASONCODE",frmScreen.cboReason.value);
				XML.CreateTag("OTHERREASONTEXT",frmScreen.txtOtherReason.value);
				XML.CreateTag("CONTACTTEXT",frmScreen.txtDetails.value);
				XML.CreateTag("USERID",m_sUserID);
				XML.CreateTag("UNITID",m_sUnitID);
				XML.CreateTag("STATUSINDICATOR","0");
				// 				XML.RunASP(document, "AddContactHistory.asp");
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
					{
					case 1: // Warning
					case 0: // OK
									XML.RunASP(document, "AddContactHistory.asp");
						break;
					default: // Error
						XML.SetErrorResponse();
					}

			}	
		}
		if (XML.IsResponseOK()==true)
		{
			var sReturn = new Array();
			sReturn[0] = IsChanged();
			window.returnValue	= sReturn;
			window.close();
		}
	}
}
<% /* ASu BMIDS00153 - Start. Note* This function exists within scScreenFunctions.asp, however if an error message is raised,
 it is hidden behind the pop ups due to scScreenFunctions being moved into framework. This is the quickest and easiest solution at present */ %>
function RestrictLength(frmScreen, sFieldId, nMaxLength, bMessage)
{
	var thisObj = frmScreen.all(sFieldId);
	if(thisObj.tagName != "TEXTAREA" || nMaxLength < 2)
		return false;
		
	var strString = thisObj.value;
	if(strString.length > nMaxLength)
	{
		thisObj.value = strString.slice(0, nMaxLength);
		if(bMessage == true)
		{
			alert("The maximum length of " + nMaxLength + " characters has been exceeded. The overflow has been truncated");
			thisObj.focus();
		}
		return true;
	}
	else
		return false;
}
<% /* ASu BMIDS00153 - End */ %>	
-->
</script>
</body>
</html>



