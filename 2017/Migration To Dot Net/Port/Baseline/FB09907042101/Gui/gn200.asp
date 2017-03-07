<%@ LANGUAGE="JSCRIPT" %>
<html>
<%

/*
Workfile:      gn200.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Application Enquiry Quick Search screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MV		30/10/00	Created
MV		14/11/00	Added Cancel Button 
					Added Code to Build XML string on btnSearch_Click event
AS		22/11/00	CORE000012: Added Message box to inform user that 
					Application Number takes precedence
AS		28/11/00	CORE000010: Streamlined XML request to GN215 and added UNITID and USERID tags
AS		09/01/01	SYS1782: Back button processing to use idXML
CL		05/03/01	SYS1920 Read only functionality added
AS		11/03/01	SYS2030: Declined/Cancelled reinstate processing
DRC     29/01/02    SYS3547: Changed user name text box to combo to allow wider search
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		15/08/2002  BMIDS00119  Core Upgrade Ref : SYS4612 
DRC     18/09/2002  BMIDS00119  Handle "<ALL>" User search
CL		15/10/2002	BMIDS00557  Change to button search
MDC		15/11/2002	BMIDS00958	Override User/Unit when Application or Account Number specified
GHun	27/11/2002	BM0035		CC011 Customer business search in application enquiry
GHun	18/12/2002	BM0034		Show UserName instead of UserID in user name combo
GHun	16/01/2003	BM0252		Minor fixes for CC011
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

HMA		20/09/2005	MAR46		Include Product Switch and Transfer of Equity complete cases in search.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


*/ %>
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
<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<form id="frmTogn215" method="post" action="gn215.asp" STYLE="DISPLAY: none"></form>
<form id="frmTogn210" method="post" action="gn210.asp" STYLE="DISPLAY: none"></form>
<form id="frmWelcomeMenu" method="post" action="mn010.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span> 
<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>

<form id="frmScreen" validate   ="onchange" mark>

<% /* Application Details */ %>
<div style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 210px" class="msgGroup" id="DIV1">
<span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 12px" class="msgLabel">
		Application Number
		<span style="LEFT: 102px; POSITION: absolute; TOP: -3px">
			<input id="txtApplicationNumber" maxlength="12" style="WIDTH: 130px; POSITION: absolute" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 365px; POSITION: absolute; TOP: 12px" class="msgLabel">
		Account Number
		<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
			<input id="txtAccountNumber" maxlength="20" style="WIDTH: 136px; POSITION: absolute" class="msgTxt">
		</span>
	</span>
	
	<% /* BM0035 */ %>
	<span style="TOP: 40px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Surname
		<span style="TOP: -3px; LEFT: 102px; POSITION: ABSOLUTE">
			<input id="txtSurname" maxlength="40" style="WIDTH: 220px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 68px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Forename(s)
		<span style="TOP: -3px; LEFT: 102px; POSITION: ABSOLUTE">
			<input id="txtForenames" maxlength="40" style="WIDTH: 220px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 40px; LEFT: 365px; POSITION: ABSOLUTE" class="msgLabel">
		Date of Birth
		<span style="TOP: -3px; LEFT: 90px; POSITION: ABSOLUTE">
			<input type="text" id="txtDateOfBirth" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 68px; LEFT: 365px; POSITION: ABSOLUTE" class="msgLabel">
		PostCode
		<span style="TOP: -3px; LEFT: 90px; POSITION: ABSOLUTE">
			<input id="txtPostcode" maxlength="8" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxtUpper">
		</span>
	</span>	
	<% /* BM0035 End */ %>
	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 96px" class="msgLabel">
		Unit Name
		<span style="LEFT: 102px; POSITION: absolute; TOP: -3px">
			<input id="txtUnitName" maxlength="45" style="WIDTH: 220px; POSITION: absolute; HEIGHT: 20px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 365px; POSITION: absolute; TOP: 96px" class="msgLabel">
		User Name
		<span style="LEFT: 84px; POSITION: absolute; TOP: -3px">&nbsp;
		    <%/* DRC SYS3547 Change from text box to combo */%>
		    <select id="cboUserName" style="WIDTH: 136px; POSITION: absolute" class="msgCombo"></select>
			<%/*<input id="txtUserName" maxlength="10" style="LEFT: 30px; WIDTH: 150px; POSITION: absolute; TOP: -2px" class="msgTxt"> */%>
		</span>
	</span>
</span>
	<span style="LEFT: 5px; WIDTH: 300px; POSITION: absolute; TOP:  130px; HEIGHT: 22px" class="msgLabelHead" >
		<input id="ChkCancelledorDeclinedApplications" type="checkbox" value="1">
		<label for="ChkCancelledorDeclinedApplications">Include Cancelled/Declined Applications</label>
	</span>

	<span style="LEFT: 5px; WIDTH: 300px; POSITION: absolute; TOP:  155px; HEIGHT: 22px" class="msgLabelHead" >
		<input id="ChkProductSwitchApplications" type="checkbox" value="1">
		<label for="ChkProductSwitchApplications">Include Product Switch Completed Applications</label>
	</span>

	<span style="LEFT: 5px; WIDTH: 300px; POSITION: absolute; TOP:  180px; HEIGHT: 22px" class="msgLabelHead" >
		<input id="ChkTOEApplications" type="checkbox" value="1">
		<label for="ChkTOEApplications">Include Transfer of Equity Completed Applications</label>
	</span>

	<span style="LEFT: 405px; POSITION: absolute; TOP: 178px">
		<input id="btnSearch" value="Search" type="button" style="WIDTH: 75px" class="msgButton">
	</span>
	<span style="LEFT: 490px; POSITION: absolute; TOP: 178px">
		<input id="btnExtendedSearch" value="ExtendedSearch" type="button" style="WIDTH: 100px" class="msgButton">
	</span>
</div> 
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 290px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" --> 
</div> 


<% /* Span to keep tabbing within this screen  */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/gn200Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var scScreenFunctions;
var m_iAppEnqExtendedSearchRole;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Application Enquiry - Quick Search ","GN200",scScreenFunctions); 
	//DRC  SYS3547 added call to PopulateUserListCombo() function
	PopulateUserListCombo();
	RetrieveContextData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	GetParameterValues();
	ValidateUserAuthorityLevel();
	
	scScreenFunctions.SetFieldState(frmScreen, "txtUnitName", "R");
	//DRC  SYS3547  txt field no longer exists
	//scScreenFunctions.SetFieldState(frmScreen, "txtUserName", "R");
	scScreenFunctions.SetFocusToFirstField(frmScreen);	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	// APS SYS1782
	// var sXML  = scScreenFunctions.GetContextParameter(window,"idXML2","");
	var sXML  = scScreenFunctions.GetContextParameter(window,"idXML","");
	if (sXML != "" ) 
		PopulateScreen(sXML);
	else
	{
		frmScreen.txtUnitName.value = scScreenFunctions.GetContextParameter(window,"idUnitName",null);
			//DRC  SYS3547  txt field no longer exists
	       //frmScreen.txtUserName.value = scScreenFunctions.GetContextParameter(window,"idUserName",null);
     		frmScreen.cboUserName.value = scScreenFunctions.GetContextParameter(window,"idUserID","");		       
	}
}

function frmScreen.btnSearch.onclick()
{	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bValidCriteria = true;	<% /* BM0035 */ %>
		
	if ((frmScreen.txtApplicationNumber.value != "") && (frmScreen.txtAccountNumber.value != "")) 
	{
		alert ("Application Number will take precedence in Search");	
		frmScreen.txtAccountNumber.value = "";
	}
	
	//Preparing XML Request string 
	XML.CreateRequestTag(window,"QUICKSEARCH");
	XML.CreateTag("APPLICATIONNUMBER",frmScreen.txtApplicationNumber.value);
	XML.CreateTag("ACCOUNTNUMBER",frmScreen.txtAccountNumber.value);
	<% /* MV - 15/08/2002 - BMIDS00119 - Core Upgrade Ref : SYS4612 */ %>
	if ((frmScreen.txtApplicationNumber.value.length == 0) && (frmScreen.txtAccountNumber.value.length == 0))  
	{	
		<% /* BM0035 Add customer details to search criteria */ %>			
		if (IsCorrectWildcard())
		{
			if ((!IsSurnameRequired()) && (!IsForenameRequired()))
			{
				XML.CreateTag("SURNAME", frmScreen.txtSurname.value);
	
				<% /* Split the forename field on white space characters */ %>
				var sForenames = frmScreen.txtForenames.value.split(/\s+/); 
				XML.CreateTag("FIRSTFORENAME", sForenames[0]);
							
				if (sForenames[1] != null) 
					XML.CreateTag("SECONDFORENAME", sForenames[1]);
	
				if (sForenames[2] != null)
				{
					<% /* Any forenames left after the second one are added to Other forenames */ %>
					var sOtherNames = "";
					for(var i=2; i <= (sForenames.length - 1); i++)
						sOtherNames += sForenames[i] + " ";
					XML.CreateTag("OTHERFORENAMES", sOtherNames);
				}
	
				XML.CreateTag("DATEOFBIRTH", frmScreen.txtDateOfBirth.value);
				XML.CreateTag("POSTCODE", frmScreen.txtPostcode.value);
			}
			else
				bValidCriteria = false;
		}
		else
			bValidCriteria = false;
		
		<% /* BM0035 End */ %>	
	}
	<% /* BM0252 Unit is now set is the same place as user
	else
	{
		//BMIDS00557 15/10/02
		//XML.CreateTag("UNITID");
		//XML.CreateTag("UNITNAME"); 
		
		XML.CreateTag("UNITID", "");
		XML.CreateTag("UNITNAME", ""); 
		//BMIDS00557 15/10/02 END
	}
	*/ %>
	
	<% /* BM0035 */ %>
	if (bValidCriteria)
	{
	<% /* BM0035 End */ %>
		
		<% /* BMIDS00958 MDC 15/11/2002 - If application/account specified, override user/unit */ %>
		// if ((frmScreen.cboUserName.value != "<ALL>" ) && (frmScreen.cboUserName.value != "")) 
		<% /* BM0252 Don't include user & unit if surname or forenames exist */ %>
		if ((frmScreen.cboUserName.value != "<ALL>" ) && (frmScreen.cboUserName.value != "")
				&& (frmScreen.txtApplicationNumber.value.length == 0) 
				&& (frmScreen.txtAccountNumber.value.length == 0)
				&& (frmScreen.txtSurname.value.length == 0)
				&& (frmScreen.txtForenames.value.length == 0)) 
		{
			XML.CreateTag("USERID",frmScreen.cboUserName.value);
			//DRC SYS3547 txt field no longer exists
			//XML.CreateTag("USERNAME",frmScreen.txtUserName.value);
			//This Username is only used in the next screen GN215 and not in the search
			XML.CreateTag("USERNAME",frmScreen.cboUserName.item(frmScreen.cboUserName.selectedIndex).text);
			XML.CreateTag("UNITID",scScreenFunctions.GetContextParameter(window,"idUnitId",null));
			XML.CreateTag("UNITNAME",frmScreen.txtUnitName.value);
		}
		else
		{
		    XML.CreateTag("USERID");
		    XML.CreateTag("USERNAME", "");
	  		XML.CreateTag("UNITID", "");
			XML.CreateTag("UNITNAME", ""); 
		}

		if (frmScreen.ChkCancelledorDeclinedApplications.checked == true) 
		{
			XML.CreateTag("INCCANCELLEDORDECLINEDAPPSCHECKED","1");
			XML.CreateTag("INCLUDECANCELLEDAPPS","1");
			XML.CreateTag("INCLUDEDECLINEDAPPS","1");
		}
		
		<% /* MAR46 Add Product Switch and Transfer of Equity searches */ %>
		if (frmScreen.ChkProductSwitchApplications.checked == true) 
		{
			XML.CreateTag("INCPRODUCTSWITCHAPPSCHECKED","1");
		}
		if (frmScreen.ChkTOEApplications.checked == true) 
		{
			XML.CreateTag("INCTOEAPPSCHECKED","1");
		}
		
		scScreenFunctions.SetContextParameter(window,"idXML",XML.ActiveTag.xml);
		// APS SYS1782
		//scScreenFunctions.SetContextParameter(window,"idXML2",XML.XMLDocument.xml);
		scScreenFunctions.SetContextParameter(window,"idMetaAction","QUICKSEARCH");
	
		frmTogn215.submit();
	}	<% /* BM0035 */ %>
}

function frmScreen.btnExtendedSearch.onclick()
{	
	// On Click Extended Search Button Navigate to GN210 Screen
	// APS 09/01/2001 : SYS1782 We do not want to try to populate gn210 with
	// information from gn200
	scScreenFunctions.SetContextParameter(window,"idXML","");
	frmTogn210.submit();
}

function btnCancel.onclick()
{
	// On Click cancel button Navigate to Main Menu
	scScreenFunctions.SetContextParameter(window,"idXML","");
	frmWelcomeMenu.submit();				
}

function PopulateScreen(sXML) 
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
	XML.LoadXML(sXML);
	if (XML.SelectTag(null,"REQUEST") != null)
	{
		frmScreen.txtApplicationNumber.value  = XML.GetTagText("APPLICATIONNUMBER");
		frmScreen.txtAccountNumber.value = XML.GetTagText("ACCOUNTNUMBER");
		frmScreen.txtUnitName.value = XML.GetTagText("UNITNAME");
		//DRC SYS3547 txt field no longer exists
		//frmScreen.txtUserName.value = XML.GetTagText("USERNAME"); */ %>
		var sUserID = XML.GetTagText("USERID");
		if (sUserID == "" )
		{
		  frmScreen.cboUserName.value = "<ALL>";
		}
		else
		{
		   frmScreen.cboUserName.value = sUserID;
		}
		if ( XML.GetTagText("INCCANCELLEDORDECLINEDAPPSCHECKED")  ==  "1" ) 	
			frmScreen.ChkCancelledorDeclinedApplications.checked = true; 
			
		<% /* MAR46 Add Product Switch and Transfer of Equity switches*/ %>			
		if ( XML.GetTagText("INCPRODUCTSWITCHAPPSCHECKED")  ==  "1" ) 	
			frmScreen.ChkProductSwitchApplications.checked = true; 

		if ( XML.GetTagText("INCTOEAPPSCHECKED")  ==  "1" ) 	
			frmScreen.ChkTOEApplications.checked = true; 
			
		<% /* BM0252 Repopulate search criteria when returning back from GN210 or GN215 */ %>
		frmScreen.txtSurname.value = XML.GetTagText("SURNAME");
		var sForenames = XML.GetTagText("FIRSTFORENAME") + " " + XML.GetTagText("SECONDFORENAME") + " " + XML.GetTagText("OTHERFORENAMES");
		if (sForenames.length > 2)
			frmScreen.txtForenames.value = sForenames;
		frmScreen.txtPostcode.value = XML.GetTagText("POSTCODE");
		frmScreen.txtDateOfBirth.value = XML.GetTagText("DATEOFBIRTH");
		<% /* BM0252 End */ %>
	}
	XML = null;	
}

function GetParameterValues()
{
	// This is a function to retrieve GlobalParameters from the database to validate authority level
	//save the combo settings in idXML context to use when we return
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var blnReturn = false;	
	m_iContextRole = scScreenFunctions.GetContextParameter(window,"idRole",null);
	
	//Preparing XML Request string 
	XML.CreateRequestTag(window,"SEARCH");
	XML.CreateActiveTag("GLOBALPARAMETER");
	XML.CreateTag("NAME","APPENQEXTENDSEARCHROLE");
	XML.RunASP(document,"FindCurrentParameterList.asp");
	if (XML.IsResponseOK()== true)
	{	
		blnReturn = true;
		if(XML.SelectTag(null,"GLOBALPARAMETERLIST") != null)
				m_iAppEnqExtendedSearchRole = XML.GetTagText("AMOUNT");
	}
	else 
		blnReturn = false;

	XML = null; 
	return blnReturn;
}

function ValidateUserAuthorityLevel()
{
	frmScreen.btnExtendedSearch.disabled = false;  
	if (parseInt(m_iContextRole) < parseInt(m_iAppEnqExtendedSearchRole)  )
	{
		frmScreen.btnExtendedSearch.disabled = true;
		//DRC AQR3547 - restrict access to other user's data by freezing the combo
		scScreenFunctions.SetFieldState(frmScreen, "cboUserName", "R"); 
	}	
}

//DRC AQR3547 added function 
function PopulateUserListCombo()
{
	// This is a Function to Populate userList Combo with userId and Usernames
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	//Preparing XML Request string 
	XML.CreateRequestTag(window);
	XML.CreateActiveTag("USERLIST");
	XML.CreateTag("UNITID",scScreenFunctions.GetContextParameter(window,"idUnitID",null));
	XML.RunASP(document,"FindUserList.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if(ErrorReturn[0])
	{	
		<% /* BM0034 <ALL> should be displayed at the top of the list */ %>
		var TagOPTION = document.createElement("OPTION");
		TagOPTION.value = "<ALL>"
		TagOPTION.text  = "<ALL>";
		frmScreen.cboUserName.add(TagOPTION);
		TagOPTION = null;
		
		<% /* BM0034 */ %>
		<%
		// PopulateComboFromXML can no longer be used as 2 values need to concatenated to display the name
		// blnSuccess = XML.PopulateComboFromXML(document,frmScreen.cboUserName,XML.XMLDocument,false,false,"OMIGAUSER","USERID","USERID");
		%>
		var TagListLISTENTRY = XML.CreateTagList("OMIGAUSER");
		XML.ActiveTagList = TagListLISTENTRY;
		
		for(var nLoop = 0;nLoop < TagListLISTENTRY.length;nLoop++)
		{	
			XML.SelectTagListItem(nLoop);
			TagOPTION = document.createElement("OPTION");
				
			TagOPTION.value = XML.GetTagText("USERID");
			TagOPTION.text = XML.GetTagText("USERFORENAME") + " " + XML.GetTagText("USERSURNAME");;
			frmScreen.cboUserName.options.add(TagOPTION);				
		}

		frmScreen.cboUserName.selectedIndex = 0;
		<% /* BM0034 End */ %>
	}	
	else
	if(ErrorReturn[1] == ErrorTypes[0])
	{	
		alert("No Users are available");		
	}
	XML= null;
}

<% /* BM0035 Validation for customer related search criteria - based on version in CR010 */ %>
function IsCorrectWildcard()
{
<% /*	Checks the fields which may be wildcarded to ensure that the wildcarding is valid
		The postcode field must have at least one character preceding the wildcard
		and there must be no characters following it.
		The surname fields must have at least the number of characters specified by the
		GlobalParameter MinCharsBeforeSurnameWildcard preceding the wildcard and there must be
		no characters following it.
		The forenames field does not require a preceding character.  It does, however, allow multiple
		wildcards in order to search on first forename and other forenames.  The rule here is that
		a wildcard may only be immediately followed by a space character.
	*/
%>

	<% /* Surname checks */ %>
	var sField = frmScreen.txtSurname.value;
	if (sField.length > 0)
	{
		var GlobalXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var iMinChars = parseInt(GlobalXML.GetGlobalParameterAmount(document,"MinCharsBeforeSurnameWildcard"));
	
		var iWildcardPos = sField.indexOf("*");
		if (iWildcardPos != -1) {
			if(iWildcardPos < iMinChars)
			{
				alert("You must enter at least " + iMinChars + " preceding character(s) to wildcard the Surname.");
				frmScreen.txtSurname.focus();
				return false;
			}

			if(iWildcardPos < sField.length - 1)
			{
				alert("There must be no characters after the wildcard for the Surname.");
				frmScreen.txtSurname.focus();
				return false;
			}
		}
	}
	
	<% /* Forename checks */ %>
	sField = frmScreen.txtForenames.value;
	var nWildcardIndex = 0;
	do
	{
		nWildcardIndex = sField.indexOf("*",nWildcardIndex);
		if(nWildcardIndex != -1)
		{
			if(nWildcardIndex < sField.length - 1)
				if(sField.charAt(nWildcardIndex + 1) != ' ')
				{
					alert("A wildcard is being followed by a character other than a space in the Forename(s).");
					frmScreen.txtForenames.focus();
					return false;
				}

			nWildcardIndex++;
		}
	}
	while(nWildcardIndex != -1)

	<% /* Postcode checks */ %>
	sField = frmScreen.txtPostcode.value;
	if(sField.indexOf("*") == 0)
	{
		alert("You must enter at least 1 preceding character to wildcard the Postcode.");
		frmScreen.txtPostcode.focus();
		return false;
	}

	if(sField.indexOf("*") < sField.length - 1 && sField.indexOf("*") != -1)
	{
		alert("There must be no characters after the wildcard for the Postcode.");
		frmScreen.txtPostcode.focus();
		return false;
	}

	return true;
}

function IsSurnameRequired()
{			
	if	((frmScreen.txtSurname.value == "") && ((frmScreen.txtForenames.value != "") || (frmScreen.txtDateOfBirth.value != "")))
	{
		alert ("You must enter Surname for Forename(s)/Date of Birth.");
		frmScreen.txtSurname.focus();
		return true;
	}

	return false;
}

function IsForenameRequired()
{
	if ((frmScreen.txtForenames.value == "") && (frmScreen.txtSurname.value != ""))
	{
		alert ("You must enter Forename(s) (Name or Initial or Wildcard) with Surname.");
		frmScreen.txtForenames.focus();
		return true;	
	}
	return false;
}
<% /* BM0035 End */ %>

-->
</script>
</body>
</html>
