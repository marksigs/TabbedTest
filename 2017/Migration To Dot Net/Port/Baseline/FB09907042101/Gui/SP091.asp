<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      SP091.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Cheque payment Authorisation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AW		05/08/2001	BMID029  created 
GD		04/10/2002	BMIDS00580 - Added field to display last cheque number.
GD		14/10/2002	BMIDS00572	- Added more graceful handling of wrong password.
PJO     13/10/2003  BMIDS643	- Password length increased to 15
MC		21/04/2004	BMIDS517	- Background DIV size changed to fix title text (IE 6.0 Display bigfont)
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
<title>Sanction/Unsanction Authorisation <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen"  year4 validate="onchange" mark><!--style="VISIBILITY: hidden"-->
	<div id="divSanctionDetails" style="HEIGHT: 210px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 365px" class="msgGroup">

		<span style="LEFT: 20px; POSITION: absolute; TOP: 15px" class="msgLabel">
			Total number of cheque payments
			<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
				<input id="txtNumberOfPayments" maxlength="10" style="WIDTH: 40px" class ="msgReadOnly" readonly tabindex="-1">
			</span>
		</span>
		<span style="LEFT: 20px; POSITION: absolute; TOP: 45px" class="msgLabel">
			Total value of cheque payments
			<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
				<input id="txtValueOfPayments" maxlength="10" style="WIDTH: 80px" class ="msgReadOnly" readonly tabindex="-1">
			</span>
		</span>
		<!--GD BMIDS00580-->
		<span style="LEFT: 20px; POSITION: absolute; TOP: 75px" class="msgLabel">
			Last Cheque Number
			<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
				<input id="txtLastChequeNumber" maxlength="10" style="WIDTH: 90px" class ="msgReadOnly" readonly>
			</span>
		</span>		
		
		<span style="LEFT: 20px; POSITION: absolute; TOP: 105px" class="msgLabel">
			<label id="idChequeNo"></label>
			Enter the number of the first cheque
			<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
				<input id="txtChequeNum" maxlength="10" style="WIDTH: 90px" class ="msgTxt">
			</span>
		</span>
		<span style="LEFT: 20px; POSITION: absolute; TOP: 135px" class="msgLabel">
			User Password
			<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
	<% /* PJO 13/10/2003 BMIDS643 - Password length increased to 15 */ %>
			<input id="txtPassword" type="password" maxlength="15" style="WIDTH: 100px" class="msgTxt">
			</span>
		</span><!-- Buttons -->
		<span id="spnButtons" style="LEFT: 65px; POSITION: absolute; TOP: 175px">
			<span style="LEFT: 1px; POSITION: absolute; TOP: 0px">
				<input id="btnOK" value="OK" type="button" 
					style="WIDTH: 80px" class="msgButton">
				</span>
			<span style="LEFT: 104px; POSITION: absolute; TOP: 0px">
				<input id="btnCancel" value="Cancel" type="button" 
					style="WIDTH: 80px" class="msgButton">
			</span>
		</span>
	</div>
</form>

<% /* End of form */ %>
<!-- #include FILE="attribs/sp091attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_ArraySelectedRows
var m_blnReadOnly = false;
var m_sUserID = "0";
var m_sUnitID = "0";
var m_nFirstChequeNumber = "0";
//GD BMIDS00580
var m_sStoredNum = "";

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();

	//frmScreen.txtValueOfPayments = null;
	//frmScreen.txtNumberOfPayments = null; 
	scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	SetMasks();
	Validation_Init();
	RetrieveData();
	scScreenFunctions.ShowCollection(frmScreen);
	frmScreen.txtLastChequeNumber.value = m_sStoredNum;
	frmScreen.txtChequeNum.focus()
	window.returnValue = null;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function RetrieveData()
{

	var ArraySelectedRows = new Array();
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];

	var sParameters	= sArguments[4];
	frmScreen.txtValueOfPayments.value = sParameters[0];
	frmScreen.txtNumberOfPayments.value = sParameters[1];
	
	if (sParameters[1] == 1) {idChequeNo.nextSibling.nodeValue = "Enter the cheque number"}

	m_ArraySelectedRows = sParameters[2];
	m_sUserID = sParameters[3];
	m_sUnitID = sParameters[4];
	//GD BMIDS00580
	FindUnitsForUser();
}

function frmScreen.btnOK.onclick()
{
	if (frmScreen.txtChequeNum.value == "")
	{
		alert("Please enter the first cheque number.");
		frmScreen.txtChequeNum.focus;
		//exit function;
		return;
	}
	
	if (frmScreen.txtPassword.value == "")
	{
		alert("Please enter your user password.");
		frmScreen.txtPassword.focus;
		//exit function;
		return;
	}
	
	ValidatePassword();	
}

function frmScreen.btnCancel.onclick()
{
	var sReturn = new Array();
	sReturn[0] = false;
	window.returnValue = sReturn;
	window.close();
}

function ValidatePassword()
{
	var XML = new scXMLFunctions.XMLObject();
	XML.CreateActiveTag("REQUEST");
	XML.SetAttribute("USERID", m_sUserID);
	XML.CreateActiveTag("OMIGAUSER");
	XML.CreateTag("USERID", m_sUserID);
	XML.CreateTag("PASSWORDVALUE", frmScreen.txtPassword.value);
	XML.CreateTag("UNITID", m_sUnitID);
	XML.CreateTag("AUDITRECORDTYPE", "1");
	// 	XML.RunASP(document, "ValidateCurrentPassword.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "ValidateCurrentPassword.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}


//GD BMIDS00572 START
	//if (XML.IsResponseOK() == true)
	//{
	//	ValidateChequeNumber();
	//}
	var ErrorTypes = new Array("WRONGPASSWORD");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[1] == ErrorTypes[0])
		alert("Incorrect Password details have been entered.  Please check your Password is correct before retrying.");
	else if(ErrorReturn[0] == true)
	{
		ValidateChequeNumber();
	}
//GD BMIDS00572 END
}

function ValidateChequeNumber()
{				
	<%// BMIDS00580 Moved to FindUnitsForUser()
				
	//var XML = new scXMLFunctions.XMLObject();

	//XML.CreateActiveTag("REQUEST");
	//XML.SetAttribute("USERID", m_sUserID);
	//XML.CreateActiveTag("SEARCH");
	//XML.CreateTag("USERID", m_sUserID);
	//XML.CreateTag("UNITID", m_sUnitID);
				
	//XML.RunASP(document, "FindUnitsForUser.asp");
					
	//if (XML.IsResponseOK() == true)
	//{
		//var sStoredNum = XML.GetTagText("UNHIGHCHEQUENUMBER");%>
		var sEnteredNum = frmScreen.txtChequeNum.value;
		
		m_nFirstChequeNumber = parseInt(sEnteredNum, 10);
		
		if(isNaN(m_sStoredNum, 10))
		{
			alert("Field UNHIGHCHEQUENUMBER on unit table contains non numeric data"); 
			frmScreen.txtChequeNum.value = "";
			return;
		}
		else
		{
			if(m_nFirstChequeNumber <= parseInt(m_sStoredNum, 10))	
			{
				alert("Enter a cheque number greater than the number currently stored for this unit"); 
				frmScreen.txtChequeNum.value = "";
				return;
			}
			else
			{
				var sReturn = new Array();

				sReturn[0] = true;
				sReturn[1] = m_nFirstChequeNumber;
				window.returnValue = sReturn;
				window.close();
			}
		}		
	//}
				
	//XML = null;
}	
//GD BMIDS00580	
function FindUnitsForUser()
{
	var XML = new scXMLFunctions.XMLObject();

	XML.CreateActiveTag("REQUEST");
	XML.SetAttribute("USERID", m_sUserID);
	XML.CreateActiveTag("SEARCH");
	XML.CreateTag("USERID", m_sUserID);
	XML.CreateTag("UNITID", m_sUnitID);
				
	// 	XML.RunASP(document, "FindUnitsForUser.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "FindUnitsForUser.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	if (XML.IsResponseOK() == true)
	{
		m_sStoredNum = XML.GetTagText("UNHIGHCHEQUENUMBER");
		return(true);
	} else return(false);
}

-->
</script>
</body>
</html>


