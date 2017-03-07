<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /* 
Workfile:      DC240.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Third Party Summary
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		18/01/2000	Created
AD		02/02/2000	Rework
AY		14/02/00	Change to msgButtons button types
RF		01/03/00	Route to DC220 on cancel as DC230 is currently out of scope.
AD		13/03/00	Fixed SYS0454
AY		31/03/00	New top menu/scScreenFunctions change
MC		27/04/00	Route to DC295 on OK
MC		04/05/00	Fixed SYS0633
MC		04/05/00	Fixed SYS0632
BG		17/05/00	SYS0752 Removed Tooltips
CL		05/03/01	SYS1920 Read only functionality added
CL		12/03/01	SYS2034 Change to onload (Data freeze)
JLD		10/12/01	SYS2806 Completeness Check routing
LD		23/05/02	SYS4727 Use cached versions of frame functions
JLD		12/06/02	sys4728 use stylesheet at all times
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

BMids History

Prog	Date		Description
GHun	04/12/2002	BM0117 Disable buttons when the screen is read only
BM0193	16/12/2002	BM0193 Add ScreenRules to OK button.
JR		23/01/2003	BM0271 Add ProcessingIndicator check for disabling buttons.
BS		14/04/2003	BM0271 Disable Add button when data frozen
KRW     21/09/2004  BMIDS833 Disable overide when case cancelled/Frozen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History

Prog	Date		AQR		Description
MF		08/08/2005	MAR20	Parameterised routing for Global Parameter "ThirdPartySummary"
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom Specific History

Prog	Date		AQR		Description
AShaw	30/10/2006	EP2_8	Additional Borrowing code.
MHeys	05/12/2006	EP2_112 Added extra validtation to IsLegalRepToBeUsed functionality
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
<meta name=vs_targetSchema content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>
<span id="spnListScroll">
	<span style="LEFT: 290px; POSITION: absolute; TOP: 288px">
		<OBJECT data=scTableListScroll.asp id=scScrollTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>

<% /* FORMS */ %>
<form id="frmToDC220" method="post" action="DC220.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC230" method="post" action="DC230.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC295" method="post" action="DC295.asp" STYLE="DISPLAY: none"></form>
<form id="frmToMN060" method="post" action="MN060.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC260" method="post" action="DC260.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC290" method="post" action="DC290.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC250" method="post" action="DC250.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC270" method="post" action="DC270.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC280" method="post" action="DC280.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC201" method="post" action="DC201.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCM010" method="post" action="CM010.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* SCREEN LAYOUT */ %>
<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="HEIGHT: 268px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<strong>Third Party Summary</strong> 
	</span>

	<span id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 36px">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="15%" class="TableHead">Type</td>			<td width="20%" class="TableHead">Contact Name/ Account Number</td><td width="20%" class="TableHead">Company Name</td>	<td width="30%" class="TableHead">Address</td>		<td width="15%" class="TableHead">Repayment Bank Account</td></tr>
			<tr id="row01">		<td width="15%" class="TableTopLeft">&nbsp;</td>	<td width="20%" class="TableTopCenter">&nbsp;</td>				  <td width="20%" class="TableTopCenter">&nbsp;</td>	<td width="30%" class="TableTopCenter">&nbsp;</td>	<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		<td width="15%" class="TableLeft">&nbsp;</td>		<td width="20%" class="TableCenter">&nbsp;</td>					  <td width="20%" class="TableCenter">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>	    <td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		<td width="15%" class="TableLeft">&nbsp;</td>		<td width="20%" class="TableCenter">&nbsp;</td>					  <td width="20%" class="TableCenter">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>	    <td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		<td width="15%" class="TableLeft">&nbsp;</td>		<td width="20%" class="TableCenter">&nbsp;</td>					  <td width="20%" class="TableCenter">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>	    <td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		<td width="15%" class="TableLeft">&nbsp;</td>		<td width="20%" class="TableCenter">&nbsp;</td>					  <td width="20%" class="TableCenter">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>	    <td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		<td width="15%" class="TableLeft">&nbsp;</td>		<td width="20%" class="TableCenter">&nbsp;</td>					  <td width="20%" class="TableCenter">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>	    <td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		<td width="15%" class="TableLeft">&nbsp;</td>		<td width="20%" class="TableCenter">&nbsp;</td>					  <td width="20%" class="TableCenter">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>	    <td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		<td width="15%" class="TableLeft">&nbsp;</td>		<td width="20%" class="TableCenter">&nbsp;</td>					  <td width="20%" class="TableCenter">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>	    <td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		<td width="15%" class="TableLeft">&nbsp;</td>		<td width="20%" class="TableCenter">&nbsp;</td>					  <td width="20%" class="TableCenter">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>	    <td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		<td width="15%" class="TableBottomLeft">&nbsp;</td>	<td width="20%" class="TableBottomCenter">&nbsp;</td>			  <td width="20%" class="TableBottomCenter">&nbsp;</td><td width="30%" class="TableBottomCenter">&nbsp;</td><td class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 230px">
		<input id="btnAdd" value="Add" type="button" style="WIDTH: 60px" class="msgButton"> 
	</span>

	<span style="LEFT: 68px; POSITION: absolute; TOP: 230px">
		<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
	</span>

	<span style="LEFT: 132px; POSITION: absolute; TOP: 230px">
		<input id="btnDelete" value="Delete" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
</div> 

<div id="divThirdPartyType" style="HEIGHT: 40px; LEFT: 10px; POSITION: absolute; TOP: 336px; WIDTH: 604px; visibility: hidden" class="msgGroup">
	<span style="LEFT: 16px; POSITION: absolute; TOP: 12px" class="msgLabel">
		Please select the type of Third Party Agent to add:
		<span style="LEFT: 256px; POSITION: absolute; TOP: -2px">
			<select id="cboThirdPartyType" name="cboThirdPartyType" style="WIDTH: 176px" class="msgCombo">
				<option id="optSelect">&lt;SELECT&gt;</option>
				<option id="opt2">Architect</option>
				<option id="opt3">Bank/Building Society</option>
				<option id="opt4">Builder</option>
				<option id="opt6">Estate Agent</option>
				<option id="opt10">Legal Rep</option>
			</select>
		</span> 
	</span>

	<span style="LEFT: 464px; POSITION: absolute; TOP: 8px">
		<input id="btnContinue" value="Continue" type="button" style="WIDTH: 60px" class="msgButton">
	</span>

	<span style="LEFT: 528px; POSITION: absolute; TOP: 8px">
		<input id="btnUndo" value="Undo" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
</div>
</form>

<% /* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/DC240Attribs.asp" -->

<% /* CODE */ %>
<script language="JScript">
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_sTypeOfApplication = "";

var ThirdPartyXML = null;
var m_iTableLength = 10;
var scScreenFunctions;
var m_blnReadOnly = false;
var m_sProcessInd = ""; //JR BM0271
var m_sDataFreezeInd = ""; // BS BM0271 14/04/03


// EP2_8 - New variables
var lsBankAccountName = "";
var lsBankSortCode = "";
var lsBankAccountNo = "";
var lsBankName = "";
var lsBankDXID = "";
var lsBankDXLocation = "";
var lBankAccount = "";
var lsTPGUID = "";
var lsIsLegalReptoBeUsed = 1;  // Default to true.
var lbBankDetailsFound = 0;	   // Default to false.

/* EVENTS */

function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var bNewPropertySummary = XML.GetGlobalParameterBoolean(document,"NewPropertySummary");			
		if(bNewPropertySummary)
			frmToDC201.submit();				
		else
			frmToDC230.submit();
}

function btnSubmit.onclick()
{
	<% /* BM0193 MDC 16/12/2002 */ %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			if(scScreenFunctions.CompletenessCheckRouting(window))
				frmToGN300.submit();
			else
			{
				// EP2_8 - Additional Borrowing logic
				// Get the lsIsLegalReptoBeUsed flag.
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				XML.CreateRequestTag(window);
				var xn = XML.XMLDocument.documentElement;
				xn.setAttribute("CRUD_OP","READ");
				xn.setAttribute("SCHEMA_NAME","omCRUD");
				xn.setAttribute("ENTITY_REF","APPLICATION");
				var xe = XML.XMLDocument.createElement("APPLICATION");
				xe.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
				xn.appendChild(xe);
				XML.RunASP(document, "omCRUDIf.asp");
				if(XML.IsResponseOK())
				{
					XML.SelectTag(null,"APPLICATION");
					lsIsLegalReptoBeUsed = XML.GetTagInt("ISLEGALREPTOBEUSED");
				}
				
				// Set general variable values.
				var m_sidMortgageApplicationValue = scScreenFunctions.GetContextParameter(window,"MortgageApplicationValue","");
				XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

				// Get BankBuildingSoc list 
				XML.CreateRequestTag(window,null);
				var tagApplication = XML.CreateActiveTag("APPLICATIONBANKBUILDINGSOC");
				XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
				XML.CreateTag("APPLICATIONFACTFINDNUMBER", "1");
				
				XML.RunASP(document,"FindBankBuildingSocietyList.asp");
				lbBankDetailsFound = 0;
				var ErrorTypes = new Array("RECORDNOTFOUND");
				var ErrorReturn = XML.CheckResponse(ErrorTypes);
				if (ErrorReturn[1] != ErrorTypes[0])
				{	// Set variables if record found.
					lbBankDetailsFound = 1;
					// Now get Bank/BSoc Data (Stored in two tables).
					XML.SelectTag(null, "APPLICATIONBANKBUILDINGSOC");
					lsBankAccountName = XML.GetTagText("ACCOUNTNAME");  
					lsBankAccountNo = XML.GetTagText("ACCOUNTNUMBER");
					
					XML.SelectTag(null, "THIRDPARTY");
					lsBankSortCode = XML.GetTagText("THIRDPARTYBANKSORTCODE");
					lsBankName = XML.GetTagText("COMPANYNAME");
					lsBankDXID = XML.GetTagText("DXID");
					lsBankDXLocation = XML.GetTagText("DXLOCATION");
				}
				// Create our new records if required. 
				if(lsIsLegalReptoBeUsed == "0" 
				&& XML.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, Array("F")) 
				&& XML.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, Array("M"))
				&& lbBankDetailsFound == 1)
				{
					// Get the Payee Historylist
					XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
					XML.CreateRequestTag(window, "FindPayeeHistoryList");
					XML.CreateActiveTag("PAYEEHISTORY");
					XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
					XML.RunASP(document,"PaymentProcessingRequest.asp");
				
					var ErrorTypes = new Array("RECORDNOTFOUND");
					var ErrorReturn = XML.CheckResponse(ErrorTypes);
					
					if (ErrorReturn[1] == ErrorTypes[0]) // NO Payee history found
					{
						// Create Third Party record and return GUID.
						XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
						// Create Third Party record. 
						XML.CreateRequestTag(window,null);
						var tagApplication = XML.CreateActiveTag("THIRDPARTY");
						XML.CreateTag("THIRDPARTYTYPE", "13");
						XML.CreateTag("COMPANYNAME", lsBankAccountName);
						
						XML.RunASP(document,"CreateThirdParty.asp");
						var ErrorTypes = new Array("RECORDNOTFOUND");
						var ErrorReturn = XML.CheckResponse(ErrorTypes);
						if (ErrorReturn[1] != ErrorTypes[0]) // Get ThirdPartyGUID back 
						{	
							XML.SelectTag(null, "GENERATEDKEYS");
							lsTPGUID = XML.GetTagText("THIRDPARTYGUID");

							// Create a Payee history.
				 			XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
							XML.CreateRequestTag(window);
							var xn = XML.XMLDocument.documentElement;
							xn.setAttribute("CRUD_OP","CREATE");
							xn.setAttribute("SCHEMA_NAME","omCRUD");
							xn.setAttribute("ENTITY_REF","PAYEEHISTORY");
							var xe = XML.XMLDocument.createElement("PAYEEHISTORY");
							xe.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
							xe.setAttribute("PAYEEHISTORYSEQNUMBER", "1");
							xe.setAttribute("BANKSORTCODE", lsBankSortCode);
							xe.setAttribute("ACCOUNTNUMBER", lsBankAccountNo);
							xe.setAttribute("BANKNAME", lsBankName);
							xe.setAttribute("USRID", scScreenFunctions.GetContextParameter(window,"UserID",""));
							xe.setAttribute("CREATIONDATETIME", scScreenFunctions.DateTimeToString(scScreenFunctions.GetAppServerDate(true)));
							xe.setAttribute("DXID", lsBankDXID);
							xe.setAttribute("DXLOCATION", lsBankDXLocation);
							xe.setAttribute("THIRDPARTYGUID", lsTPGUID);
							xn.appendChild(xe);
							XML.RunASP(document, "omCRUDIf.asp");


/*	New Code		XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
Not working 		XML.RunASP(document,"PaymentProcessingRequest.asp");
				
					var ErrorTypes = new Array("RECORDNOTFOUND");
					var ErrorReturn = XML.CheckResponse(ErrorTypes);

							XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
							XML.CreateRequestTag(window, "CreatePayeeHistoryDetails");
							XML.CreateActiveTag("PAYEEHISTORY");
							XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
							XML.SetAttribute("PAYEEHISTORYSEQNUMBER", "1");
							XML.SetAttribute("BANKSORTCODE", lsBankSortCode);
							XML.SetAttribute("ACCOUNTNUMBER", lsBankAccountNo);
							XML.SetAttribute("BANKNAME", lsBankName);
							XML.SetAttribute("USRID", scScreenFunctions.GetContextParameter(window,"UserID",""));
							XML.SetAttribute("CREATIONDATETIME", scScreenFunctions.DateTimeToString(scScreenFunctions.GetAppServerDate(true)));
							XML.SetAttribute("DXID", lsBankDXID);
							XML.SetAttribute("DXLOCATION", lsBankDXLocation);
							XML.SetAttribute("THIRDPARTYGUID", lsTPGUID);
							XML.RunASP(document,"PaymentProcessingRequest.asp");
End New code */

						} // Got TPGUID back.
						
					} // Payee list not in error.

				}	// New Records required.		  
				
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				var bNewPropertySummary = XML.GetGlobalParameterBoolean(document,"NewPropertySummary");			
				if(bNewPropertySummary)
					frmToCM010.submit();
				else
					frmToDC295.submit();
				
				break;
			}
		default: // Error
			// Do not route to another screen
	}
	<% /* BM0193 MDC 16/12/2002  - End */ %>

}

function frmScreen.btnAdd.onclick()
{
	frmScreen.cboThirdPartyType.selectedIndex = 0;
	//SYS1522 Need onchange method incase keyboard used 
	//frmScreen.cboThirdPartyType.onclick();
	frmScreen.cboThirdPartyType.onchange();
	
	scScreenFunctions.ShowCollection(divThirdPartyType);
	SetFunctionButtonAccess();
	frmScreen.cboThirdPartyType.focus();
}

function frmScreen.btnContinue.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "Add");
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber", m_sApplicationNumber);
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber", m_sApplicationFactFindNumber);

	sThirdPartyType = frmScreen.cboThirdPartyType.value;

	if (sThirdPartyType == "4") frmToDC260.submit(); // Builder
	if (sThirdPartyType == "2") frmToDC290.submit(); // Architect
	if (sThirdPartyType == "6") frmToDC250.submit(); // Estate Agent
	if (sThirdPartyType == "3") frmToDC270.submit(); // Bank/Building Society
	// EP2_8 - Extra logic.
	if (sThirdPartyType == "10")					 // Legal Representative
	{
		// EP2_8 - Check whether Legal rep to be used
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window,null);
		XML.CreateActiveTag("APPLICATION");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.RunASP(document,"GetApplicationData.asp");
		if(XML.IsResponseOK())
		{
			XML.SelectTag(null,"APPLICATION");
			var lsIsLegalReptoBeUsed = XML.GetTagBoolean("ISLEGALREPTOBEUSED");
			<%/* MAH EP2_112 05/11/2006 Start */%>
			var m_sidMortgageApplicationValue = scScreenFunctions.GetContextParameter(window,"MortgageApplicationValue","");
			XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			<%/* MAH EP2_112 05/11/2006 End */%>
		}
		// Fire alert if NOT using a solicitor.
		if ((lsIsLegalReptoBeUsed == "0") 
			&& XML.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, Array("F")) <%/* MAH EP2_112 */%>
			&& XML.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, Array("M")))<%/* MAH EP2_112 */%>
		{
			var lAlertMesage = "You cannot add the Solicitor for this application because you have opted not to use a Solicitor.";
			alert(lAlertMesage);
			return;
		}
		else		
		frmToDC280.submit(); 
	}

	divThirdPartyType.style.visibility = "hidden";
	SetFunctionButtonAccess();
}

function frmScreen.btnDelete.onclick()
{
	var bAllowDelete = confirm("Are you sure?");

	if(bAllowDelete)
	{
		//Get the XML that just contains the dependant chosen in the listbox
		var XML = GetXMLBlock(false);
		var sThirdPartyType = XML.GetTagText("THIRDPARTYTYPE");

		// 		if (sThirdPartyType == "4") XML.RunASP(document,"DeleteBuilder.asp"); // Builder
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					if (sThirdPartyType == "4") XML.RunASP(document,"DeleteBuilder.asp"); // Builder
				break;
			default: // Error
				if (sThirdPartyType == "4") XML.SetErrorResponse();
			}

		// 		if (sThirdPartyType == "2") XML.RunASP(document,"DeleteArchitect.asp"); // Architect
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					if (sThirdPartyType == "2") XML.RunASP(document,"DeleteArchitect.asp"); // Architect
				break;
			default: // Error
				if (sThirdPartyType == "2") XML.SetErrorResponse();
			}

		// 		if (sThirdPartyType == "6") XML.RunASP(document,"DeleteEstateAgent.asp"); // Estate Agent
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					if (sThirdPartyType == "6") XML.RunASP(document,"DeleteEstateAgent.asp"); // Estate Agent
				break;
			default: // Error
				if (sThirdPartyType == "6") XML.SetErrorResponse();
			}

		// 		if (sThirdPartyType == "3") XML.RunASP(document,"DeleteBankBuildingSociety.asp"); // Bank/Building Society
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					if (sThirdPartyType == "3") XML.RunASP(document,"DeleteBankBuildingSociety.asp"); // Bank/Building Society
				break;
			default: // Error
				if (sThirdPartyType == "3") XML.SetErrorResponse();
			}

		// 		if (sThirdPartyType == "10") XML.RunASP(document,"DeleteLegalRep.asp"); // Legal Representative
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					if (sThirdPartyType == "10") XML.RunASP(document,"DeleteLegalRep.asp"); // Legal Representative
				break;
			default: // Error
				if (sThirdPartyType == "10") XML.SetErrorResponse();
			}


		if (XML.IsResponseOK() == true)
		{
			//Rather than re-search the database, remove the relevant section
			//from the ThirdPartyXML
			var nDeletedRow = DeleteItemFromXML();

			PopulateListBox(0);

			if(nDeletedRow > ThirdPartyXML.ActiveTagList.length)
				nDeletedRow = nDeletedRow - 1;
			if(nDeletedRow >=0)
				scScrollTable.setRowSelected(nDeletedRow);
			else
			{
				frmScreen.btnDelete.disabled = true;
				frmScreen.btnEdit.disabled = true;
			}
			if(frmScreen.btnAdd.disabled == false)
				frmScreen.btnAdd.focus();
		}
	}
	else
	{
		frmScreen.btnDelete.focus();
	}
}

function frmScreen.btnEdit.onclick()
{
	//Set idXML in context to the APPLICATIONTHIRDPARTY record for the selected list box line
	var XML = GetXMLBlock(true);
	var sThirdPartyType = XML.GetTagText("THIRDPARTYTYPE");
	if (sThirdPartyType == "") sThirdPartyType = XML.GetTagText("NAMEANDADDRESSTYPE");

	scScreenFunctions.SetContextParameter(window,"idXML", XML.ActiveTag.xml);
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "Edit");

	if (sThirdPartyType == "4") frmToDC260.submit(); // Builder
	if (sThirdPartyType == "2") frmToDC290.submit(); // Architect
	if (sThirdPartyType == "6") frmToDC250.submit(); // Estate Agent
	if (sThirdPartyType == "3") frmToDC270.submit(); // Bank/Building Society
	if (sThirdPartyType == "10") frmToDC280.submit(); // Legal Representative
}

function frmScreen.btnUndo.onclick()
{
	scScreenFunctions.HideCollection(divThirdPartyType);
	<% /* SetFunctionButtonAccess(); */ %>
	SetFunctionButtonAccessForTable();
	frmScreen.btnAdd.focus();
}
//SYS1522 Need onchange method incase keyboard used 
//function frmScreen.cboThirdPartyType.onclick()
function frmScreen.cboThirdPartyType.onchange()
{
	frmScreen.btnContinue.disabled = (frmScreen.cboThirdPartyType.selectedIndex < 1);
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Third Party Summary","DC240",scScreenFunctions);
	
	RetrieveContextData();
	
	<% /* BM0177 Also disable buttons if read only
	if (scScreenFunctions.GetContextParameter(window,"idFreezeDataIndicator",null) == 1)
	*/ %>
	<% /* BS BM0271 14/04/03
	if ((scScreenFunctions.GetContextParameter(window,"idFreezeDataIndicator",null) == 1) || (m_sReadOnly == "1") ||  */ %>
	if ((m_sDataFreezeInd == "1") || (m_sReadOnly == "1") || 
		(m_sProcessInd == "0")) //JR BM0271 Add ProcessInd check.		
	{
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnAdd.disabled = true;
	}
	else
	{
		frmScreen.btnDelete.disabled = false;
		frmScreen.btnAdd.disabled = false;
	}
	
	PopulateScreen();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();

	<% /* scScreenFunctions.SetFocusToFirstField(frmScreen); */ %>
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC240");
	
	if (m_blnReadOnly == true) 
	{
		m_sReadOnly = "1";
		<% /* KRW - 21/09/2004 - BMIDS833 */ %>
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnDelete.disabled = true;
	}
	
	scScreenFunctions.HideCollection(divThirdPartyType);
			
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

/* FUNCTIONS */

function AddToAddressString(sThisAddressLine, sCurrentAddress)
{
	if (sThisAddressLine != "")
	{
		if (sCurrentAddress == "") return sThisAddressLine;
		else
			return sCurrentAddress + ", " + sThisAddressLine;
	}
	else return sCurrentAddress;
}

function DeleteItemFromXML()
/* Deletes an APPLICATIONTHIRDPARTY block from the ThirdPartyXML
The APPLICATIONTHIRDPARTY block deleted is the one selected in the listbox */
{
	var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	ThirdPartyXML.ActiveTag = null;

	var tagNode = ThirdPartyXML.SelectTag(null, "APPLICATIONTHIRDPARTYLIST");
	tagNode.removeChild(tagNode.childNodes.item(nRowSelected -1));

	return(nRowSelected);
}

function GetXMLBlock(bForEdit)
/* Returns an XML DOM for a single APPLICATIONTHIRDPARTY block
from ThirdPartyXML - the one relating to the selected
line in the listbox */
{
	//From the row selected, find the index into the XML from which the table
	//was created. Create a new XML request block with just this block of XML
	var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	ThirdPartyXML.ActiveTag = null;
	//ThirdPartyXML.CreateTagList("APPLICATIONTHIRDPARTY");	
	ThirdPartyXML.SelectTagListItem(nRowSelected -1);

	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);

	if (bForEdit)
		XML.ActiveTag.appendChild(ThirdPartyXML.ActiveTag);
	else
	{
		var sTableName = ThirdPartyXML.GetAttribute("TABLE");
		var sThirdPartyType = ThirdPartyXML.GetTagText("THIRDPARTYTYPE");
		if (sThirdPartyType == "") sThirdPartyType = ThirdPartyXML.GetTagText("NAMEANDADDRESSTYPE");

		XML.CreateActiveTag(sTableName);
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER",	m_sApplicationFactFindNumber);
		XML.CreateTag("THIRDPARTYTYPE",	sThirdPartyType);
		if (sThirdPartyType == "3") // Bank/Building Society
			XML.CreateTag("BANKACCOUNTSEQUENCENUMBER", ThirdPartyXML.GetTagText("BANKACCOUNTSEQUENCENUMBER"));
	}
	return(XML);
}

function PopulateCombos()
// Populates all combos on the screen
{
	var blnSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("ThirdPartyType");

	if(XML.GetComboLists(document,sGroupList))
	{
		// Populate Country combo
		blnSuccess = XML.PopulateCombo(document,frmScreen.cboThirdPartyType,"ThirdPartyType",true);

		if(blnSuccess == false)
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}

	XML = null;
}

function PopulateListBox(nStart)
{
	ThirdPartyXML.ActiveTag = null;
	ThirdPartyXML.CreateTagList("APPLICATIONTHIRDPARTY");
	var iNumberOfThirdParties = ThirdPartyXML.ActiveTagList.length;

	PopulateThirdPartyTypes();
	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfThirdParties);
	ShowList(nStart);

	if(iNumberOfThirdParties > 0)
	{
		scScrollTable.setRowSelected(1);
		frmScreen.btnDelete.disabled = false;
		frmScreen.btnEdit.disabled = false;
	}
	else
	{
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnEdit.disabled = true;
	}
	frmScreen.btnAdd.disabled = (frmScreen.cboThirdPartyType.length == 1)
}

function PopulateScreen()
{
	// Retrieve main data
	ThirdPartyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ThirdPartyXML.CreateRequestTag(window,null);

	var tagApplication = ThirdPartyXML.CreateActiveTag("APPLICATION");
	ThirdPartyXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	ThirdPartyXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	ThirdPartyXML.RunASP(document,"FindApplicationThirdPartyList.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = ThirdPartyXML.CheckResponse(ErrorTypes);

	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		if(ErrorReturn[1] == ErrorTypes[0])
		{
			//record not found
			<% /* BS BM0271 14/04/03
			if(m_sReadOnly == "1") */ %>
			if ((m_sReadOnly == "1") || (m_sDataFreezeInd == "1") || (m_sProcessInd == "0"))
				frmScreen.btnAdd.disabled = true;
			else
				frmScreen.btnAdd.disabled = false;
			frmScreen.btnDelete.disabled = true;
			frmScreen.btnEdit.disabled = true;
		}
		else
		{
			PopulateListBox(0);
			<% /* BS BM0271 14/04/03
			if(m_sReadOnly == "1" || m_sProcessInd == "0") //JR BM0271 */ %>
			if ((m_sReadOnly == "1") || (m_sProcessInd == "0") || (m_sDataFreezeInd == "1"))
			{
				frmScreen.btnAdd.disabled = true;
				frmScreen.btnDelete.disabled = true;
			}

			
			if (!frmScreen.btnAdd.disabled)
				frmScreen.btnAdd.focus();
			else if(!frmScreen.btnEdit.disabled)
				frmScreen.btnEdit.focus();
		}
	}

	if (tblTable.rows > 0)
		scScrollTable.setRowSelected(1);

	ErrorTypes = null;
	ErrorReturn = null;
}

function PopulateThirdPartyTypes()
{
	var nBankBuildingSocietyCount = 0;
	var aThirdPartyTypes = new Array();
	PopulateCombos();

	<%/* Cannot add estate agents if application type is 'further advance' */%>
	var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	if (TempXML.IsInComboValidationList(document,"TypeOfMortgage",m_sTypeOfApplication,["F"]))
		aThirdPartyTypes[6] = "Y";			
	TempXML = null;

	// Mark those third party types in the list for deletion which already are represented
	// in the list of third party agents (bank/building societies can have up to 2 agents)
	for (iCount = 0; iCount < ThirdPartyXML.ActiveTagList.length; iCount++)
	{
		ThirdPartyXML.SelectTagListItem(iCount);
		var sThirdPartyType = ThirdPartyXML.GetTagText("THIRDPARTYTYPE");
		if (sThirdPartyType == "") sThirdPartyType = ThirdPartyXML.GetTagText("NAMEANDADDRESSTYPE");

		// Check whether the third party is a bank/building society
		if (sThirdPartyType == "3") nBankBuildingSocietyCount++;

		var optOption = null;

		// Remove the item corresponding to this third party type from the dropdown list
		if (((sThirdPartyType == "3") & (nBankBuildingSocietyCount > 1)) | (sThirdPartyType != "3"))
			aThirdPartyTypes[parseInt(sThirdPartyType)] = "Y";
	}

	// Remove all items marked for deletion (or of the wrong validation type) from the list
	optOption = null;
	<% /* SYS0632: Do not remove <Select> option */ %>
	for (iCount = frmScreen.cboThirdPartyType.length - 1; iCount > 0; iCount--)
	{
		optOption = frmScreen.cboThirdPartyType.item(iCount);

		if ((aThirdPartyTypes[parseInt(optOption.value)] == "Y") |
			!(scScreenFunctions.IsOptionValidationType(frmScreen.cboThirdPartyType, iCount, "T")))
			frmScreen.cboThirdPartyType.remove(iCount);
	}
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","Edit");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","B00003743");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");  
	m_sTypeOfApplication = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", "10");
	m_sProcessInd = scScreenFunctions.GetContextParameter(window,"idProcessingIndicator",null); //JR BM0271
	m_sDataFreezeInd = scScreenFunctions.GetContextParameter(window,"idFreezeDataIndicator",null); //BS BM0271 14/04/03
}

function SetFunctionButtonAccess()
{
	frmScreen.btnAdd.disabled = (divThirdPartyType.style.visibility == "visible");
	frmScreen.btnDelete.disabled = (divThirdPartyType.style.visibility == "visible");
	frmScreen.btnEdit.disabled = (divThirdPartyType.style.visibility == "visible");
}

function SetFunctionButtonAccessForTable()
{
	frmScreen.btnAdd.disabled = false;
	if (scScrollTable.getRowSelected() >= 0)
	{
		frmScreen.btnDelete.disabled = false;
		frmScreen.btnEdit.disabled = false;
	}
	else
	{
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnEdit.disabled = true;
	}
}

function ShowList(nStart)
{
	var iCount;
	scScrollTable.clear();		
	for (iCount = 0; iCount < ThirdPartyXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		ThirdPartyXML.SelectTagListItem(iCount + nStart);

		var sCustomerName = scScreenFunctions.GetContextCustomerName(window, ThirdPartyXML.GetTagText("CUSTOMERNUMBER"));
		var sThirdPartyType = ThirdPartyXML.GetTagAttribute("THIRDPARTYTYPE", "TEXT");
		if (sThirdPartyType == "") sThirdPartyType = ThirdPartyXML.GetTagAttribute("NAMEANDADDRESSTYPE", "TEXT");
		var sACNumber = ThirdPartyXML.GetTagText("ACCOUNTNUMBER");
		var sTemp = ThirdPartyXML.GetTagText("REPAYMENTBANKACCOUNTINDICATOR");
		var sIsRepaymentBankAccount = "N/A";
		if (sTemp != "") sIsRepaymentBankAccount = (sTemp == "1") ? "Yes" : "No";
		var sCompanyName = ThirdPartyXML.GetTagText("COMPANYNAME");

		// Contact
		var sContactName = ThirdPartyXML.GetTagText("CONTACTFORENAME")  + ' ' + 
						   ThirdPartyXML.GetTagText("CONTACTSURNAME");
		// Address
		var sAddress = "";
		sAddress = AddToAddressString(ThirdPartyXML.GetTagText("POSTCODE"), sAddress);
		sAddress = AddToAddressString(ThirdPartyXML.GetTagText("FLATNUMBER"), sAddress);
		sAddress = AddToAddressString(ThirdPartyXML.GetTagText("BUILDINGORHOUSENAME"), sAddress);
		sAddress = AddToAddressString(ThirdPartyXML.GetTagText("BUILDINGORHOUSENUMBER"), sAddress);
		sAddress = AddToAddressString(ThirdPartyXML.GetTagText("STREET"), sAddress);
		sAddress = AddToAddressString(ThirdPartyXML.GetTagText("DISTRICT"), sAddress);
		sAddress = AddToAddressString(ThirdPartyXML.GetTagText("TOWN"), sAddress);
		sAddress = AddToAddressString(ThirdPartyXML.GetTagText("COUNTY"), sAddress);
		sAddress = AddToAddressString(ThirdPartyXML.GetTagAttribute("COUNTRY","TEXT"), sAddress);

		// Add to the search table
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0), sThirdPartyType);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),
			(sACNumber != "" ? sACNumber : sContactName));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),sCompanyName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3),sAddress);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4),sIsRepaymentBankAccount);
		tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount + nStart);
	}
}


</script>
</body>
</html>




