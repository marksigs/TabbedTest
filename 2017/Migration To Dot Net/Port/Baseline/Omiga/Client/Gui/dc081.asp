<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      dc081.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Add/Edit Bank/Credit Cards
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
APS		02/02/2000	ie5 changes 
AY		14/02/00	Change to msgButtons button types
IW		10/02/00	SYS0119 - Misc cosmetic changes
AY		30/03/00	New top menu/scScreenFunctions change
MH		03/05/00	SYS0492 Allow collection of guarantor details
IW		10/05/00	SYS0704 Misc Validation Changes
BG		17/05/00	SYS0752 Removed Tooltips
MC		06/06/00	SYS0837 Standardise approach to restricting length
					of text in textarea fields.
MC		09/06/00	SYS0880 Increase width of Card Type combo
MC		09/06/00	SYS0866 Standardise format of balance and payment fields
MH		22/06/00	SYS0933 Standardise format of balance and payment fields
CL		05/03/01	SYS1920 Read only functionality added
LD		23/05/02	SYS4727 Use cached versions of frame functions
HMA     17/09/03    BM0063  Amend HTML text for radio buttons
SR		01/06/2004	BMIDS772 Update FinancialSummary record on Submit (only for create)	

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS specific History:

Prog	Date		Description
SD		14/11/2005	MAR258 Critical Data check changes
DRC     03/02/2006  MAR1189 endable disable all fileds dep on setting of glob param  FSDisableBankCredCardAddButton
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM specific History:

Prog	Date		Description
IK		12/04/2006	EP353 - fix error re: btnAdd
IK		03/05/2006	EP495 - incorporate MAR1306
SAB		15/05/2006	EP519 - Return to DC080 when the user has clicked OK
IK		15/08/2006	EP1085 - save details occasional failure
DS		19/03/2007	EP2_1814 - In edit mode, disabled Bank/Credit card owner
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript"></script>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>

<form id="frmToDC080" method="post" action="dc080.asp" STYLE="DISPLAY: none"></form>

<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 250px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Bank/Credit Card Owner
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<select id="cboApplicant" name="Applicant"  style="WIDTH: 200px" class="msgCombo">
			</select>
		</span>
	</span>
	<span style="TOP: 36px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Card Type
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<select id="cboCardType" name="CardType" style="WIDTH: 120px" class="msgCombo">
			</select>
		</span>
	</span>
	<span style="TOP: 62px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Card Provider
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="txtCardProvider" name="CardProvider" maxlength="50" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 88px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Total Outstanding Balance
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="txtTotalOutstandingBalance" name="TotalOutstandingBalance" maxlength="6" style="WIDTH: 65px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 114px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Average Monthly Repayment
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="txtAverageMonthlyRepayment" name="AverageMonthlyRepayment" maxlength="9" style="WIDTH: 65px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
	<span style="TOP: 140px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Are there any additional Card Holders?
		<span style="TOP: -3px; LEFT: 196px; POSITION: ABSOLUTE">
			<input id="optAdditionalIndYes" name="AdditionalInd" type="radio" value="1"><label for="optAdditionalIndYes" class="msgLabel">Yes</label>
		</span> 
		<span style="TOP: -3px; LEFT: 250px; POSITION: ABSOLUTE">
			<input id="optAdditionalIndNo" name="AdditionalInd" type="radio" value="0" checked><label for="optAdditionalIndNo" class="msgLabel">No</label>
		</span> 
	</span>		
	<span style="TOP: 166px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Additional Card Holder(s) Details
		<span style="TOP: 0px; LEFT: 200px; POSITION: ABSOLUTE">
			<textarea id="txtAdditionalDetails" name="AdditionalDetails" rows="5" style="WIDTH: 300px; POSITION: absolute" class="msgTxt"></textarea>
		</span>
	</span>
</div>
</form>

<!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/dc081attribs.asp" -->

<script language="JScript">
<!--		
var m_sMetaAction = null;		
var m_sXMLOnEntry = null;
var m_sSequenceNumber = null;		
var scScreenFunctions;
var m_blnReadOnly = false;
<% /* SR 10/06/2004 : BMIDS772 */ %>
var m_sApplicationNumber			= null;
var m_sApplicationFactFindNumber	= null;
<% /* SR 10/06/2004 : BMIDS772 - End */ %>
<% /* MAR1306 / EP495 */ %>
var m_isBtnSubmit = false;

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	RetrieveContextData();

	<% /* Make the required buttons available on the bottom of the screen
		(see msgButtons.asp for details) */ %>
	var sButtonList = new Array("Submit","Cancel","Another");

	<% /* If not in add mode then the another button is not required */ %>
	if(m_sMetaAction != "Add")
	{
		sButtonList = new Array("Submit","Cancel");
	}

	ShowMainButtons(sButtonList);
	FW030SetTitles("Add/Edit Bank/Credit Cards","DC081",scScreenFunctions);

	GetComboLists();
	PopulateScreen();
			
	frmScreen.cboApplicant.onchange();
	frmScreen.optAdditionalIndNo.onclick();
						
	SetMasks();

	Validation_Init();
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");

	if (m_sReadOnly=="1") scScreenFunctions.SetScreenToReadOnly(frmScreen);

	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC081");
	//DRC MAR1189 - check if ADD button should be disabled
	var GlobXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bDisableCredCardCapture = GlobXML.GetGlobalParameterBoolean(document,"FSDisableBankCredCardAddButton");				
	if (bDisableCredCardCapture)
	  scScreenFunctions.SetScreenToReadOnly(frmScreen);
// ik_EP353_12/04/2006
//	else
//		frmScreen.btnAdd.disabled=false; 

		<% /*MAR1306 / EP495  */ %>
		btnAnother.disabled=false; 
		btnSubmit.disabled=false;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function frmScreen.txtAdditionalDetails.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalDetails", 255, true);
}


<% /* keep the focus within this screen when using the tab key */ %>
function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

<% /* keep the focus within this screen when using the tab key */ %>
function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

<% /* Retrieves all context data required for use within this screen */ %>
function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	<% /* SR 10/06/2004 : BMIDS772 */ %>
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber");
	<% /* SR 10/06/2004 : BMIDS772 - End */ %>
}

<% /* Populates all combos with their options */ %>
function GetComboLists()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();			
	var sGroups = new Array("CreditCardType");
						
	if (XML.GetComboLists(document, sGroups) == true)
	{
		XML.PopulateCombo(document, frmScreen.cboCardType, "CreditCardType", true);
	}
	PopulateApplicantCombo();
	XML = null;
}

<% /* Populates the Applicant combo with all applicants	currently held in context*/ %>
function PopulateApplicantCombo()
{
	<% /* Clear any <OPTION> elements from the combo */ %>
	while(frmScreen.cboApplicant.options.length > 0)
	{
		frmScreen.cboApplicant.options.remove(0);
	}

	<% /* Add a <SELECT> option */ %>
	var TagOPTION	= document.createElement("OPTION");
	TagOPTION.value	= "";
	TagOPTION.text	= "<SELECT>";
	frmScreen.cboApplicant.add(TagOPTION);

	var nCustomerCount = 0;

	<% /* Loop through all customer context entries */ %>
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerName = scScreenFunctions.GetContextParameter(window,"idCustomerName" + nLoop);
		var sCustomerNumber	= scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

		<% /* If the customer is an applicant, add him/her as an option */ %>
		<% /* SYS0492 - or guarantor */ %>
		if(sCustomerRoleType == "1" || sCustomerRoleType == "2")
		{
			TagOPTION = document.createElement("OPTION");
			TagOPTION.value	= sCustomerNumber;
			TagOPTION.text = sCustomerName;
			TagOPTION.setAttribute("CustomerVersionNumber", sCustomerVersionNumber);
			frmScreen.cboApplicant.add(TagOPTION);
			nCustomerCount++;
		}
	}
		
	if(nCustomerCount == 1)
	{
		frmScreen.cboApplicant.remove(0);
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboApplicant");
	}

	<% /* Default to the first option */ %>
	frmScreen.cboApplicant.selectedIndex = 0;
}
		
function frmScreen.cboApplicant.onchange()
{
	if(frmScreen.cboApplicant.value != "")
	{
		if(frmScreen.cboApplicant.item(0).text == "<SELECT>")
		{
			frmScreen.cboApplicant.remove(0);
		}
	}
}

<% /* Retrieves the data and sets the screen accordingly */ %>
function PopulateScreen()
{
	if(m_sMetaAction == "Edit")
	{
		//DS - EP2_1814
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboApplicant");
		m_sXMLOnEntry = scScreenFunctions.GetContextParameter(window,"idXML",null);
		var XML	= new top.frames[1].document.all.scXMLFunctions.XMLObject();
				
		XML.LoadXML(m_sXMLOnEntry);		
				
		if(XML.SelectTag(null,"BANKCREDITCARD") != null)
		{
			var sCustomerNumber = XML.GetTagText("CUSTOMERNUMBER");

			frmScreen.cboApplicant.value = sCustomerNumber;
			frmScreen.cboCardType.value	= XML.GetTagText("CARDTYPE");
			frmScreen.txtCardProvider.value	= XML.GetTagText("CARDPROVIDER");
			frmScreen.txtTotalOutstandingBalance.value = XML.GetTagText("TOTALOUTSTANDINGBALANCE");
			frmScreen.txtAverageMonthlyRepayment.value = XML.GetTagText("AVERAGEMONTHLYREPAYMENT");
			frmScreen.txtAdditionalDetails.value = XML.GetTagText("ADDITIONALDETAILS");
					
			scScreenFunctions.SetRadioGroupValue(frmScreen, "AdditionalInd", XML.GetTagText("ADDITIONALINDICATOR"));
			m_sSequenceNumber = XML.GetTagText("SEQUENCENUMBER");
		}
	}
}

<% /* Sets the Additional Card Holder field to writable */ %>
function frmScreen.optAdditionalIndYes.onclick()
{
	if(frmScreen.optAdditionalIndYes.checked == true)
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtAdditionalDetails");
	}
}

<% /* Sets the Additional Card Holder field to read only */ %>
function frmScreen.optAdditionalIndNo.onclick()
{
	if(frmScreen.optAdditionalIndNo.checked == true)
	{
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtAdditionalDetails");
	}
}

function btnSubmit.onclick()
{
	<% /* MAR1306 / EP495 Disable the button and display an hourglass until completenesscheck finishes */ %>
	btnSubmit.disabled = true;
	btnSubmit.blur();
	btnSubmit.style.cursor = "wait";
	btnAnother.disabled = true;
	btnAnother.blur();
	btnAnother.style.cursor = "wait";
	m_isBtnSubmit = true;

	if (frmScreen.onsubmit() == true)
	{
		<% /* MAR1306 Moved processing from here to CommitScreen & finishProcessing. */ %>
		<% /* Call CommitScreen after a timeout to allow the cursor time to change. */ %>
		window.setTimeout("CommitScreen()", 0)
	}
	else
	{
		btnAnother.style.cursor = "hand";
		btnAnother.disabled = false;
		btnSubmit.style.cursor = "hand";
		btnSubmit.disabled = false;

		m_isBtnSubmit = false;
	}

	<% /* EP519 - Return to DC080 */%>
	<% /* EP1085 - move to  CommitScreen() 
	if (m_isBtnSubmit == true)
	{
		frmToDC080.submit();
	}
	*/%>
}
		
function btnCancel.onclick()
{
	frmToDC080.submit();
}

<% /* Event handler for the Another frame button Saves the record and 
	clears all fields for new input */ %>
function btnAnother.onclick()
{
	if (frmScreen.onsubmit() == true)
	{
		if(CommitScreen() == true)
		{
			PopulateApplicantCombo();
			scScreenFunctions.ClearCollection(frmScreen);
			frmScreen.optAdditionalIndNo.checked = true;
			frmScreen.optAdditionalIndNo.onclick();
			scScreenFunctions.SetFocusToFirstField(frmScreen);
		}
	}
}

<% /* Commits the screen data to the database, either by a create or update */ %>
function CommitScreen()
{
	var bOK = false;			
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var TagRequestType = null;
			
	if (m_sReadOnly=="1") return true;
	
	if(scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalDetails", 255, true))
		return false;
						
	<% /* <REQUEST USERID=? USERTYPE=? UNIT=?> */ %>
	/*if(m_sMetaAction == "Add")
	{
		TagRequestType = XML.CreateRequestTag(window, "CREATE");
	}
	else
	{
		TagRequestType = XML.CreateRequestTag(window, "UPDATE");
	}*/

	TagRequestType = XML.CreateRequestTag(window, null);
	XML.SetAttribute("OPERATION","CriticalDataCheck");
	
	
	if(TagRequestType != null)
	{
		XML.CreateActiveTag("BANKCREDITCARD");
				
		var strCustomerNumber = frmScreen.cboApplicant.value;				
		XML.CreateTag("CUSTOMERNUMBER", strCustomerNumber);

		var nSelectedCustomer = frmScreen.cboApplicant.selectedIndex;
		var TagOption = frmScreen.cboApplicant.options.item(nSelectedCustomer);
								
		XML.CreateTag("CUSTOMERVERSIONNUMBER",TagOption.getAttribute("CustomerVersionNumber"));
				
		<% /* For an update we need to specify the sequence number */ %>
		if(m_sMetaAction == "Edit")
		{
			XML.CreateTag("SEQUENCENUMBER",m_sSequenceNumber);								
		}

		XML.CreateTag("ADDITIONALDETAILS",frmScreen.txtAdditionalDetails.value);
		XML.CreateTag("ADDITIONALINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"AdditionalInd"));
		XML.CreateTag("AVERAGEMONTHLYREPAYMENT",frmScreen.txtAverageMonthlyRepayment.value);
		XML.CreateTag("CARDPROVIDER",frmScreen.txtCardProvider.value);
		XML.CreateTag("CARDTYPE",frmScreen.cboCardType.value);
		XML.CreateTag("TOTALOUTSTANDINGBALANCE",frmScreen.txtTotalOutstandingBalance.value);
		<% /* SR 09/06/2004 : BMIDS772 - If FinancialSummary(FS) record needs to be updated, pass the 
			 node to CustomerFinancialBO. 
		*/ %>
		XML.ActiveTag = TagRequestType
		XML.CreateActiveTag("FINANCIALSUMMARY");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("BANKCARDINDICATOR", 1);
		<% /* SR 09/06/2004 : BMIDS772 - End */ %>
		if(m_sMetaAction == "Add")
		{
			// 			XML.RunASP(document, "CreateBankCard.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			/*
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document, "CreateBankCard.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}
			*/
		}
		else
		{
			<% /* A PREVIOUS KEY section may need to be created because we may have 
				changed the customer to who the dank card applies*/ %>					
			var XMLOnEntry = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XMLOnEntry.LoadXML(m_sXMLOnEntry);
			
			if(XMLOnEntry.SelectTag(null,"BANKCREDITCARD") != null)
			{
				var strOrigCustomerNumber = XMLOnEntry.GetTagText("CUSTOMERNUMBER");					
				if (strOrigCustomerNumber != strCustomerNumber)
				{						
					XML.CreateActiveTag("PREVIOUSKEY");
					XML.CreateActiveTag("BANKCREDITCARD");
					XML.CreateTag("CUSTOMERNUMBER", strOrigCustomerNumber);
					XML.CreateTag("CUSTOMERVERSIONNUMBER", XMLOnEntry.GetTagText("CUSTOMERVERSIONNUMBER"));
					XML.CreateTag("SEQUENCENUMBER", XMLOnEntry.GetTagText("SEQUENCENUMBER"));
				}					
			}
			XMLOnEntry = null;
		}
		
		XML.SelectTag(null,"REQUEST");
		XML.CreateActiveTag("CRITICALDATACONTEXT");
		XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		XML.SetAttribute("SOURCEAPPLICATION","Omiga");
		XML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
		XML.SetAttribute("ACTIVITYINSTANCE","1");
		XML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
		XML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
		XML.SetAttribute("COMPONENT","omCF.CustomerFinancialBO");
		
		if(m_sMetaAction == "Add")
			XML.SetAttribute("METHOD","CreateBankCard");	
		else
			XML.SetAttribute("METHOD","UpdateBankCard");	
	
					
		window.status = "Critical Data Check - please wait";
		
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
						XML.RunASP(document, "OmigaTMBO.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

		window.status = "";
			
		bOK = XML.IsResponseOK();
	}
	else
	{
		alert("CommitScreen - Invalid MetaAction");
	}

	<% /* EP1085  */%>
	if(m_isBtnSubmit)
	{
		if(bOK) frmToDC080.submit();
		else
		{
			btnAnother.style.cursor = "hand";
			btnAnother.disabled = false;
			btnSubmit.style.cursor = "hand";
			btnSubmit.disabled = false;
			m_isBtnSubmit = false;
		}
	}
	
	return bOK;
}
-->
</script>
</body>
</html>


