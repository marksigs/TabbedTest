<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /* 
Workfile:      DC181.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Employed Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		10/02/2000	Created
IW		22/03/2000  SYS0527
IW		24/03/2000  Homezone Mods: AboveFivePercentSharesOwnedIndicator
AY		30/03/00	New top menu/scScreenFunctions change
IVW		07/04/2000	Corrected percentage validation
IVW		11/04/2000	Changed Prev/Next/Cancel to Ok/Cancel
BG		17/05/00	SYS0752 Removed Tooltips
MC		23/05/00	SYS0756 If Read-only mode, disabled fields
CL		05/03/01	SYS1920 Read only functionality added
DJP		31/01/02	SYS2564(parent) SYS3806(client) Cosmetic changes
DPF		21/06/02	BMIDS00077 - File Changed to be in line with Core V7.0.2
					SYS3119 Hide shares owned percentage when less than 5% button selected
					SYS4727 Use cached versions of frame functions
					
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		15/05/2002	BMIDS00008	Added a New Button Accountants,Added a New function frmScreen.btnAccountants.onclick()
								Modified DefaultFields(),Added frmScreen.optEmploymentRelationshipYes.onclick()
								frmScreen.optEmploymentRelationshipNo.onclick()
MV		17/05/2002	BMIDS00008	Modified SaveEmployedDetails(); added AccountantGUID to the Request								
MV		24/05/2002	BMIDS00008	Modified Initialise(),PopulateScreen(),DefaultFields(),Added a function SetButtons()
MV		10/06/2002	BMIDS00008	Modified PopulateScreen(),SetButtons()
DPF		05/07/2002  BMIDS00091	Have added line to Default Fields to diasble Accountants button.
MV		22/08/2002	BMIDS00355	IE 5.5 Upgrade - Modified the Msgbuttons Position 
ASu		04/09/2002	BMIDS00305	Accountant button not disabled on initialise in Edit mode when required.
HMA     17/09/2003  BM0063      Amend HTML text for radio buttons
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
<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>

<% /* removed - DPF 21/06/02 - BMIDS00077
<OBJECT data="scScreenFunctions.asp" height=1 id=scScrFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<object data="scMathFunctions.asp" height="1px" id="scMathFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<OBJECT data="scXMLFunctions.asp" height=1 id=scXMLFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
*/ %>

<% /* FORMS */ %>
<form id="frmToDC160" method="post" action="DC160.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC170" method="post" action="DC170.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC190" method="post" action="DC190.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange" year4>
<div style="HEIGHT: 230px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Employer's Name
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="txtEmployerName" maxlength="50" style="POSITION: absolute; WIDTH: 200px" class="msgReadOnly" readonly tabindex=-1>
		</span> 
	</span>

	<span style="TOP: 34px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		<LABEL id=idPayrollNumber></LABEL>
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="txtPayrollNumber" maxlength="50" style="POSITION: absolute; WIDTH: 200px" class="msgTxt">
		</span> 
	</span>
	<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
	<span id=idWageslipsSeen style="TOP: 58px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Last 3 Wage Slips Seen?
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="optWageSlipsSeenYes" name="WageSlipsSeenGroup" type="radio" value="1"><label for="optWageSlipsSeenYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 274px; POSITION: absolute; TOP: -3px">
			<input id="optWageSlipsSeenNo" name="WageSlipsSeenGroup" type="radio" value="0"><label for="optWageSlipsSeenNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="TOP: 82px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		<LABEL id=idP60Seen></LABEL>
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="optP60SeenYes" name="P60SeenGroup" type="radio" value="1"><label for="optP60SeenYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 274px; POSITION: absolute; TOP: -3px">
			<input id="optP60SeenNo" name="P60SeenGroup" type="radio" value="0"><label for="optP60SeenNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="TOP: 106px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Any Notice of Termination or Redundancy?
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="optNoticeProblemYes" name="NoticeProblemGroup" type="radio" value="1"><label for="optNoticeProblemYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 274px; POSITION: absolute; TOP: -3px">
			<input id="optNoticeProblemNo" name="NoticeProblemGroup" type="radio" value="0"><label for="optNoticeProblemNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="TOP: 130px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Own 5% or more of shares in Company?
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="optSharesOwnedIndicatorYes" name="SharesOwnedIndicatorGroup" type="radio" value="1"><label for="opSharesOwnedIndicatorYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 274px; POSITION: absolute; TOP: -3px">
			<input id="optSharesOwnedIndicatorNo" name="SharesOwnedIndicatorGroup" type="radio" value="0" checked=true><label for="optSharesOwnedIndicatorNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="TOP: 154px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Percentage of Shares held
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="txtPercentSharesHeld" maxlength="3" style="POSITION: absolute; WIDTH: 40px" class="msgTxt">
			<span style="LEFT: 40px; TOP: 3px; POSITION: absolute" class="msgLabel">%</span>
		</span> 
	</span>

	<span style="TOP: 178px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Employment subject to Probationary Period?
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="optProbationaryYes" name="ProbationaryGroup" type="radio" value="1"><label for="optProbationaryYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 274px; POSITION: absolute; TOP: -3px">
			<input id="optProbationaryNo" name="ProbationaryGroup" type="radio" value="0"><label for="optProbationaryNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="TOP: 202px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Are you related to your Employer?
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="optEmploymentRelationshipYes" name="EmploymentRelationshipGroup" type="radio" value="1"><label for="optEmploymentRelationshipYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 274px; POSITION: absolute; TOP: -3px">
			<input id="optEmploymentRelationshipNo" name="EmploymentRelationshipGroup" type="radio" value="0"><label for="optEmploymentRelationshipNo" class="msgLabel">No</label> 
		</span> 
		
		<span style="LEFT: 350px; POSITION: absolute; TOP: -3px" class="msgLabel">
			<input id="btnAccountants" value="Accountants" type="button" style="WIDTH: 94px" class="msgButton">
		</span>
		
	</span>
</div>
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 310px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/DC181attribs.asp" -->
<!-- #include FILE="customise/DC181customise.asp" -->
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sCustomerNumber = "";
var m_sCustomerVersionNumber = "";
var m_sEmploymentSequenceNumber = "";
var m_sEmployerName = "";
var EmployedDetailsXML = null;
var scScreenFunctions;
var m_blnReadOnly = false;
var m_bCanAddToDirectory = false;
var m_sAccountantGUID = "";
var m_dblPercentSharesHeld = 0;

<% /* EVENTS */%>

function btnCancel.onclick()
{
	<% /* clear the contexts */ %>
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "Edit");
	frmToDC170.submit();		
}

function btnSubmit.onclick()
{
	if (CommitChanges())
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
		frmToDC190.submit();
	}
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
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	//next line replaced with line below - DPF 21/06/02 - BMIDS00077
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Other Details for Employed Period","DC181",scScreenFunctions);

	RetrieveContextData();
	SetMasks();

	Validation_Init();	
	Initialise(true);

	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC181");	
	if (m_blnReadOnly == true) m_sReadOnly = "1";
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);

	Customise();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function CommitChanges()
{
	var bSuccess = true;

	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (frmScreen.txtPercentSharesHeld.value > 100)
			{
				alert("A valid percentage must be specified (0-100).");
				frmScreen.txtPercentSharesHeld.focus();
				bSuccess = false;
			}
			else
				bSuccess = SaveEmployedDetails();
		}
	return(bSuccess);
}

<% /*  Inserts default values into all fields */ %>
function DefaultFields()
{
	frmScreen.txtPayrollNumber.value = "";
	frmScreen.txtPercentSharesHeld.value = "";
	scScreenFunctions.SetRadioGroupValue(frmScreen,"WageSlipsSeenGroup","0");
	scScreenFunctions.SetRadioGroupValue(frmScreen,"NoticeProblemGroup","0");
	scScreenFunctions.SetRadioGroupValue(frmScreen,"SharesOwnedIndicatorGroup","0");
	scScreenFunctions.SetRadioGroupValue(frmScreen,"P60SeenGroup","0");
	scScreenFunctions.SetRadioGroupValue(frmScreen,"ProbationaryGroup","0");
	scScreenFunctions.SetRadioGroupValue(frmScreen,"EmploymentRelationshipGroup","0");
	
	//line added - DPF 21/06/02 - BMIDS00077
	frmScreen.txtPercentSharesHeld.style.visibility = "hidden";
	//BMIDS0091 - DPF 05/07/02 - line added to disable Accountants button until it is activated by user
	frmScreen.btnAccountants.disabled = true;
}

<% /*  Initialises the screen */ %>
function Initialise(bOnLoad)
{
	frmScreen.txtEmployerName.value = m_sEmployerName;
	
	if (!PopulateScreen())
	{
		DefaultFields();
		SetButtons();
	}

	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

<% /*  Populates the screen with details of the item selected in dc160 */ %>
function PopulateScreen()
{
	//next line replaced with line below - DPF 21/06/02 - BMIDS0077
	//EmployedDetailsXML = new scXMLFunctions.XMLObject();
	EmployedDetailsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	EmployedDetailsXML.CreateRequestTag(window,null);
	EmployedDetailsXML.CreateActiveTag("EMPLOYEDDETAILS");
	EmployedDetailsXML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	EmployedDetailsXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	EmployedDetailsXML.CreateTag("EMPLOYMENTSEQUENCENUMBER", m_sEmploymentSequenceNumber);
	EmployedDetailsXML.RunASP(document,"GetEmployedDetails.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = EmployedDetailsXML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[1] == ErrorTypes[0]))
	{
		m_sMetaAction = "Add";
	}
	else if (ErrorReturn[0] == true)
	{
		m_sMetaAction = "Edit";

		if(EmployedDetailsXML.SelectTag(null, "EMPLOYEDDETAILS") != null)
		{
			frmScreen.txtPayrollNumber.value = EmployedDetailsXML.GetTagText("PAYROLLNUMBER");
			frmScreen.txtPercentSharesHeld.value = EmployedDetailsXML.GetTagText("PERCENTSHARESHELD");
			scScreenFunctions.SetRadioGroupValue(frmScreen,"WageSlipsSeenGroup",EmployedDetailsXML.GetTagText("WAGESLIPSSEENINDICATOR"));
			scScreenFunctions.SetRadioGroupValue(frmScreen,"NoticeProblemGroup",EmployedDetailsXML.GetTagText("NOTICEPROBLEMINDICATOR"));
			scScreenFunctions.SetRadioGroupValue(frmScreen,"SharesOwnedIndicatorGroup",EmployedDetailsXML.GetTagText("SHARESOWNEDINDICATOR"));
			//lines added - DPF 21/06/02 - BMIDS00077
			m_dblPercentSharesHeld = EmployedDetailsXML.GetTagText("PERCENTSHARESHELD");
			//
			// Show / Hide the Shares held amount
			//
			if (EmployedDetailsXML.GetTagText("SHARESOWNEDINDICATOR") == "1")
			{
				frmScreen.optSharesOwnedIndicatorYes.onclick();
			}
			else
			{
				frmScreen.optSharesOwnedIndicatorNo.onclick();
			}
			//BMIDS00077
			scScreenFunctions.SetRadioGroupValue(frmScreen,"P60SeenGroup",EmployedDetailsXML.GetTagText("P60SEENINDICATOR"));
			scScreenFunctions.SetRadioGroupValue(frmScreen,"ProbationaryGroup",EmployedDetailsXML.GetTagText("PROBATIONARYINDICATOR"));
			scScreenFunctions.SetRadioGroupValue(frmScreen,"EmploymentRelationshipGroup",EmployedDetailsXML.GetTagText("EMPLOYMENTRELATIONSHIPIND"));
			SetButtons();
			
		}
	}

	ErrorTypes = null;
	ErrorReturn = null;

	return(m_sMetaAction == "Edit");
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","Edit");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber","1325");
	m_sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber","1");
	m_sEmploymentSequenceNumber = scScreenFunctions.GetContextParameter(window,"idEmploymentSequenceNumber","1");    
	m_sEmployerName = scScreenFunctions.GetContextParameter(window,"idEmployerName","");    
	var sUserRole = scScreenFunctions.GetContextParameter(window,"idRole","0");
	//next line replaced with line below - DPF 25/06/02 - BMIDS00077
	//var XML = new scXMLFunctions.XMLObject();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_bCanAddToDirectory = (sUserRole >= XML.GetGlobalParameterAmount(document,"AddToDirectoryRole"));

}

function SetButtons()
{
	if (EmployedDetailsXML != null) 
	{
		if (EmployedDetailsXML.GetTagText("EMPLOYMENTRELATIONSHIPIND") == "1" )
		{
			frmScreen.btnAccountants.disabled = false;
			m_sAccountantGUID = EmployedDetailsXML.GetTagText("ACCOUNTANTGUID");
		}
		 // ASu 04/09/02 BMIDS00305 - Check for 0. Start 
		else
		{
			frmScreen.btnAccountants.disabled = true;
			m_sAccountantGUID = EmployedDetailsXML.GetTagText("ACCOUNTANTGUID");
		}
		// ASu End.		
	}
	else
		frmScreen.btnAccountants.disabled = true;
		
}

function SaveEmployedDetails()
{
	var bSuccess = true;
	//next line replaced by line below - DPF 21/06/02 - BMIDS00077
	//var XML = new scXMLFunctions.XMLObject();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateRequestTag(window,null)

	XML.CreateActiveTag("EMPLOYEDDETAILS");
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	XML.CreateTag("EMPLOYMENTSEQUENCENUMBER",m_sEmploymentSequenceNumber);
	XML.CreateTag("NOTICEPROBLEMINDICATOR", scScreenFunctions.GetRadioGroupValue(frmScreen, "NoticeProblemGroup"));
	XML.CreateTag("SHARESOWNEDINDICATOR", scScreenFunctions.GetRadioGroupValue(frmScreen, "SharesOwnedIndicatorGroup"));
	XML.CreateTag("PAYROLLNUMBER", frmScreen.txtPayrollNumber.value);
	XML.CreateTag("PERCENTSHARESHELD",frmScreen.txtPercentSharesHeld.value);
	XML.CreateTag("PROBATIONARYINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen, "ProbationaryGroup"));
	XML.CreateTag("P60SEENINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen, "P60SeenGroup"));
	XML.CreateTag("WAGESLIPSSEENINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen, "WageSlipsSeenGroup"));
	XML.CreateTag("EMPLOYMENTRELATIONSHIPIND",scScreenFunctions.GetRadioGroupValue(frmScreen, "EmploymentRelationshipGroup"));
	
	if (frmScreen.optEmploymentRelationshipYes.checked == true ) 
		XML.CreateTag("ACCOUNTANTGUID",m_sAccountantGUID);
	else
		XML.CreateTag("ACCOUNTANTGUID","");
	
	<%/* Save the details */%>
	// 	XML.RunASP(document,"SaveEmployedDetails.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"SaveEmployedDetails.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}


	bSuccess = XML.IsResponseOK();

	XML = null;
	return(bSuccess);
}


//2 functions added - DPF 24/06/02 - BMIDS00077
function frmScreen.optSharesOwnedIndicatorYes.onclick()

{
	frmScreen.txtPercentSharesHeld.value = m_dblPercentSharesHeld
	frmScreen.txtPercentSharesHeld.style.visibility = "visible"
}
function frmScreen.optSharesOwnedIndicatorNo.onclick()

{
	m_dblPercentSharesHeld = frmScreen.txtPercentSharesHeld.value
	frmScreen.txtPercentSharesHeld.value = ""
	frmScreen.txtPercentSharesHeld.style.visibility = "hidden"
}

function frmScreen.optEmploymentRelationshipYes.onclick()
{
	frmScreen.btnAccountants.disabled = false; 
}

function frmScreen.optEmploymentRelationshipNo.onclick()
{
	frmScreen.btnAccountants.disabled = true; 
}

function frmScreen.btnAccountants.onclick()
{
	<%/* Interface to Accountant Details popup */%>
	//next line replaced with line below - DPF 25/06/02 - BMIDS00077
	//var XML = new scXMLFunctions.XMLObject();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ArrayArguments = XML.CreateRequestAttributeArray(window);
	ArrayArguments[4] = m_sAccountantGUID;
	ArrayArguments[5] = m_bCanAddToDirectory;
	ArrayArguments[6] = m_sReadOnly;
	ArrayArguments[7] = "DC181";
	ArrayArguments[8] = m_sCustomerNumber;
	ArrayArguments[9] = m_sCustomerVersionNumber;
	ArrayArguments[10]= m_sEmploymentSequenceNumber;
	ArrayArguments[11] = scScreenFunctions.GetRadioGroupValue(frmScreen, "EmploymentRelationshipGroup");
	
	var sReturn = scScreenFunctions.DisplayPopup(window, document, "DC191.asp", ArrayArguments, 628, 480);

	if(sReturn != null)
		m_sAccountantGUID = sReturn[1];

	XML = null;
	
} 
 

-->
</script>
</body>
</html>


