<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /* 
Workfile:      DC182.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Self Employed Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		15/02/2000	Created
AY		30/03/00	New top menu/scScreenFunctions change
IVW		11/04/2000	Changed Prev/Next/Cancel to Ok/Cancel
BG		17/05/00	SYS0752 Removed Tooltips
MC		23/05/00	SYS0756 If Read-only mode, disabled fields
MH      14/06/00    SYS0791 DC191 Popup width
MC		21/06/00	SYS0956 Accountant not mandatory. Enable popup
					when in read only mode
CL		05/03/01	SYS1920 Read only functionality added
DJP		06/08/01	SYS2564 (parent) SYS3959 (child) Client cosmetic customisation						
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
SR		27/07/2004	BMIDS750 - Modified field width for Nature of business 

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ 
%>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MC		24/06/2004	   BMIDS750 NatureOfBusiness field added
MAH		17/11/2006		E2CR35 Addition of Combo CompanyStatusType	
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
<OBJECT id=scClientFunctions style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 type=text/x-scriptlet height=1 width=1 data=scClientFunctions.asp VIEWASTEXT></OBJECT>
<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>

<% /* FORMS */ %>
<form id="frmToDC160" method="post" action="DC160.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC170" method="post" action="DC170.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC192" method="post" action="DC192.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" year4 validate="onchange" mark>
<div style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 332px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Trading Name
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="txtEmployerName" maxlength="50" style="WIDTH: 200px; POSITION: absolute" class="msgReadOnly" readonly tabindex=-1>
		</span> 
	</span>
	<%/*=[MC]BMIDS750 Nature of Business field added*/%>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 34px" class="msgLabel">
		Nature of Business
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="txtNatureOfBusiness" maxlength="50" style="WIDTH: 300px; POSITION: absolute" class="msgTxt">
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 58px" class="msgLabel">
		Date Financial Interest held
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="txtDateFinancialInterestHeld" maxlength="10" style="WIDTH: 100px; POSITION: absolute" class="msgTxt">
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 82px" class="msgLabel">
		Registration Number
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="txtRegistrationNumber" maxlength="50" style="WIDTH: 200px; POSITION: absolute" class="msgTxt">
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 106px" class="msgLabel">
		<LABEL id=idVATNumber></LABEL>
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="txtVATNumber" maxlength="50" style="WIDTH: 200px; POSITION: absolute" class="msgTxt">
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 130px" class="msgLabel">
		<LABEL id=idPercentageSharesHeld></LABEL>
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="txtPercentSharesHeld" maxlength="3" style="WIDTH: 30px; POSITION: absolute" class="msgTxt">
			<span style="LEFT: 40px; POSITION: absolute; TOP: 3px" class="msgLabel">%</span>
		</span> 
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 152px" class="msgLabel">
		Is the business; LTD, Partnership<br>or Sole Trader?
		<span style="LEFT: 180px; POSITION: absolute; TOP: 4px">
			<select id="cboCompanyStatusType" style="WIDTH: 200px" class="msgCombo" NAME="cboCompanyStatusType"></select>
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 193px" class="msgLabel">
		Accountant
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="txtAccountantName" name="Accountant Name" maxlength="30" style="WIDTH: 200px; POSITION: absolute" class="msgReadOnly" readonly tabindex=-1>
		</span>
		<span style="LEFT: 385px; POSITION: absolute; TOP: -9px">				
			<input id="btnAccountant" type="button" style="WIDTH: 32px; HEIGHT: 32px" class ="msgDDButton">				
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 217px" class="msgLabel">
		Connections with other businesses
		<span style="LEFT: 0px; POSITION: absolute; TOP: 16px"><TEXTAREA class=msgTxt id=txtOtherBusinessConnections style="WIDTH: 300px; POSITION: absolute" rows=6></TEXTAREA> 
		</span> 
	</span> 
</div>
</form>

<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 410px; HEIGHT: 19px">
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/DC182attribs.asp" -->
<!-- #include FILE="Customise/DC182Customise.asp" -->
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sCustomerNumber = "";
var m_sCustomerVersionNumber = "";
var m_sEmploymentSequenceNumber = "";
var m_sEmployerName = "";
var SelfEmployedDetailsXML = null;
var m_sAccountantGUID = "";
var m_bCanAddToDirectory = false;
var scScreenFunctions;
var m_blnReadOnly = false;


/* EVENTS */

function btnCancel.onclick()
{
	//clear the contexts
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "Edit");
	frmToDC170.submit();		
}

function btnSubmit.onclick()
{
	if (CommitChanges())
	{
		//clear the contexts
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
		frmToDC192.submit();
	}
}

function frmScreen.btnAccountant.onclick()
{
	<%/* Interface to Accountant Details popup */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ArrayArguments = XML.CreateRequestAttributeArray(window);
	ArrayArguments[4] = m_sAccountantGUID;
	ArrayArguments[5] = m_bCanAddToDirectory;
	ArrayArguments[6] = m_sReadOnly;

	var sReturn = scScreenFunctions.DisplayPopup(window, document, "DC191.asp", ArrayArguments, 628, 480);

	if(sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_sAccountantGUID = sReturn[1];
		frmScreen.txtAccountantName.value = sReturn[2];
	}
	XML = null;
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function frmScreen.txtOtherBusinessConnections.onblur()
{
if (this.value.length>255)
	{
		alert("Cannot enter more than 255 characters");
		this.value = this.value.substr(0,255)
		this.focus();
	}
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

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Self Employed Details","DC182",scScreenFunctions);

	RetrieveContextData();
	SetMasks();

	Validation_Init();	
	Initialise(true);

	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC182");
	if (m_blnReadOnly == true) m_sReadOnly = "1";
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	// DJP SYS2564 (parent) SYS3959 (child) for client customisation
	Customise();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

/* FUNCTIONS */

function SetScreenToReadOnly()
{
	scScreenFunctions.SetScreenToReadOnly(frmScreen);
	// frmScreen.btnAccountant.disabled = true;
}

function CommitChanges()
{
	var bSuccess = true;

	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
			if (IsChanged())
				if (frmScreen.txtPercentSharesHeld.value > 100)
				{
					alert("A valid percentage must be specified (0-100).");
					frmScreen.txtPercentSharesHeld.focus();
					bSuccess = false;
				}
				<% /* SYS0956 Accountant not mandatory 
				else if (m_sAccountantGUID == "")
				{
					alert("An accountant must be specified.");
					bSuccess = false;
				} */ %>
				else
					bSuccess = SaveSelfEmployedDetails();
	return(bSuccess);
}

function DefaultFields()
// Inserts default values into all fields
{
	with (frmScreen)
	{
		txtDateFinancialInterestHeld.value = "";
		//*=[MC]BMIDS750 - Field added
		txtNatureOfBusiness.value="";
		//section end
		txtOtherBusinessConnections.value = "";
		txtPercentSharesHeld.value = "";
		txtRegistrationNumber.value = "";
		txtVATNumber.value = "";
		cboCompanyStatusType.value = "";
	}
}

function Initialise(bOnLoad)
// Initialises the screen
{
	PopulateCombos(); <%/* MAH 17/11/2006 E2CR35*/%>
	frmScreen.txtEmployerName.value = m_sEmployerName;

	if (!PopulateScreen())
		DefaultFields();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function PopulateScreen()
// Populates the screen with details of the item selected in dc160
{
	SelfEmployedDetailsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	SelfEmployedDetailsXML.CreateRequestTag(window,null);
	SelfEmployedDetailsXML.CreateActiveTag("SELFEMPLOYEDDETAILS");
	SelfEmployedDetailsXML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	SelfEmployedDetailsXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	SelfEmployedDetailsXML.CreateTag("EMPLOYMENTSEQUENCENUMBER", m_sEmploymentSequenceNumber);
	SelfEmployedDetailsXML.RunASP(document,"GetSelfEmployedDetails.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = SelfEmployedDetailsXML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[1] == ErrorTypes[0]))
	{
		//Error: record not found
		m_sMetaAction = "Add";
	}
	else if (ErrorReturn[0] == true)
	{
		// No error
		m_sMetaAction = "Edit";

		if(SelfEmployedDetailsXML.SelectTag(null, "SELFEMPLOYEDDETAILS") != null)
			with (frmScreen)
			{
				txtDateFinancialInterestHeld.value = SelfEmployedDetailsXML.GetTagText("DATEFINANCIALINTERESTHELD");
				txtOtherBusinessConnections.value = SelfEmployedDetailsXML.GetTagText("OTHERBUSINESSCONNECTIONS");
				txtPercentSharesHeld.value = SelfEmployedDetailsXML.GetTagText("PERCENTSHARESHELD");
				txtRegistrationNumber.value = SelfEmployedDetailsXML.GetTagText("REGISTRATIONNUMBER");
				txtVATNumber.value = SelfEmployedDetailsXML.GetTagText("VATNUMBER");
				//*=[MC]BMIDS750
				txtNatureOfBusiness.value = SelfEmployedDetailsXML.GetTagText("NATUREOFBUSINESS");
				//SECTION END
				<%/* MAH 17/11/2006 E2CR35 start*/%>
				if(SelfEmployedDetailsXML.GetTagText("COMPANYSTATUSTYPE") != "")
					cboCompanyStatusType.value = SelfEmployedDetailsXML.GetTagText("COMPANYSTATUSTYPE");
				else
					cboCompanyStatusType.value = "";
				<%/* MAH 17/11/2006 E2CR35 end */%>
				
				if (SelfEmployedDetailsXML.SelectTag(null, "ACCOUNTANT") != null)
				{
					m_sAccountantGUID = SelfEmployedDetailsXML.GetTagText("ACCOUNTANTGUID");
					frmScreen.txtAccountantName.value = SelfEmployedDetailsXML.GetTagText("COMPANYNAME");
				}
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
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_bCanAddToDirectory = (sUserRole >= XML.GetGlobalParameterAmount(document,"AddToDirectoryRole"));

	XML = null;
}

function SaveSelfEmployedDetails()
{
	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateRequestTag(window,null)

	XML.CreateActiveTag("SELFEMPLOYEDDETAILS");
	XML.CreateTag("ACCOUNTANTGUID", m_sAccountantGUID);
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	XML.CreateTag("EMPLOYMENTSEQUENCENUMBER",m_sEmploymentSequenceNumber);
	XML.CreateTag("DATEFINANCIALINTERESTHELD",frmScreen.txtDateFinancialInterestHeld.value);
	XML.CreateTag("OTHERBUSINESSCONNECTIONS",frmScreen.txtOtherBusinessConnections.value);
	XML.CreateTag("PERCENTSHARESHELD",frmScreen.txtPercentSharesHeld.value);
	XML.CreateTag("REGISTRATIONNUMBER",frmScreen.txtRegistrationNumber.value);
	XML.CreateTag("VATNUMBER",frmScreen.txtVATNumber.value);
	//*=[MC]BMIDS750
	XML.CreateTag("NATUREOFBUSINESS",frmScreen.txtNatureOfBusiness.value);
	//SECTION END
	XML.CreateTag("COMPANYSTATUSTYPE", frmScreen.cboCompanyStatusType.value); <%/* MAH 17/11/2006 E2CR35 */%>
	<%/* Save the details */%>
	// 	XML.RunASP(document,"SaveSelfEmployedDetails.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"SaveSelfEmployedDetails.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}


	bSuccess = XML.IsResponseOK();

	XML = null;
	return(bSuccess);
}
<%/* MAH 17/11/2006 E2CR35 start */%>
function PopulateCombos()
{
	var blnSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	<%/* Company Status Type */%>
	var sGroupList = new Array("CompanyStatusType");
	if(XML.GetComboLists(document,sGroupList))
		blnSuccess = blnSuccess & XML.PopulateCombo(document,frmScreen.cboCompanyStatusType,"CompanyStatusType",true);
	
	if(!blnSuccess) scScreenFunctions.SetScreenToReadOnly(frmScreen);
	XML = null;
}
<%/* MAH 17/11/2006 E2CR35 End */%>
-->
</script>
</body>
</html>


