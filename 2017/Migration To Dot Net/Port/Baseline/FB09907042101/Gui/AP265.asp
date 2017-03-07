<%@ LANGUAGE="JSCRIPT" %>
<%
/*
Workfile:      AP265.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   ?????
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
0000	17.01.2001	Created By GD
JR		21/09/01	Omiplus24, Telephone Number changes
DRC     04/02/02    AQR 3070 - Use of omAQ.GetLoanCompositionDetails
SF		16/04/02	AQR 3342 - Telephone access was not appearing
AT		30/04/02	SYS4359 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		10/05/2002	BMIDS00004	Modified frmScreen.btnEdit.onclick(),RetrieveData()
MV		15/05/2002	BMIDS00004	Modified frmScreen.btnEdit.onclick(),Added DisableButtons()
GHun	22/05/2002	BMIDS00005  Added View and Get Contact Data buttons, disabled Edit for validation type C
GHun	28/05/2002	BMIDS00108	Fix errors introduced by Core upgrade
GHun	13/09/2002	BMIDS00442	Get CIF number for passing to GetCRSContactData
GHun	23/09/2002	BMIDS00506	m_iTotalRecords not always initialised when changing between customers
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MV		09/10/2002	BMIDS00556	Removed Validation_Init()
MDC		15/11/2002	BMIDS00955  Remove hardcoded ValueId for Further Advance TypeOfMortgage
Mc		19/04/2004	BMIDS517	blank space padded to the title text
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<html>

<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Valuation Property Details <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<% /* Scriptlets - remove any which are not required */ %>

<form id="frmScreen" >
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 200px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">

	<span style="TOP: 24px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<LABEL id="idValuationRequired"></LABEL>
		<span style="TOP: -3px; LEFT: 105px; POSITION: ABSOLUTE">
			<input id="txtValuationRequired" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 48px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Type Of Property
		<span style="TOP: -3px; LEFT: 105px; POSITION: ABSOLUTE">
			<input id="txtTypeOfProperty" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 72px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Property Description
		<span style="TOP: -3px; LEFT: 105px; POSITION: ABSOLUTE">
			<input id="txtPropertyDescription" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	
	<span style="TOP: 96px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Year Built
		<span style="TOP: -3px; LEFT: 105px; POSITION: ABSOLUTE">
			<input id="txtYearBuilt" style="WIDTH: 40px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>	
	
	<span style="TOP: 114px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Purchase Price/<br>Estimated Price
		<span style="TOP: 3px; LEFT: 105px; POSITION: ABSOLUTE">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtPurchaseEstimatedPrice" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>	
	
	<span style="TOP: 144px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<LABEL id="idValuationPrice"></LABEL>
		<span style="TOP: -3px; LEFT: 105px; POSITION: ABSOLUTE">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtValuationPrice" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>	
	
	<span style="TOP: 168px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Amount Requested
		<span style="TOP: -3px; LEFT: 105px; POSITION: ABSOLUTE">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtAmountRequested" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>	
	
	<span style="TOP: 24px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		Purpose Of Loan
		<span style="TOP: -3px; LEFT: 115px; POSITION: ABSOLUTE">
			<input id="txtPurposeOfLoan" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>	
	
	<span style="TOP: 48px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		Account No.
		<span style="TOP: -3px; LEFT: 115px; POSITION: ABSOLUTE">
			<input id="txtAccountNo" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 72px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		Current balance
		<span style="TOP: -3px; LEFT: 115px; POSITION: ABSOLUTE">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtCurrentBalance" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 96px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		<LABEL id="idLastValuer"></LABEL>
		<span style="TOP: -3px; LEFT: 115px; POSITION: ABSOLUTE">
			<input id="txtLastValuer" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 120px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		<LABEL id="idLastValuationDate"></LABEL>
		<span style="TOP: -3px; LEFT: 115px; POSITION: ABSOLUTE">
			<input id="txtLastValuationDate" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 144px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		Tenure
		<span style="TOP: -3px; LEFT: 115px; POSITION: ABSOLUTE">
			<input id="txtTenure" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	
	<span style="TOP: 168px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		Unexpired Lease
		<span style="TOP: -3px; LEFT: 115px; POSITION: ABSOLUTE">
			<input id="txtUnexpiredLease" style="WIDTH: 40px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>



</div>

<div style="TOP: 215px; LEFT: 10px; HEIGHT: 225px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">

		<span style="TOP: 24px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			<LABEL id="idPostcode"></LABEL>
			<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
				<input id="txtPostcode" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		
		<span style="TOP: 48px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			<LABEL id="idFlatNo"></LABEL>
			<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
				<input id="txtFlatNo" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		
		<span style="TOP: 72px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			<LABEL id="idHouseNameAndNo"></LABEL>
			<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
				<input id="txtHouseName" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
			</span>

			<span style="TOP: -3px; LEFT: 264px; POSITION: ABSOLUTE">
				<input id="txtHouseNumber" style="WIDTH: 45px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		
		<span style="TOP: 96px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			Street
			<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
				<input id="txtStreet" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>		
		
		<span style="TOP: 120px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			<LABEL id="idDistrict"></LABEL>
			<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
				<input id="txtDistrict" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>		
		
		<span style="TOP: 144px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			<LABEL id="idTown"></LABEL>
			<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
				<input id="txtTown" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>		
		
		<span style="TOP: 168px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			<LABEL id="idCounty"></LABEL>
			<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
				<input id="txtCounty" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>	
		
		<span style="TOP: 192px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			Country
			<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
				<input id="txtCountry" style="WIDTH: 200px" class="msgTxt">
			</span>
		</span>			
				
		<span style="TOP: 24px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
			Telephone for Access
			<span style="TOP: -3px; LEFT: 115px; POSITION: ABSOLUTE">
				<input id="txtTelephoneForAccess" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 72px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
			Ground Rent
			<span style="TOP: -3px; LEFT: 115px; POSITION: ABSOLUTE">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtGroundRent" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 96px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
			<LABEL id="idServiceCharge"></LABEL>
			<span style="TOP: -3px; LEFT: 115px; POSITION: ABSOLUTE">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtServiceCharge" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 144px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
			Vendor/Applicant
			<span style="TOP: -3px; LEFT: 115px; POSITION: ABSOLUTE">
				<input id="txtVendorApplicant" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 168px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
			Selling Agent
			<span style="TOP: -3px; LEFT: 115px; POSITION: ABSOLUTE">
				<input id="txtSellingAgent" style="WIDTH: 150px" class="msgTxt">
			</span>
		</span>
</div>


</form>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 450px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<!-- #include FILE="Customise/AP265Customise.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/AP265Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;

var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_sAccountGUID = "";
var m_sReadOnly = "";
var sParameters;
var m_sMortgageApplicationValue = "";
var m_sCustomerNumber = "";
var m_sCustomerVersionNumber = "";
var m_sCustomerAddressSequenceNumber = "";
var m_sCurrency = "";
var m_sStageId = "";
var m_sMortgageSubQuoteNumber = "";
var m_sQuoteNumber = "";
var m_BaseNonPopupWindow = null;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	var sButtonList = new Array("Submit");
	RetrieveData();
	Customise();
	scScreenFunctions.SetThisCurrency(frmScreen, m_sCurrency);
	ShowMainButtons(sButtonList);
	SetScreenOnReadOnly();	
	GetRecordOnEntry();
	window.returnValue = null;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function GetRecordOnEntry()
{
	var bGetNewProperty			= false;
	var bGetLoanComposition		= false;
	var bGetNewLoan				= false;
	var bGetQuotation           = false;
	var bGetMortgageAccount		= false;
	var bGetPropertyInsurance	= false;
	var bGetMortgageLoanList	= false;
	var bGetThirdParty			= false;
	
	if ((m_sApplicationNumber !="") && (m_sApplicationFactFindNumber !=""))
	{
		
		bGetNewProperty	= GetNewProperty();
		bGetNewLoan = GetNewLoan();
      //DRC AQR SYS3070		
		bGetQuotation = GetQuotationData();
		bGetLoanComposition	= GetLoanCompositionDetails();
      // AQR SYS3070
		bGetThirdParty	= GetThirdParty();
	} 

	<% /* BMIDS00955 MDC 15/11/2002 - Remove hardcoded ValueId for Further Advance TypeOfMortgage */ %>
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	if( XML.IsInComboValidationList(document,"TypeOfMortgage", m_sMortgageApplicationValue, ["F", "M"]))
//	if (m_sMortgageApplicationValue == "40") <%//a Further Advance (Should really read VALIDATIONTYPE)%>
	{		

			if (m_sAccountGUID !="")
			{
				
				bGetMortgageAccount			=	GetMortgageAccount();
				bGetMortgageLoanList		=	GetMortgageLoanList();
			} 		
			if (!((m_sCustomerNumber=="") || (m_sCustomerVersionNumber=="") || (m_sCustomerAddressSequenceNumber=="")))
			{
				
				bGetPropertyInsurance		=	GetPropertyInsurance();
			} 
	}
	
}
function GetNewProperty()
{
	var XMLnewproperty = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject;	
	var blnRetval=false;

	XMLnewproperty.CreateRequestTagFromArray(sParameters[0], "GetFullNewPropertyDetails");
	XMLnewproperty.CreateActiveTag("NEWPROPERTY");
	XMLnewproperty.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XMLnewproperty.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XMLnewproperty.RunASP(document,"GetFullNewPropertyDetails.asp");
	<%//CHECK RECORD NOT FOUND BEFORE POPULATING FIELDS
	%>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XMLnewproperty.CheckResponse(ErrorTypes);	
	if (ErrorReturn[0] || (ErrorReturn[1] == ErrorTypes[0])) 
	{
		if (ErrorReturn[1] != ErrorTypes[0])
		{
			<%//RECORD FOUND
			%>
			PopulateFromNewProperty(XMLnewproperty);
			blnRetval=true;			
		}

	}
	XMLnewproperty = null;
	return(blnRetval);
}

function GetNewLoan()
{
	var XMLnewloan = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject;
	var blnRetval=false;
	XMLnewloan.CreateRequestTagFromArray(sParameters[0], "GetNewLoanDetails");
	XMLnewloan.CreateActiveTag("NEWLOAN");
	XMLnewloan.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XMLnewloan.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);	
	XMLnewloan.RunASP(document,"GetNewLoanDetails.asp");
	<%//CHECK RECORD NOT FOUND BEFORE POPULATING FIELDS
	%>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XMLnewloan.CheckResponse(ErrorTypes);	
	if (ErrorReturn[0] || (ErrorReturn[1] == ErrorTypes[0])) 
	{
		if (ErrorReturn[1] != ErrorTypes[0])
		{
			<%//RECORD FOUND
			%>
			PopulateFromNewLoan(XMLnewloan);
			blnRetval=true;			
		}

	}		
	XMLnewloan = null;
	return(blnRetval);
}
// DRC AQR SYS 3070 Function Added

function GetQuotationData()
{
	var XMLQuotation = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject;
	var blnRetval=false;
	XMLQuotation.CreateRequestTagFromArray(sParameters[0], "GetData");
	XMLQuotation.CreateActiveTag("QUOTATION");
	XMLQuotation.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XMLQuotation.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
    XMLQuotation.CreateTag("QUOTATIONNUMBER", m_sQuoteNumber );
	
	XMLQuotation.RunASP(document,"GetQuotationData.asp");
	<%//CHECK RECORD NOT FOUND BEFORE POPULATING FIELDS
	%>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XMLQuotation.CheckResponse(ErrorTypes);	
	if (ErrorReturn[0] || (ErrorReturn[1] == ErrorTypes[0])) 
	{
		if (ErrorReturn[1] != ErrorTypes[0])
		{
			<%//RECORD FOUND
			%>
			XMLQuotation.SelectTag(null,"QUOTATION");
			m_sMortgageSubQuoteNumber = XMLQuotation.GetTagText("MORTGAGESUBQUOTENUMBER");
			blnRetval=true;			
		}

	}		
	XMLQuotation = null;
	return(blnRetval);
}

// DRC AQR SYS 3070 Function Added

function GetLoanCompositionDetails()
{
	var XMLloan = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject;
	var blnRetval=false;
	XMLloan.CreateRequestTagFromArray(sParameters[0], "GetLoanCompositionDetails");
	XMLloan.CreateActiveTag("LOANCOMPOSITION");
	XMLloan.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XMLloan.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
    XMLloan.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubQuoteNumber );
	//XMLloan.CreateTag("MORTGAGESUBQUOTENUMBER", "1" );
	XMLloan.CreateTag("CURRENTSTAGEID", m_sStageId );	
	XMLloan.RunASP(document,"AQGetLoanCompositionDetails.asp");
	<%//CHECK RECORD NOT FOUND BEFORE POPULATING FIELDS
	%>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XMLloan.CheckResponse(ErrorTypes);	
	if (ErrorReturn[0] || (ErrorReturn[1] == ErrorTypes[0])) 
	{
		if (ErrorReturn[1] != ErrorTypes[0])
		{
			<%//RECORD FOUND
			%>
			PopulateFromLoanComposition(XMLloan);
			blnRetval=true;			
		}

	}		
	XMLloan = null;
	return(blnRetval);
}

function GetMortgageAccount()
{
	var XMLmortgageaccount = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject;
	var blnRetval=false;
	XMLmortgageaccount.CreateRequestTagFromArray(sParameters[0], "GetNewLoanDetails");
	XMLmortgageaccount.CreateActiveTag("MORTGAGEACCOUNT");
	XMLmortgageaccount.CreateTag("ACCOUNTGUID", m_sAccountGUID);
	XMLmortgageaccount.RunASP(document,"GetMortgageAccount.asp");
	<%//CHECK RECORD NOT FOUND BEFORE PULLING OUT VALUES
	%>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XMLmortgageaccount.CheckResponse(ErrorTypes);	
	if (ErrorReturn[0] || (ErrorReturn[1] == ErrorTypes[0])) 
	{
		if (ErrorReturn[1] != ErrorTypes[0])
		{
			<%//RECORD FOUND
			%>
			XMLmortgageaccount.SelectTag(null,"MORTGAGEACCOUNT");
			<%//Populate Globals
			%>
			m_sCustomerNumber = XMLmortgageaccount.GetTagText("CUSTOMERNUMBER");
			m_sCustomerVersionNumber = XMLmortgageaccount.GetTagText("CUSTOMERVERSIONNUMBER");
			m_sCustomerAddressSequenceNumber = XMLmortgageaccount.GetTagText("CUSTOMERADDRESSSEQUENCENUMBER");
			blnRetval=true;			
		}

	}
	XMLmortgageaccount = null;
	return(blnRetval);
}

function GetPropertyInsurance()
{
	var XMLpropertyinsurance = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject;
	var blnRetval=false;
	XMLpropertyinsurance.CreateRequestTagFromArray(sParameters[0], "GetPropertyInsurance");
	XMLpropertyinsurance.CreateActiveTag("CUSTOMER");
	XMLpropertyinsurance.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XMLpropertyinsurance.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	XMLpropertyinsurance.CreateTag("CUSTOMERADDRESSSEQUENCENUMBER", m_sCustomerAddressSequenceNumber);
	XMLpropertyinsurance.RunASP(document,"GetPropertyInsuranceDetails.asp");
	<%//CHECK RECORD NOT FOUND BEFORE PULLING OUT VALUES	
	%>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XMLpropertyinsurance.CheckResponse(ErrorTypes);	
	if (ErrorReturn[0] || (ErrorReturn[1] == ErrorTypes[0])) 
	{
		if (ErrorReturn[1] != ErrorTypes[0])
		{
			<%//RECORD FOUND
			%>
			PopulateFromCurrentProperty(XMLpropertyinsurance);
			blnRetval=true;			
		}

	} 
	XMLpropertyinsurance = null;
	return(blnRetval);
}

function GetMortgageLoanList()
{
	var XMLmortgageloanlist = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject;	
	var XMLmortgageloanlist = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject;
	var blnRetval=false;

	XMLmortgageloanlist.CreateRequestTagFromArray(sParameters[0], "GetPropertyInsurance");
	XMLmortgageloanlist.CreateActiveTag("MORTGAGELOAN");
	XMLmortgageloanlist.CreateTag("ACCOUNTGUID", m_sAccountGUID);
	XMLmortgageloanlist.RunASP(document,"FindMortgageLoanList.asp");
	<%//CHECK RECORD NOT FOUND BEFORE PULLING OUT VALUES	
	%>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XMLmortgageloanlist.CheckResponse(ErrorTypes);	
	if (ErrorReturn[0] || (ErrorReturn[1] == ErrorTypes[0])) 
	{
		if (ErrorReturn[1] != ErrorTypes[0])
		{
			<%//RECORD FOUND
			%>
			PopulateFromMortgageLoanList(XMLmortgageloanlist);
			blnRetval=true;			
		}

	} 
	XMLmortgageloanlist = null;
	return(blnRetval);	
	

}
function GetThirdParty()
{
	var XMLthirdpartylist = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject;	

	var blnRetval=false;

	XMLthirdpartylist.CreateRequestTagFromArray(sParameters[0], "GetThirdParty");
	XMLthirdpartylist.CreateActiveTag("APPLICATION");
	XMLthirdpartylist.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XMLthirdpartylist.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	
	XMLthirdpartylist.RunASP(document,"FindApplicationThirdPartyList.asp");
	<%//CHECK RECORD NOT FOUND BEFORE PULLING OUT VALUES	
	%>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XMLthirdpartylist.CheckResponse(ErrorTypes);	
	if (ErrorReturn[0] || (ErrorReturn[1] == ErrorTypes[0])) 
	{
		if (ErrorReturn[1] != ErrorTypes[0])
		{
			<%//RECORD FOUND
			%>
			PopulateFromThirdParty(XMLthirdpartylist);
			blnRetval=true;			
		}

	} 
	XMLthirdpartylist = null;
	return(blnRetval);
}



function PopulateFromNewProperty(XML)
{
	XML.SelectTag(null,"NEWPROPERTY");
	frmScreen.txtValuationRequired.value		=		XML.GetTagAttribute("VALUATIONTYPE", "TEXT");
	frmScreen.txtTypeOfProperty.value			=		XML.GetTagAttribute("TYPEOFPROPERTY", "TEXT");
	frmScreen.txtPropertyDescription.value		=		XML.GetTagAttribute("DESCRIPTIONOFPROPERTY","TEXT");
	frmScreen.txtYearBuilt.value				=		XML.GetTagText("YEARBUILT");
	frmScreen.txtValuationPrice.value			=		XML.GetTagText("VALUATIONPRICE");
	frmScreen.txtTenure.value					=		XML.GetTagAttribute("TENURETYPE","TEXT");
	frmScreen.txtUnexpiredLease.value			=		XML.GetTagText("UNEXPIREDTERMOFLEASEYEARS");
	frmScreen.txtPostcode.value					=		XML.GetTagText("POSTCODE");
	frmScreen.txtFlatNo.value					=		XML.GetTagText("FLATNUMBER");
	frmScreen.txtHouseName.value				=		XML.GetTagText("BUILDINGORHOUSENAME");
	frmScreen.txtHouseNumber.value				=		XML.GetTagText("BUILDINGORHOUSENUMBER");
	frmScreen.txtStreet.value					=		XML.GetTagText("STREET");
	frmScreen.txtDistrict.value					=		XML.GetTagText("DISTRICT");
	frmScreen.txtTown.value						=		XML.GetTagText("TOWN");
	frmScreen.txtCounty.value					=		XML.GetTagText("COUNTY");
	frmScreen.txtCountry.value					=		XML.GetTagAttribute("COUNTRY","TEXT");
	//frmScreen.txtTelephoneForAccess.value		=		XML.GetTagText("ACCESSTELEPHONENUMBER");
	frmScreen.txtGroundRent.value				=		XML.GetTagText("GROUNDRENT");	
	frmScreen.txtServiceCharge.value			=		XML.GetTagText("SERVICECHARGE");
	frmScreen.txtVendorApplicant.value			=		XML.GetTagText("COMPANYNAME");
	frmScreen.txtAccountNo.value				=		XML.GetTagText("OTHERSYSTEMACCOUNTNUMBER");	
	
	//JR - Omiplus24
	var sTelephone = XML.GetTagText("COUNTRYCODE");
	sTelephone = sTelephone + " " + XML.GetTagText("AREACODE");
	sTelephone = sTelephone + " " + XML.GetTagText("ACCESSTELEPHONENUMBER");
	sTelephone = sTelephone + " " + XML.GetTagText("EXTENSIONNUMBER");
	frmScreen.txtTelephoneForAccess.value = sTelephone;
	//End
}

function PopulateFromNewLoan(XML)
{
	XML.SelectTag(null,"NEWLOAN");
// AQR SYS3203 & 3070 DRC change from New Loan to Loan Composition	
//	frmScreen.txtAmountRequested.value			=		XML.GetTagText("AMOUNTREQUESTED");
//	frmScreen.txtPurposeOfLoan.value			=		XML.GetTagAttribute("PURPOSEOFLOAN","TEXT");
	XML.SelectTag(null,"APPLICATION");	
	<%//Set global
	%>
	m_sAccountGUID = XML.GetTagText("ACCOUNTGUID");
	frmScreen.txtAccountNo.value				=		XML.GetTagText("OTHERSYSTEMACCOUNTNUMBER");	
	XML.SelectTag(null,"APPLICATIONFACTFIND");	
// AQR SYS3203 & 3070 DRC change from New LOan to Loan Composition	
	m_sQuoteNumber = XML.GetTagText("ACCEPTEDQUOTENUMBER");
	if ( m_sQuoteNumber == "")
		m_sQuoteNumber = XML.GetTagText("ACTIVEQUOTENUMBER");
//	frmScreen.txtPurchaseEstimatedPrice.value	=		XML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE");
//
}

function PopulateFromLoanComposition(XML)
{
    var sPurposeOfLoanID = "";
	XML.SelectTag(null,"APPLICATIONFACTFIND");	
	frmScreen.txtPurchaseEstimatedPrice.value	=		XML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE"); 
	XML.SelectTag(null,"MORTGAGESUBQUOTE");
	frmScreen.txtAmountRequested.value			=		XML.GetTagText("AMOUNTREQUESTED");
	XML.SelectTag(null,"LOANCOMPONENTLIST");
	// First One by default
	XML.SelectTag(null,"LOANCOMPONENT");
	sPurposeOfLoanID = XML.GetTagText("PURPOSEOFLOAN");
	frmScreen.txtPurposeOfLoan.value			=		GetPurposeOfLoanText(sPurposeOfLoanID);
}

function PopulateFromCurrentProperty(XML)
{
	XML.SelectTag(null,"CURRENTPROPERTY");
	frmScreen.txtLastValuer.value				=		XML.GetTagText("LASTVALUERNAME");
	frmScreen.txtLastValuationDate.value		=		XML.GetTagText("LASTVALUATIONDATE");
}

function PopulateFromMortgageLoanList(XML)
{
	XML.SelectTag(null,"MORTGAGELOANLIST");
	frmScreen.txtCurrentBalance.value			=		XML.GetTagText("LOANSNOTREDEEMED");
}
function PopulateFromThirdParty(XML)
{
	<%//Parse list for estate agents
	%>
	XML.ActiveTag = null;
	XML.CreateTagList("APPLICATIONTHIRDPARTY");
	var iCount;
	var sTypeText="";
	var iNumberOfTPs = 0;
	iNumberOfTPs = XML.ActiveTagList.length;
	
	for (iCount = 0; (iCount < iNumberOfTPs); iCount++)
	{
		XML.SelectTagListItem(iCount);
		var sTPType	= XML.GetTagText("THIRDPARTYTYPE");
			
		if (sTPType == 6)<%//Estate Agents%>
		{
			sTypeText = XML.GetTagText("COMPANYNAME");
		}
	}
	frmScreen.txtSellingAgent.value=sTypeText;
}




function SetScreenOnReadOnly()
{
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		btnSubmit.focus();
	}
}

function RetrieveData()
{
	var sArguments = window.dialogArguments;
	
	//Grab the window sizing properties...
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];
	sParameters					 = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sApplicationNumber		 = sParameters[1];	
	m_sApplicationFactFindNumber = sParameters[2];
	m_sReadOnly					 = sParameters[3];
	m_sMortgageApplicationValue	 = sParameters[4];
	m_sCurrency					 = sParameters[5];
	//SYS3070 added
	m_sStageId                   = sParameters[6];
<% /* fetch any other parameters passed by the calling screen here */ %>
}
// AQR SYS3070 DRC
function GetPurposeOfLoanText(sPurposeOfLoanID)
{
<%	// return the valuename from the PurposeOfLoan group for the valueid sPurposeOfLoanID
%>
	PurposeOfLoanXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("PurposeOfLoan");
	PurposeOfLoanXML.GetComboLists(document, sGroups);
		PurposeOfLoanXML.SelectTag(null, "LISTNAME");
	PurposeOfLoanXML.CreateTagList("LISTENTRY");
	var sValueName = "";
	for(var iCount = 0; iCount < PurposeOfLoanXML.ActiveTagList.length && sValueName == ""; iCount++)
	{
		PurposeOfLoanXML.SelectTagListItem(iCount);
		if(PurposeOfLoanXML.GetTagText("VALUEID") == sPurposeOfLoanID)
			sValueName = PurposeOfLoanXML.GetTagText("VALUENAME");
	}
	return sValueName;
}
function btnSubmit.onclick()
{
	//nothing is returned from here....
	window.close();
}
-->
</script>
</body>
</html>


<% /* OMIGA BUILD VERSION 020.02.04.11.00 */ %>




