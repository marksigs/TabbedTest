<%@ LANGUAGE="JSCRIPT" %>
<HTML>
	<HEAD>
		<title>Credit Check Summary <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
		<%
/*
Workfile:      ra021.asp
Copyright:     Copyright © 2005 Marlborough Stirling

Description:   Credit Check Summary Screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS History:

Prog	Date		Description
HM		24/08/2005	MAR13 Original
PE		13/12/2005	MAR309
PE		03/03/2006	MAR1037 - CIF number not showing for first application
DRC     16/03/2006  MAR1305 - & MAR1403 - Dispaly Exp Ref No
DRC     25/03/2006  MAR1482 - substitute for £ sign
PE		27/03/2006	MAR1482 - Fixed comment (no closing%)
PE		29/03/2006	MAR1504 - Error with the display of CAIS data 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM History:

Prog	Date		Description
SAB		14/04/2006	EP341 - Replaced Risk Indicator with Eligibility Indicator
SAB		09/05/2006	EP481 - Updated to handle new address XML - also removed the call
							to FindCustomerAddressList
PB		21/04/2006  EP529 - Merged MAR1623 - Failure to display the bankruptcy flag in the credit check screen
							changed BX02_BANKRUPTCYPRESENT to BX02_BANKRUPTCYDETECTED.
PB		27/04/2006  EP529 - Merged MAR1674 - Failure to display the unsecured arrears flag in the credit check screen
							changed BX18_UNSECUREDARREARS to BX18_UNSECUREDARREARS_2.
*/
%>
		<META content="Microsoft Visual Studio 6.0" name="GENERATOR">
		<LINK href="stylesheet.css" type="text/css" rel="STYLESHEET">
	</HEAD>
	<BODY>
		<OBJECT id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px"
			tabIndex="-1" type="text/x-scriptlet" height="1" width="1" data="scClientFunctions.asp"
			VIEWASTEXT>
		</OBJECT>
		<OBJECT id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex="-1" type="text/x-scriptlet"
			height="1" data="scTable.htm" VIEWASTEXT>
		</OBJECT>
		<SPAN style="DISPLAY: none; LEFT: 0px; TOP: 0px">
			<OBJECT id="scScrollTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex="-1" type="text/x-scriptlet"
				data="scTableListScroll.asp" VIEWASTEXT>
			</OBJECT>
		</SPAN>
		<% /* Scriptlets - remove any which are not required */ %>
		<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
		<FORM id="frmScreen" style="VISIBILITY: hidden" year4 validate="onchange" mark>
			<TABLE cellSpacing="5" cellPadding="0" width="607" border="0" ID="Table4">
				<TR valign="middle">  
					<TD class="msgGroup" colSpan="2">
						<TABLE class="msgGroup" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="msgLabelHead" colSpan="4">Credit Check Summary</TD>
							</TR>
							<TR>
								<TD class="msgLabel" width="20%">First Applicant</TD>
								<TD align="left" width="40%"><INPUT class="msgReadOnly" id="txtFirstApplicant" style="WIDTH: 200px" maxLength="30" name="FirstApplicant">
								</TD>
								<TD class="msgLabel" width="25%">CIF</TD>
								<TD align="center">
									<DIV align="left"><INPUT class="msgReadOnly" id="txtCIF1" style="WIDTH: 100px" maxLength="10" name="CIF1">
									</DIV>
								</TD>
							</TR>
							<TR>
								<TD class="msgLabel">Second Applicant</TD>
								<TD align="left"><INPUT class="msgReadOnly" id="txtSecondApplicant" style="WIDTH: 200px" maxLength="30"
										name="SecondApplicant">
								</TD>
								<TD class="msgLabel">CIF</TD>
								<TD align="center">
									<DIV align="left"><INPUT class="msgReadOnly" id="txtCIF2" style="WIDTH: 100px" maxLength="10" name="CIF2">
									</DIV>
								</TD>
							</TR>
							<TR>
								<TD class="msgLabelHead" colSpan="4"><BR>
									Voters Roll
								</TD>
							</TR>
							<TR>
								<TD colSpan="4">
									<TABLE class="msgTable" id="tblAddressSummary" cellSpacing="0" cellPadding="0" width="97%"
										border="0">
										<TR id="rowTitles">
											<TD class="TableHead" width="20%">Type&nbsp;</TD>
											<TD class="TableHead" width="65%">Address&nbsp;</TD>
											<TD class="TableHead" width="15%">No. of Years&nbsp;</TD>
										</TR>
										<TR id="row01">
											<TD class="TableTopLeft">&nbsp;</TD>
											<TD class="TableTopCenter">&nbsp;</TD>
											<TD class="TableTopRight">&nbsp;</TD>
										</TR>
										<TR id="row02">
											<TD class="TableLeft">&nbsp;</TD>
											<TD class="TableCenterCenter">&nbsp;</TD>
											<TD class="TableRight">&nbsp;</TD>
										</TR>
										<TR id="row03">
											<TD class="TableLeft">&nbsp;</TD>
											<TD class="TableCenterCenter">&nbsp;</TD>
											<TD class="TableRight">&nbsp;</TD>
										</TR>
										<TR id="row04">
											<TD class="TableBottomLeft">&nbsp;</TD>
											<TD class="TableBottomCenter">&nbsp;</TD>
											<TD class="TableBottomRight">&nbsp;</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="msgGroup" colSpan="2">
						<TABLE class="msgGroup" id="Table1" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="msgLabelHead" colSpan="4">Credit Decision Detail</TD>
							</TR>
							<TR>
								<TD class="msgLabel" width="20%">Experian Ref</TD>
								<TD width="40%"><INPUT class="msgReadOnly" id="txtExperianRef" style="WIDTH: 100px" maxLength="30" name="ExperianRef"></TD>
								<TD class="msgLabel" width="22%">Detect Score</TD>
								<TD align="center"><INPUT class="msgReadOnly" id="txtDetectScore" style="WIDTH: 120px" maxLength="30" name="DetectScore"></TD>
							</TR>
							<TR>
								<TD class="msgLabel">Delphi Total</TD>
								<TD><INPUT class="msgReadOnly" id="txtDelphiTotal" style="WIDTH: 100px" maxLength="30" name="DelphiTotal"></TD>
								<TD class="msgLabel">Eligibility Indicator</TD>
								<TD align="center"><INPUT class="msgReadOnly" id="txtRiskIndicator" style="WIDTH: 120px" maxLength="30" name="RiskIndicator"></TD>
							</TR>
							<TR>
								<TD class="msgLabel">Flags:</TD>
								<TD colSpan="3"><INPUT id="chkNOCNOD" type="checkbox" value="checkbox" name="NOCNOD">
									<SPAN class="msgLabel" style="WIDTH: 65px">NOC / 
            NOD</SPAN>
									<INPUT id="chkCIFAS" type="checkbox" value="checkbox" name="CIFAS">
									<SPAN class="msgLabel" style="WIDTH: 100px">CIFAS</SPAN>
									<INPUT id="chkCML" type="checkbox" value="checkbox" name="CML">
									<SPAN class="msgLabel" style="WIDTH: 100px">CML</SPAN>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="msgGroup" vAlign="top" width="60%">
						<TABLE class="msgGroup" id="Table2" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="msgLabelHead" colSpan="3">Arrears Defaults and Adverse Data</TD>
							</TR>
							<TR>
								<TD class="msgLabel">&nbsp;</TD>
								<TD class="msgLabelHead" align="center">Volume</TD>
								<TD class="msgLabelHead" align="center">Value</TD>
							</TR>
							<TR>
								<TD class="msgLabel">All CAIS 8/9</TD>
								<TD>
									<DIV align="center"><INPUT class="msgReadOnly" id="txtCAIS89Volume" style="WIDTH: 100px" maxLength="30" size="30"
											name="CAIS89Volume">
									</DIV>
								</TD>
								<TD>
									<DIV align="center"><INPUT class="msgReadOnly" id="txtCAIS89Value" style="WIDTH: 100px" maxLength="30" name="CAIS89Value">
									</DIV>
								</TD>
							</TR>
							<TR>
								<TD class="msgLabel">CCJs</TD>
								<TD>
									<DIV align="center"><INPUT class="msgReadOnly" id="txtCCJVolume" style="WIDTH: 100px" maxLength="30" name="CCJVolume">
									</DIV>
								</TD>
								<TD>
									<DIV align="center"><INPUT class="msgReadOnly" id="txtCCJValue" style="WIDTH: 100px" maxLength="30" name="CCJValue">
									</DIV>
								</TD>
							</TR>
							<TR>
								<TD></TD>
							</TR>
							<TR>
								<TD class="msgLabelHead" vAlign="top" rowSpan="2">Adverse Flags:</TD>
								<TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT id="chkBankruptcy" type="checkbox" value="checkbox" name="Bankruptcy">
									<SPAN class="msgLabel">Bankruptcy</SPAN></TD>
								<TD><INPUT id="chkSecuredArrears" type="checkbox" value="checkbox" name="SecuredArrears">
									<SPAN class="msgLabel">Secured 
        Arrears</SPAN></TD>
							</TR>
							<TR>
								<TD>&nbsp;&nbsp;&nbsp;&nbsp; <INPUT id="chkIVA" type="checkbox" value="checkbox" name="IVA">
									<SPAN class="msgLabel">IVA</SPAN></TD>
								<TD><INPUT id="chkUnsecuredArrears" type="checkbox" value="checkbox" name="UnsecuredArrears">
									<SPAN class="msgLabel">Unsecured Arrears</SPAN>
								</TD>
							</TR>
							<TR>
								<TD></TD>
							</TR>
							<TR>
								<TD class="msgLabelHead" colSpan="3">Previous Searches
								</TD>
							</TR>
							<TR>
								<TD class="msgLabel">&nbsp;</TD>
								<TD class="msgLabelHead" align="center">Total</TD>
								<TD></TD>
							</TR>
							<TR>
								<TD class="msgLabel">last 3 months</TD>
								<TD>
									<DIV align="center"><INPUT class="msgReadOnly" id="txtLast3monthsTotal" style="WIDTH: 100px" maxLength="30"
											name="Last3monthsTotal">
									</DIV>
								</TD>
							</TR>
							<TR>
								<TD class="msgLabel">last 6 months</TD>
								<TD>
									<DIV align="center"><INPUT class="msgReadOnly" id="txtLast6monthsTotal" style="WIDTH: 100px" maxLength="30"
											name="Last6monthsTotal">
									</DIV>
								</TD>
							</TR>
							<TR>
								<TD class="msgLabel">last 12 months</TD>
								<TD>
									<DIV align="center"><INPUT class="msgReadOnly" id="txtLast12monthsTotal" style="WIDTH: 100px" maxLength="30"
											name="Last12monthsTotal">
									</DIV>
								</TD>
							</TR>
							<TR>
								<TD></TD>
							</TR>
						</TABLE>
					</TD>
					<TD class="msgGroup">
						<TABLE class="msgGroup" id="Table3" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="msgLabelHead" colSpan="3">All CAIS Status</TD>
							</TR>
							<TR>
								<TD width="54%">&nbsp;</TD>
								<TD class="msgLabelHead" align="center">Status</TD>
								<TD></TD>
							</TR>
							<TR>
								<TD class="msgLabel">Worst Current</TD>
								<TD align="center"><INPUT class="msgReadOnly" id="txtWorstCurrentStatus" style="WIDTH: 100px" maxLength="30"
										name="WorstCurrentStatus">
								</TD>
							</TR>
							<TR>
								<TD class="msgLabel">Worst in Last 6 months</TD>
								<TD align="center"><INPUT class="msgReadOnly" id="txtWorstLast6monthStatus" style="WIDTH: 100px" maxLength="30"
										name="WorstLast6monthStatus">
								</TD>
							</TR>
							<TR>
								<TD></TD>
							</TR>
							<TR>
								<TD class="msgLabelHead" colSpan="3">Active CAIS Worst Status in Last 3 months
								</TD>
							</TR>
							<TR>
								<TD width="54%">&nbsp;</TD>
								<TD class="msgLabelHead" align="center">Status</TD>
								<TD></TD>
							</TR>
							<TR>
								<TD class="msgLabel">Opened &lt; 12 months ago</TD>
								<TD align="center"><INPUT class="msgReadOnly" id="txtOpenedLess12monthsAgo" style="WIDTH: 100px" maxLength="30"
										name="OpenedLess12monthsAgo">
								</TD>
							</TR>
							<TR>
								<TD class="msgLabel">Opened &gt; 12 months ago</TD>
								<TD align="center"><INPUT class="msgReadOnly" id="txtOpenedMore12monthsAgo" style="WIDTH: 100px" maxLength="30"
										name="OpenedMore12monthsAgo">
								</TD>
							</TR>
							<TR>
								<TD></TD>
							</TR>
							<TR>
								<TD class="msgLabelHead" colSpan="3">Delinquent Accounts
								</TD>
							</TR>
							<TR>
								<TD class="msgLabel" width="54%">Total Value £</TD>
								<TD>
									<DIV align="center"><INPUT class="msgReadOnly" id="txtTotalValue" style="WIDTH: 100px" maxLength="30" name="TotalValue">
									</DIV>
								</TD>
							</TR>
							<TR>
								<TD class="msgLabel">Age Most Recent</TD>
								<TD>
									<DIV align="center"><INPUT class="msgReadOnly" id="txtAgeMostRecent" style="WIDTH: 100px" maxLength="30" name="AgeMostRecent">
									</DIV>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD colSpan="2">
						<TABLE class="msgGroup" id="Table6" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="msgLabelHead" width="217">Total CAIS Volume
								</TD>
								<TD width="108"><INPUT class="msgReadOnly" id="txtTotalCAISVolume" style="WIDTH: 100px" maxLength="30"
										name="TotalCAISVolume"></TD>
								<TD class="msgLabelHead" width="162">Total Unsecured CAIS Value
								</TD>
								<TD><INPUT class="msgReadOnly" id="txtTotalUnsecuredCAISValue" style="WIDTH: 100px" maxLength="30"
										name="TotalUnsecuredCAISValue"></TD>
							</TR>
							<TR>
								<TD class="msgLabelHead" width="217">CAIS Special Instruction Indicator</TD>
								<TD width="108"><INPUT class="msgReadOnly" id="txtCAISSpecialInstructionInd" style="WIDTH: 100px" maxLength="30"
										name="CAISSpecialInstructionInd"></TD>
								<TD class="msgLabelHead" width="162">Total Secured CAIS Value</TD>
								<TD><INPUT class="msgReadOnly" id="txtTotalSecuredCAISValue" style="WIDTH: 100px" maxLength="30"
										name="TotalSecuredCAISValue"></TD>
							</TR>
						</TABLE>
					<TD></TD>
				</TR>
			</TABLE>
		</FORM>
		<%/* Main Buttons */ %>
		<DIV id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 500px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" --></DIV>
		<% /* File containing field attributes */ %>
		<!-- #include FILE="attribs/ra021attribs.asp" -->
		<% /* Specify Code Here */ %>
		<SCRIPT language="JScript">

var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_nCurrentCustomerOrder = 1;
var m_XMLCC = null;
var m_XMLAddr = null;

var m_sMetaAction = null;

var m_sCreditCheckGUID = "";
var m_sCustomerName1 = "";
var m_sCustomerName2 = "";
var m_sCustomerNumber1 = "";
var m_sCustomerNumber2 = "";
var m_sCustomerVersion1 = "";
var m_sCustomerVersion2 = "";
var m_sCustomerCIF1 = "";
var m_sCustomerCIF2 = "";

//Voters Roll total rows and last column 
var m_iTableLength = 4;

/*
var m_NoOfYearsCurr1 = "";
var m_NoOfYearsCurr2 = "";
var m_NoOfYearsPrev1 = "";
var m_NoOfYearsPrev2 = "";
*/

var scScreenFunctions;
var m_sRequestAttribs = "";
var m_nMaxCustomers = 1;
var m_sReadOnly ="1"
var m_BaseNonPopupWindow = null;
var m_sArgArray = null;

var scClientScreenFunctions;

function window.onload()
{
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	SetMasks();
	
	<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
	%>	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);
	
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];	
	m_sArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();

	RetrieveContextData();
	Initialise();
	SetAllFieldsToReadOnly();
	window.returnValue = null;
	ClientPopulateScreen();		//attribs func
}

function RetrieveContextData()
{
	m_sApplicationNumber			= m_sArgArray[0];
	m_sApplicationFactFindNumber	= m_sArgArray[1];
	m_sCustomerNumber1				= m_sArgArray[2];
	m_sCustomerVersionNumber1		= m_sArgArray[3];
	m_sCustomerNumber2				= m_sArgArray[4];
	m_sCustomerVersionNumber2		= m_sArgArray[5];
	m_sCustomerName1				= m_sArgArray[6];
	m_sCustomerName2				= m_sArgArray[7];
	m_sRequestAttribs				= m_sArgArray[10];
	
	//update number of applicants
	if (m_sCustomerNumber2 == "") 
		m_nMaxCustomers = 1;
	else
		m_nMaxCustomers = 2;
}

function Initialise()
{
	GetOtherSystemCustomerNumbers();
	PopulateCustomerData();
	if (GetCreditCheckData())
	{
		PopulateCreditCheckData();
	}
	PopulateTable();
	//GetAddressList();
}

function SetAllFieldsToReadOnly()
{
	scScreenFunctions.ShowCollection(frmScreen);
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
}

function GetOtherSystemCustomerNumbers()
{
	var sCustomerNumber, sCustomerVersionNumber, sOtherSystemCustomerNumber;
	
	var ListXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	ListXML.CreateRequestTag(m_BaseNonPopupWindow,"SEARCH");
	ListXML.CreateActiveTag("APPLICATION");
	ListXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	ListXML.RunASP(document,"FindCustomersForApplication.asp");

	if (ListXML.IsResponseOK())
	{
		//var tagCustomerList = m_XMLCustomerList.CreateActiveTag("CUSTOMERLIST");
		ListXML.CreateTagList("CUSTOMERROLE");
		var iNoOfCustomers = ListXML.ActiveTagList.length;
		for(var nCount = 0; nCount < iNoOfCustomers; nCount++)
		{	
			ListXML.SelectTagListItem(nCount);
			sCustomerNumber = ListXML.GetTagText("CUSTOMERNUMBER");
			sCustomerVersionNumber = ListXML.GetTagText("CUSTOMERVERSIONNUMBER");
			sCustomerRoleType = ListXML.GetTagText("CUSTOMERROLETYPE");
			sOtherSystemCustomerNumber = ListXML.GetTagText("OTHERSYSTEMCUSTOMERNUMBER");
			
			//pick up only customer1 or customer2 info
			if(sCustomerNumber != "" && sCustomerVersionNumber != "" )
			{
				<% // MAR1037 - Peter Edney - 03/03/06 %>
				if (sCustomerNumber == m_sCustomerNumber1) {m_sCustomerCIF1 = sOtherSystemCustomerNumber;}
				if (sCustomerNumber == m_sCustomerNumber2) {m_sCustomerCIF2 = sOtherSystemCustomerNumber;}
			}
		}
	}
}

function PopulateCustomerData()
{
with (frmScreen)
	{
		<% /* Customer Specific Data */ %>
		txtFirstApplicant.value	= m_sCustomerName1;
		txtSecondApplicant.value = m_sCustomerName2; 
		txtCIF1.value = m_sCustomerCIF1;
		txtCIF2.value = m_sCustomerCIF2;
	}
}

function GetCreditCheckData()
{
	var bSuccess = false;

	<% /* Retrieve Credit Check data */ %>
	m_XMLCC = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	with (m_XMLCC)
	{
		<% /* Create request block */ %>
		CreateRequestTag(m_BaseNonPopupWindow,"SEARCH");
		CreateActiveTag("SEARCH");
		CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		CreateTag("CUSTOMERNUMBER1", m_sCustomerNumber1);
		CreateTag("CUSTOMERVERSIONNUMBER1", m_sCustomerVersionNumber1);
		CreateTag("CUSTOMERNUMBER2", m_sCustomerNumber2);
		CreateTag("CUSTOMERVERSIONNUMBER2", m_sCustomerVersionNumber2);
		
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					RunASP(document, "GetCreditCheckResult.asp");
				break;
			default: // Error
				SetErrorResponse();
			}

		if (IsResponseOK())
		{
			//return success even no credit check exists for the applicants 
			//m_sCreditCheckGUID = GetTagText("CREDITCHECKGUID");
			bSuccess = true;
		}
	}
	return bSuccess;
}

function PopulateCreditCheckData()
{
	with (frmScreen)
	{
		<% /* Credit Check Data */ %>
		txtExperianRef.value = m_XMLCC.GetTagText("");
		
		m_XMLCC.SelectTag(null,"DETECT");
		if (m_XMLCC.ActiveTag)
		{
			txtDetectScore.value = m_XMLCC.GetAttribute("DETECTCREDITSCORE");
		}
		
		if (m_XMLCC.SelectTag(null,"CB1") != null)
		{
		
			/*
			m_NoOfYearsCurr1 = m_XMLCC.GetAttribute("BX01_ERYEARSSPCURRENT1"); 
			m_NoOfYearsCurr2 = m_XMLCC.GetAttribute("BX01_ERYEARSSPCURRENT2");  
			m_NoOfYearsPrev1 = m_XMLCC.GetAttribute("BX01_ERYEARSSPPREVIOUS1"); 
			m_NoOfYearsPrev2 = m_XMLCC.GetAttribute("BX01_ERYEARSSPPREVIOUS2"); 
			*/
		
			// Peter Edney - MAR309 - 13/12/2005			
			// txtDelphiTotal.value = m_XMLCC.GetAttribute("BX13_OPTINSCORE");
			txtDelphiTotal.value = m_XMLCC.GetAttribute("DEC1_SCORE");
			// DRC MAR1503 
			txtExperianRef.value = m_XMLCC.GetAttribute("EXPERIAN_REF");
			txtRiskIndicator.value = m_XMLCC.GetAttribute("ELIGIBILITYINDICATOR");
		
			if (m_XMLCC.GetAttribute("BX08_OTHERNOC") == "Y") {chkNOCNOD.checked = true;}
			if (m_XMLCC.GetAttribute("BX05_CIFASDETECTED") == "Y") {chkCIFAS.checked = true;}
			if (m_XMLCC.GetAttribute("BX06_CMLDETECTED") == "Y") {chkCML.checked = true;}
		
			// Peter Edney - MAR309 - 13/12/2005			
			// txtCAIS89Volume.value = m_XMLCC.GetAttribute("BX03_TOTALCAIS89");
			// txtCAIS89Value.value = m_XMLCC.GetAttribute("BX03_TOTALVALUECAIS89");
			// txtCCJVolume.value = m_XMLCC.GetAttribute("BX02_TOTALCCJS");
			// txtCCJValue.value = m_XMLCC.GetAttribute("BX02_TOTALVALUECCJS");
			txtCAIS89Volume.value = FormatCreditResult("BX03_TOTALCAIS89", "BX03_TOTALCAIS89_2");
			txtCAIS89Value.value = FormatCreditResult("BX03_TOTALVALUECAIS89", "BX03_TOTALVALUECAIS89_2");
			txtCCJVolume.value = FormatCreditResult("BX02_TOTALCCJS", "BX02_TOTALCCJS_2");
			txtCCJValue.value = FormatCreditResult("BX02_TOTALVALUECCJS", "BX02_TOTALVALUECCJS_2");
		
			<% /* EP529 - MARS1623 */ %>
			//if (m_XMLCC.GetAttribute("BX02_BANKRUPTCYPRESENT") == "Y") {chkBankruptcy.checked = true;}
			if (m_XMLCC.GetAttribute("BX02_BANKRUPTCYDETECTED") == "Y") {chkBankruptcy.checked = true;}
			<% /* EP529 - MARS1623 End */ %>
			if (m_XMLCC.GetAttribute("BX18_INDIVIDUALVOLARRANGEMENT") == "Y") {chkIVA.checked = true;}
			if (m_XMLCC.GetAttribute("BX18_SECUREDARREARS") == "Y") {chkSecuredArrears.checked = true;}
			<% /* EP529 - MARS1674 */ %>
			//if (m_XMLCC.GetAttribute("BX18_UNSECUREDARREARS") == "Y") {chkUnsecuredArrears.checked = true;}
			if (m_XMLCC.GetAttribute("BX18_UNSECUREDARREARS_2") == "Y") {chkUnsecuredArrears.checked = true;}
			<% /* EP529 - MARS1674 End */ %>

			// Peter Edney - MAR309 - 13/12/2005			
			// txtLast3monthsTotal.value = m_XMLCC.GetAttribute("BX04_TOTALSEARCHESNONMO3M");
			// txtLast6monthsTotal.value = m_XMLCC.GetAttribute("BX04_TOTALSEARCHESNONMO6M");		
			txtLast3monthsTotal.value = FormatCreditResult("BX04_TOTALSEARCHESNONMO3M", "BX04_TOTALSEARCHESNONMO3M_2");
			txtLast6monthsTotal.value = FormatCreditResult("BX04_TOTALSEARCHESNONMO6M", "BX04_TOTALSEARCHESNONMO6M_2");
			
			txtLast12monthsTotal.value = m_XMLCC.GetAttribute("BX04_TOTALNDSEARCHES12M");
		
			// Peter Edney - MAR309 - 13/12/2005	
			// txtWorstCurrentStatus.value = m_XMLCC.GetAttribute("BX03_WORSTCURRSTATUSACTIVECAIS");
			// txtWorstLast6monthStatus.value = m_XMLCC.GetAttribute("BX03_WORSTSTATUSALLACTIVECAIS");
			//	
			// txtOpenedLess12monthsAgo.value = m_XMLCC.GetAttribute("BX03_WORSTSTATUSNDH3MLT12M");
			// txtOpenedMore12monthsAgo.value = m_XMLCC.GetAttribute("BX03_WORSTSTATUSNDH3MGT12M");
			//	
			// txtTotalValue.value = m_XMLCC.GetAttribute("BX03_TOTALVALUEDELINQUENTACC");
			// txtAgeMostRecent.value = m_XMLCC.GetAttribute("BX03_AGEOFLASTDELINQUENTACC");
			//	
			// txtTotalCAISVolume.value = m_XMLCC.GetAttribute("BX03_TOTALACTIVECAISACCOUNTS");
			// txtCAISSpecialInstructionInd.value = m_XMLCC.GetAttribute("BX03_CAISSPECIALINSTRUCTIONIND");
			// txtTotalUnsecuredCAISValue.value = m_XMLCC.GetAttribute("BX03_TOTALVALACTIVECAISEXCMORT");
			// txtTotalSecuredCAISValue.value = m_XMLCC.GetAttribute("BX03_TOTALVALACTIVECAISINCMORT");
			txtWorstCurrentStatus.value = FormatCreditResult("BX03_WORSTCURRSTATUSACTIVECAIS", "BX03_WORSTCURRSTATUSACTIVECAIS_2");
			txtWorstLast6monthStatus.value = FormatCreditResult("BX03_WORSTSTATUSALLACTIVECAIS", "BX03_WORSTSTATUSALLACTIVECAIS_2");
		
			txtOpenedLess12monthsAgo.value = FormatCreditResult("BX03_WORSTSTATUSNDH3MLT12M", "BX03_WORSTSTATUSNDH3MLT12M_2");
			txtOpenedMore12monthsAgo.value = FormatCreditResult("BX03_WORSTSTATUSNDH3MGT12M", "BX03_WORSTSTATUSNDH3MGT12M_2");
		
			txtTotalValue.value = FormatCreditResult("BX03_TOTALVALUEDELINQUENTACC", "BX03_TOTALVALUEDELINQUENTACC_2");
			txtAgeMostRecent.value = FormatCreditResult("BX03_AGEOFLASTDELINQUENTACC", "BX03_AGEOFLASTDELINQUENTACC_2");
		
			txtTotalCAISVolume.value = FormatCreditResult("BX03_TOTALACTIVECAISACCOUNTS", "BX03_TOTALACTIVECAISACCOUNTS_2");
			txtCAISSpecialInstructionInd.value = FormatCreditResult("BX03_CAISSPECIALINSTRUCTIONIND", "BX03_CAISSPECIALINSTRUCTIONIND_2");
			txtTotalUnsecuredCAISValue.value = FormatCreditResult("BX03_TOTALVALACTIVECAISEXCMORT", "BX03_TOTALVALACTIVECAISEXCMORT_2");
			txtTotalSecuredCAISValue.value = FormatCreditResult("BX03_TOTALVALACTIVECAISINCMORT", "BX03_TOTALVALACTIVECAISINCMORT_2");
		}
	}
}

// Peter Edney - MAR309 - 13/12/2005
function FormatCreditResult(sAttribute1, sAttribute2)
{
	var sValue1;
	var sValue2;
	
	sValue1 = m_XMLCC.GetAttribute(sAttribute1);
	
	sValue2 = m_XMLCC.GetAttribute(sAttribute2);
	
	if (sValue1 == "")
	{
		return "";
	}
	else			
	{
	<% /* this substitution necessary because SQL server won't allow a £ in stored proceudres */ %>
	    if (sValue1.substring(0,2) == "<s")
	       sValue1 = "<£" + sValue1.substring(2,sValue1.length);
		if (sValue2 == "")
		{
			return sValue1;
		}
		else
		{
		    if (sValue2.substring(0,2) == "<s")
				sValue2 = "<£" + sValue2.substring(2,sValue2.length);
			return sValue1 + ", " + sValue2;
		}
	}	
}

function PopulateTable()
{
	m_XMLCC.ActiveTag = null;
	m_XMLCC.CreateTagList("ADDRESS");
	var iNoOfCustomerAddresses = m_XMLCC.ActiveTagList.length;
	if (iNoOfCustomerAddresses > 0)
	{
		scScrollTable.initialiseTable(tblAddressSummary, 0, "", ShowTable, m_iTableLength, iNoOfCustomerAddresses);
		ShowTable(0);
		scScrollTable.setRowSelected(1);
	}
}

function ShowTable(iStart)
{			
	var sAddressType, sAddress, sNoOfYears;
	var ParentTag = null;
	var sCustAddrSeq ;

	m_XMLCC.ActiveTag = null;
	m_XMLCC.CreateTagList("ADDRESS");

	for (var iLoop = 0; iLoop < m_XMLCC.ActiveTagList.length && iLoop < m_iTableLength; iLoop++)
	{				
		m_XMLCC.SelectTagListItem(iLoop + iStart);

		sAddressType = m_XMLCC.GetTagText("ADDRESSTYPE");
		sNoOfYears = m_XMLCC.GetTagText("VOTERROLL");

		sAddress = GetAddressString(m_XMLCC);
		ShowRow(iLoop+1, sAddressType, sAddress, sNoOfYears);
	}											
}

function GetAddressString(XMLAddrLine)
{		
	function AddComma(sAddress, sTagValue, bAddComma)
	{																						
		if ((bAddComma==true) && (sTagValue != "") && (sAddress != "")) sTagValue = ", " + sTagValue;
		sAddress += sTagValue;
		return sAddress;
	}	

	var sAddress = "";
	sAddress = AddComma(sAddress, XMLAddrLine.GetTagText("POSTCODE"), false);
	sAddress = AddComma(sAddress, XMLAddrLine.GetTagText("FLATNUMBER"), true);
	sAddress = AddComma(sAddress, XMLAddrLine.GetTagText("BUILDINGORHOUSENAME"), true);
	sAddress = AddComma(sAddress, XMLAddrLine.GetTagText("BUILDINGORHOUSENUMBER"), true); 
	sAddress = AddComma(sAddress, XMLAddrLine.GetTagText("STREET"), true);
	sAddress = AddComma(sAddress, XMLAddrLine.GetTagText("DISTRICT"), true);
	sAddress = AddComma(sAddress, XMLAddrLine.GetTagText("TOWN"), true);
	sAddress = AddComma(sAddress, XMLAddrLine.GetTagText("COUNTY"), true);
	sAddress = AddComma(sAddress, XMLAddrLine.GetTagText("COUNTRY"), true);
						
	return sAddress;
}

function ShowRow(nIndex, sAddressType, sAddress, sYears)
{			 					
	scScreenFunctions.SizeTextToField(tblAddressSummary.rows(nIndex).cells(0),sAddressType);
	scScreenFunctions.SizeTextToField(tblAddressSummary.rows(nIndex).cells(1),sAddress);
	scScreenFunctions.SizeTextToField(tblAddressSummary.rows(nIndex).cells(2),sYears);
}

function btnSubmit.onclick()
{
	window.close();	
}


		</SCRIPT>
	</BODY>
</HTML>
