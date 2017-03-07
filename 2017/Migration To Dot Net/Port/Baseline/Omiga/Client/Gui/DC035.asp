<%@ Language=JScript %>
<%
/******************************************************************************************	
	BMIDS747 - CC063
	
	FILENAME	:	DC035.asp
	Copyright	:	Copyright © 2004 Marlborough Stirling
	Description	:   This screen populates verification list and cusomer version data
					(Whoseencustomer,howcustomeridseen)
	
	History:
	------------------------------------------------------------------
	CC			INITIALS	Date			Description
	------------------------------------------------------------------
	BMIDS747	MC			18/05/2004		Created verification summary screen
	BMIDS819	JD			28/07/2004		Changed label
	BMIDS864    HMA         14/09/2004      Remove ApplicationNumber and ApplicationFactFindNumber from Verification table.
	EP16		SAB			29/03/2006		Include KYC status field and associated processing
	
******************************************************************************************/
%>

<HTML>
<HEAD>
	<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<META NAME="CONTENT" Content="Omiga4, Marlborough Stirling, Mortgage Software, Optimus Back Office, Point of Sale, Life and Pensions">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
	<OBJECT id=scClientFunctions style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 type=text/x-scriptlet height=1 width=1 data=scClientFunctions.asp VIEWASTEXT></OBJECT>
	<%/* SCRIPTLETS */%>
	<script src="validation.js" language="JScript"></script>
</HEAD>

<BODY>
	<!--VERIFICATION TABLE SCROLL OBJECT-->
	<span style="TOP: 336px; LEFT: 295px; POSITION: absolute;z-index:-1">
		<TABLE><TR><TD>
		<object data="scTableListScroll.asp" id="scVerificationListTable" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
		</TD></TR>
		</TABLE>
	</span>

	<!--ROUTING FORMS ENABLE TO SUBMIT TO THE REQUIRED TARGETS-->
	<form id="frmToDC035" method="post" action="DC035.asp" STYLE="DISPLAY: none"></form>
	<form id="frmToDC031" method="post" action="DC031.asp" STYLE="DISPLAY: none"></form>
	<form id="frmToDC030" method="post" action="DC030.asp" STYLE="DISPLAY: none"></form>

<form id="frmScreen" validate  ="onchange" mark>
<DIV class=msgGroup id=divVerificationSummary name="divVerificationSummary" style="LEFT: 10px; WIDTH: 600px; POSITION: absolute; TOP: 60px; HEIGHT: 345px">
	<SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 30px">
		Customer Name <INPUT id="IdCustomerName" name="CustomerName" style="LEFT: 160px; WIDTH: 240px; POSITION: absolute; HEIGHT: 20px" class="msgInput" size=30>
	</SPAN>
	<DIV id="divWhoHasSeen" name="divWhoHasSeen" style="LEFT: 10px; WIDTH: 590px; POSITION: absolute; TOP: 65px">
		<TABLE class="msgLabel" width="580" cellpadding=0 cellspacing=0>
			<TR>
				<TD WIDTH="28%">
					KYC Status
				</TD>
				<TD COLSPAN="3">
					<SELECT id="idKYCStatus" name="KYCStatus" class="msgTxt"  style="WIDTH: 180px">
					</SELECT>
				</TD>
			</TR>
			<TR>
				<TD WIDTH="28%">
					Who has seen the customer? 
				</TD>
				<TD WIDTH="22%">
					<SELECT id="idWhoHasSeenCustomer" name="WhoHasSeenCustomer" class="msgTxt" style="WIDTH: 111px">
						<OPTION value="Broker" selected>Broker</OPTION>
						<OPTION value="Broker">Advisor</OPTION>
					</SELECT>
				</TD>
				<TD WIDTH="28%">
					How was the customer's identification seen? 
				</TD>
				<TD WIDTH="22%" align=right>
					<SELECT id="idHowWasCustomerIdSeen" name="HowWasCustomerIdSeen" class="msgTxt" style="WIDTH: 131px">
						<OPTION value="Broker" selected>Broker</OPTION>
						<OPTION value="Broker">Advisor</OPTION>
					</SELECT>
				</TD>
			</TR>
		</TABLE>
	</DIV>
	<DIV>
		<span id="spnVerificationList" name="spnVerificationList" style="LEFT: 8px; POSITION: absolute; TOP: 120px">
		<table id="tblVerificationTable" width="580" border="0" cellspacing="0" cellpadding="0" class="msgTable">
		<tr id="rowTitles">
			<td width="22%" class="TableHead">Verification Type</td>	
			<td width="21%" class="TableHead">ID Type</td>	
			<td width="21%" class="TableHead">Issuer</td>		
			<td width="21%" class="TableHead">Reference</td>
			<td width="15%" class="TableHead">Date</td>
		</tr>
		<tr id="row01">		
			<td class="TableTopLeft">&nbsp;</td>		
			<td class="TableTopCenter">&nbsp;</td>		
			<td class="TableTopCenter">&nbsp;</td>		
			<td class="TableTopCenter">&nbsp;</td>		
			<td class="TableTopRight">&nbsp;</td>		
		</tr>
		<tr id="row02">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableRight">&nbsp;</td>			
		</tr>
		<tr id="row03">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableRight">&nbsp;</td>			
		</tr>
		<tr id="row04">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableRight">&nbsp;</td>			
		</tr>
		<tr id="row05">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableRight">&nbsp;</td>			
		</tr>
		<tr id="row06">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableRight">&nbsp;</td>			
		</tr>
		<tr id="row07">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableRight">&nbsp;</td>			
		</tr>
		<tr id="row08">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableRight">&nbsp;</td>			
		</tr>
		<tr id="row09">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableRight">&nbsp;</td>			
		</tr>
		<tr id="row10">		
			<td class="TableBottomLeft">&nbsp;</td>	
			<td class="TableBottomCenter">&nbsp;</td>	
			<td class="TableBottomCenter">&nbsp;</td>	
			<td class="TableBottomCenter">&nbsp;</td>	
			<td class="TableBottomRight">&nbsp;</td>	
		</tr>
		</table>
	</span>
	</DIV>
	
	</DIV>
	<DIV class=msgLabel id=dc014Buttons style="LEFT: 17px; POSITION: absolute; TOP: 360px" name="dc014Buttons">
		<INPUT class=msgButton id=btnAdd style="WIDTH: 60px" type=button value=Add name=btnAdd Text="Add"> 
		<INPUT class=msgButton id=btnEdit style="WIDTH: 60px" type=button value=Edit name=btnEdit Text="Edit">
		<INPUT class=msgButton id=btnDelete style="WIDTH: 60px" type=button value=Delete name=btnDelete Text="Delete">  
	</DIV>


</form>
<div id="msgButtons" style="LEFT: 10px; WIDTH: 590px; POSITION: absolute; TOP: 440px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" --> 
</div>


<span id="spnToLastField" tabindex="0"></span>
<!-- #include FILE="fw030.asp" -->
<!--  #include FILE="attribs/DC035attribs.asp" -->
<script language="JScript">

	var m_sMetaAction = "";
	var m_sContext = null;
	var m_sReadOnly	= "";	
	var m_sApplicationNumber = "";
	var m_sApplicationFactFindNumber = "";
	var m_sCustomer1Number = "";
	var m_sCustomer1Version = "";
	var m_sCustomer2Number = "";
	var m_sCustomer2Version = "";
	var m_sCustomer3Number = "";
	var m_sCustomer3Version = "";
	var m_sCustomer4Number = "";
	var m_sCustomer4Version = "";
	var m_sCustomer5Number = "";
	var m_sCustomer5Version = "";
	var m_sCurrentCustomerNumber = "";
	var m_sCurrentCustomerVersionNumber = "";
	var m_sDateOfBirth = "";
	var m_sNextCustNum = "";
	var m_sNextCustVersion = "";
	var m_sSurname = "";
	var m_sFirstForename = "";
	var m_sPreferredname = "";
	var CustomerXML = null;
	var m_blnReadOnly = false;
	var m_sProcessInd = "";
	var scScreenFunctions;
	var scClientScreenFunctions;
	var ListXML = null;
	var XMLArray = new Array();

	/*******************************************************************************
	*	Function	:	Window.onload()
	*	Author		:	Mallesh Cheekoti
	*	Date		:	18/05/2004
	*	
	*	Purpose		:	This function executes onload of the document (DC035.asp).
	*******************************************************************************/
	function window.onload()
	{
		/*SETUP MAIN BUTTONS*/
		var sButtonList = new Array("Submit");
		ShowMainButtons(sButtonList);
		try
		{
			/*Get ScreenFunctions object from frame 1*/
			scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
			
			/*SET Title Section*/
			FW030SetTitles("Verification Summary","DC035",scScreenFunctions);
			
			/*Retrieve context data in document level variable holders*/
			RetrieveContextData();

			//*=[MC]Function added through ASP include DC035Attrib.asp
			SetMasks();
			Validation_Init();
			/*
				INITIALIZE DC035 PAGE
			*/		
			initialize();
		}
		catch(ex)
		{
			/*ERROR HANDLING SECTION*/		
		}
		
	}

	/*******************************************************************************
	*	Function	:	initialize()
	*	Author		:	Mallesh Cheekoti
	*	Date		:	18/05/2004
	*	
	*	Purpose		:	This function loads and populates screen data (DC035.asp).
	*******************************************************************************/
	
	function initialize()
	{
		
		//*=[MC]This function populates data on DOCUMENT initialize/LOAD.
		populateScreen();
		
		//*=[MC] Populate "Who Has Seen Customer" Combo list
		PopulateCombo("VerWhoSeenCustomer",frmScreen.idWhoHasSeenCustomer,1);
		//*=[MC] Populate "How was customer ID seen" Combo list
		PopulateCombo("VerHowCustomerIDSeen",frmScreen.idHowWasCustomerIdSeen,1);
		<% /* SAB - EP16 */ %>
		PopulateCombo("KYCStatus",frmScreen.idKYCStatus,1);
		
		if(m_sCurrentCustomerNumber == "" &&  m_sCurrentCustomerVersionNumber == "")
		{
			m_sCurrentCustomerNumber = m_sCustomer1Number;
			m_sCurrentCustomerVersionNumber = m_sCustomer1Version;
		}
		
		//Get Customer and customer VErsion data
		GetCustomer();
		
		if(CustomerXML!=null)
		{
			setWhoSeenCustomer(CustomerXML);
			setHowCustomerIDSeen(CustomerXML);
			setKYCStatus(CustomerXML);
		}
		
		/*
			The following lines can be uncommented to test freeze or readonly functionality
			SET Readonly to "1" and/or FreezeIndicator to "1" to test respective features.
		*/
		//scScreenFunctions.SetContextParameter(window,"idFreezeDataIndicator","0");
		//scScreenFunctions.SetContextParameter(window,"idReadOnly","0");
		
		/*SET DC035 SCREEN READONLY IF CONDITIONS SATISFY*/
		m_sReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC035");
		
		if (m_sReadOnly=="1")
		{
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			frmScreen.btnAdd.disabled=true;
			frmScreen.btnDelete.disabled=true;
		}
	}

	/*******************************************************************************
	*	Function	:	setWhoSeenCustomer()
	*	Arguments	:	CustomerXML object
	*	Author		:	Mallesh Cheekoti
	*	Date		:	21/05/2004
	*	Purpose		:	This function sets WhoSeenCustomer combo item data
	*******************************************************************************/

	function setWhoSeenCustomer(cusVerXMLObject)
	{
		try
		{
			if(cusVerXMLObject!=null)
			{
				cusVerXMLObject.ActiveTag = null;
				cusVerXMLObject.SelectTag(null,"CUSTOMER");
				var nodTemp = cusVerXMLObject.GetTagText("WHOSEENCUSTOMER");
				if(nodTemp!=null && nodTemp!="" && nodTemp!="0")
				{	
					frmScreen.idWhoHasSeenCustomer.value=nodTemp;
				}
			}
			
		}
		catch(ex)
		{
			alert('Unexpected Error Occured' + ex);
		}
	}

	/*******************************************************************************
	*	Function	:	setHowCustomerIDSeen
	*	Arguments	:	CustomerXML object
	*	Author		:	Mallesh Cheekoti
	*	Date		:	21/05/2004
	*	Purpose		:	This function sets HowCustomerIDSeen combo item data
	*******************************************************************************/
	
	function setHowCustomerIDSeen(cusVerXMLObject)
	{
		try
		{
			if(cusVerXMLObject!=null)
			{
				cusVerXMLObject.ActiveTag = null;
				cusVerXMLObject.SelectTag(null,"CUSTOMER");
				
				var nodTemp = cusVerXMLObject.GetTagText("HOWCUSTOMERIDSEEN");
				if(nodTemp!=null && nodTemp!="" && nodTemp!="0")
				{	
					frmScreen.idHowWasCustomerIdSeen.value=nodTemp;
				}
			}
		}
		catch(ex)
		{
			alert('Unexpected Error Occured' + ex);
		}
	}

	/*******************************************************************************
	*	Function	:	setKYCStatus()
	*	Arguments	:	CustomerXML object
	*	Author		:	Simon Bajcer
	*	Date		:	30/03/2006
	*	Purpose		:	This function sets KYCStatus combo item data
	*******************************************************************************/

	function setKYCStatus(cusVerXMLObject)
	{
		try
		{
			if(cusVerXMLObject!=null)
			{
				cusVerXMLObject.ActiveTag = null;
				cusVerXMLObject.SelectTag(null,"CUSTOMER");
				var nodTemp = cusVerXMLObject.GetTagText("CUSTOMERKYCSTATUS");

				if(nodTemp!=null && nodTemp!="" && nodTemp!="0")
				{	
					frmScreen.idKYCStatus.value=nodTemp;
					if(scScreenFunctions.IsValidationType(frmScreen.idKYCStatus,"AE"))
					{
						frmScreen.idKYCStatus.disabled=true;
					}
					else
					{
						var redundantOption = nextAEOption(frmScreen.idKYCStatus);
						while (redundantOption > -1)
						{
							frmScreen.idKYCStatus.remove(redundantOption);
							redundantOption = nextAEOption(frmScreen.idKYCStatus);
						}
					}
				}
			}
		}
		catch(ex)
		{
			alert('Unexpected Error Occured' + ex);
		}
	}

	/*******************************************************************************
	*	Function	:	GetCustomer()
	*	Arguments	:	<None>
	*	Author		:	Mallesh Cheekoti
	*	Date		:	21/05/2004
	*	Purpose		:	This function loads customerVersion XML
	*******************************************************************************/
	function GetCustomer()
	{
		
		CustomerXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		with (CustomerXML)
		{
			CreateRequestTag(window, null)
			CreateActiveTag("CUSTOMERVERSION");
			CreateTag("CUSTOMERNUMBER", m_sCurrentCustomerNumber);
			CreateTag("CUSTOMERVERSIONNUMBER", m_sCurrentCustomerVersionNumber);
			RunASP(document, "GetPersonalDetails.asp");
			IsResponseOK();
			
		}
	}

	/**************************************************************
	*	Author	:	Mallesh Cheekoti
	*	CC		:	BMIDS747 / CC063 
	*	SUMMARY	:	VERIFICATION SUMMARY AND DETAIL SCREENS
	*
	*	DESC	:
	*	This function executes 'UpdateCustomerVersion.asp' page to
	*	update "WhoSeenCustomer" and "HowCustomerIDSeen" data to 
	*	CustomerVersion DB Table
	***************************************************************/

	function updateScreenData()
	{
		
		var bReturn=false;
		var whoSeenCustomer=0;
		var howCustomerIDSeen=0;
		
		if(frmScreen.idHowWasCustomerIdSeen.selectedIndex>0)
		{
			howCustomerIDSeen = frmScreen.idHowWasCustomerIdSeen.value;
		}
		
		if(frmScreen.idWhoHasSeenCustomer.selectedIndex>0)
		{
			whoSeenCustomer = frmScreen.idWhoHasSeenCustomer.value;
		}
		
		<% /* EP16 - Update CUSTOMERKYCSTATUS */ %>
		var KYCStatus="";
		if(frmScreen.idKYCStatus.selectedIndex>0)
		{
			KYCStatus = frmScreen.idKYCStatus.value;
		}
		<% /* EP16 */ %>

		var CustXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		CustXML.CreateRequestTag(window,null);

		CustXML.CreateActiveTag("CUSTOMERVERSION");
		CustXML.CreateTag("CUSTOMERNUMBER", m_sCurrentCustomerNumber);
		CustXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCurrentCustomerVersionNumber);
		
		CustXML.CreateTag("WHOSEENCUSTOMER",whoSeenCustomer);	
		CustXML.CreateTag("HOWCUSTOMERIDSEEN",howCustomerIDSeen);
		CustXML.CreateTag("CUSTOMERKYCSTATUS",KYCStatus);
				
		switch (ScreenRules())
		{
			case 1: // Warning
			case 0: // OK
				CustXML.RunASP(document, "UpdateCustomerVersion.asp");
				break;
			default: // Error
				CustXML.SetErrorResponse();
		}
		if(CustXML.IsResponseOK())
		{
			bReturn=true;
		}
		
		return bReturn;
	}

	/**************************************************************
	*	FUNCTION:	populateScreen()
	*	Author	:	Mallesh Cheekoti
	*	CC		:	BMIDS747 / CC063 
	*	DESC	:
	*	This function call populates data on to screens. Required data has been
	*	retrieved from combolists, VERIFICATION DB Table etc.,
	***************************************************************/
	function populateScreen()
	{
		
		/*=[MC] Retrieve Customer Name from the context and set the value to relevant field*/
		var sCustomerName = scScreenFunctions.GetContextCustomerName(window, m_sCurrentCustomerNumber);
		frmScreen.IdCustomerName.value = sCustomerName;
		
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"IdCustomerName");
		try
		{
			ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			ListXML.CreateRequestTag(window,null)
			ListXML.CreateActiveTag("VERIFICATION");
		
			<% /* BMIDS864 Remove ApplicationNumber and ApplicationFactFindNumber which are no longer keys 
			ListXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
			ListXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);  */ %>
			ListXML.CreateTag("CUSTOMERNUMBER", m_sCurrentCustomerNumber);
			try
			{
				ListXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCurrentCustomerVersionNumber);
			}
			catch(ex)
			{
			}	
			
		}
		catch(e)
		{
			//Error control
		}
		/*[MC] Run "FindVerificationList.asp" to retrieve required verification data to render onto screen*/	
		ListXML.RunASP(document, "FindVerificationList.asp");
		
		/*Check Response*/
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var sResponseArray = ListXML.CheckResponse(ErrorTypes);
		
		
		if (sResponseArray[0] == true || 
			sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
		{
			try
			{
				ListXML.ActiveTag = null;
				ListXML.CreateTagList("VERIFICATION");
			
				var m_iTotalRecords = ListXML.ActiveTagList.length;
				scVerificationListTable.initialiseTable(tblVerificationTable, 0, "", PopulateTable, 10, m_iTotalRecords);
				/*[MC]Set XMLArray*/
				SetUpXMLArray();
				//Setup paging DataSet Table onto screen
				PopulateTable(0);
			}
			catch(e)
			{
				/*ERROR HANDLING*/
				
				//alert(e);
			}
			
		}
		
	}


	function SetUpXMLArray()
	{
		<% 
			/*=[MC] BMIDS747 - CC063
				ELIMINATE DUPLICATE VERIFICATION RECORDS (SEQUENCE NUMBER IS UNIQUE FOR EACH CASE)
				DISPLAY ONLY UNIQUE DATA, IF INCASE SAME SEQUENCE NUMBER RECORD HAS BEEN SAVED AWAY 
			*/
		%>	
		
		ListXML.ActiveTag = null;
		ListXML.CreateTagList("VERIFICATION");
		try
		{
			var sGUID = "";
		
			for(var nIdx = 0; ListXML.SelectTagListItem(nIdx)!= false && nIdx < 10; nIdx++)
			{
				
				if(ListXML.GetTagText("VERIFICATIONSEQUENCENUMBER") != sGUID)
				{
					sGUID = ListXML.GetTagText("VERIFICATIONSEQUENCENUMBER");
					//alert(XMLArray.length);
					XMLArray[XMLArray.length] = nIdx;
					//alert(ListXML.GetTagText("VERIFICATIONSEQUENCENUMBER"))
				}
			}
		}
		catch(e)
		{
			/*ERROR HANDLING*/
			//alert(e + 'setupXMLArray');
		}
	}


	function PopulateTable(nStart)
	{
		scVerificationListTable.clear();	

		ListXML.ActiveTag = null;
		ListXML.CreateTagList("VERIFICATION");
		
		var verificationXML=null;
		verificationXML = getComboXML("VerVerificationType");
			
		<% /* Populate the table with a set of records, starting with record nStart */ %>
		for (var nLoop=0; ListXML.SelectTagListItem(nLoop + nStart) != false && nLoop < 10; nLoop++)
		{	
			try
			{
				var verificationSeqNo = ListXML.GetTagText("VERIFICATIONSEQUENCENUMBER");
				var verificationType =ListXML.GetTagText("VERIFICATIONTYPE");
				var idType	=  ListXML.GetTagText("IDENTIFICATIONTYPE");
				var issuer =ListXML.GetTagText("ISSUER");
				var reference = ListXML.GetTagText("REFERENCE");
				var verificationDate = ListXML.GetTagText("VERIFICATIONDATE");
				
				<% /* BMIDS864  Remove ApplicationNumber and ApplicationFactFindNumber
				var appNo =ListXML.GetTagText("APPLICATIONNUMBER");
				var appFactFindNo =ListXML.GetTagText("APPLICATIONFACTFINDNUMBER");  */ %>
				var custNo =ListXML.GetTagText("CUSTOMERNUMBER");
				var custVersionNo =ListXML.GetTagText("CUSTOMERVERSIONNUMBER");
				var reasonT1T3 = ListXML.GetTagText("OTHERIDENTIFICATION");
			
				var sVerificationDesc = getVerificationDescriptionByID(verificationXML,verificationType);
				var	sIDTypeListName = sVerificationDesc.replace(/\s+/,"");
				/*[MC] BMIDS747 - CC063
					IDTypeList combo names for verification list items are as follows
					Ver[VERIFICATION LIST ITEM DESC WITH OUT WHITE SPACES]IDType
					EX: Verification List Item "Electronic ID" IDType combo name is "VerElectronicIDIDType"
					
					Note:This format has been agreed with sophie barnes.
				*/
				sIDTypeListName = "Ver" + sIDTypeListName + "IDType";
			
				var IDTypeListXML = getComboXML(sIDTypeListName);
			
				var sIDTypeDesc = getVerificationDescriptionByID(IDTypeListXML,idType);
			
			
				scScreenFunctions.SizeTextToField(tblVerificationTable.rows(nLoop + 1).cells(0), sVerificationDesc);
				scScreenFunctions.SizeTextToField(tblVerificationTable.rows(nLoop + 1).cells(1), sIDTypeDesc);
				scScreenFunctions.SizeTextToField(tblVerificationTable.rows(nLoop + 1).cells(2), issuer);
				scScreenFunctions.SizeTextToField(tblVerificationTable.rows(nLoop + 1).cells(3), reference);
				scScreenFunctions.SizeTextToField(tblVerificationTable.rows(nLoop + 1).cells(4), verificationDate);

				//tblVerificationTable.rows(nLoop + 1).title = reasonT1T3;
				tblVerificationTable.rows(nLoop+1).cells(0).title= reasonT1T3;
				tblVerificationTable.rows(nLoop+1).cells(1).title= reasonT1T3;
				tblVerificationTable.rows(nLoop+1).cells(2).title= reasonT1T3;
				tblVerificationTable.rows(nLoop+1).cells(3).title= reasonT1T3;
				tblVerificationTable.rows(nLoop+1).cells(4).title= reasonT1T3;
				
			}
			catch(e)
			{
				/*ERROR HANDLING CODE HERE*/
				//alert(e);
			}
			
		}
		
		if (nLoop + nStart > 0)
		{
			scVerificationListTable.setRowSelected(1);
			frmScreen.btnEdit.disabled = false;
			frmScreen.btnDelete.disabled = false;
		}
		
		/*[MC]
			If no records found, disable Edit and Delete buttons
		*/
		if(ListXML.SelectTagListItem(nStart) == false)
		{
			frmScreen.btnEdit.disabled=true;
			frmScreen.btnDelete.disabled=true;
		}
	}


	function PopulateCombo(comboName,objScreenCombo,clearCombo)
	{
		var comboXML = null;
		var XML =null;
		var sGroupList = null;
		var blnSuccess = true;
		
		if(comboName==null || comboName=='')
		{
			return;
		}
		
		if(clearCombo==1)
		{
			while(objScreenCombo.options.length > 0) 
				objScreenCombo.remove(0);
		}
		
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		sGroupList = new Array(comboName);

		if(XML!=null)
		{
			if(XML.GetComboLists(document, sGroupList))
			{
				comboXML = XML.GetComboListXML(comboName);
				
				try
				{
					blnSuccess = blnSuccess & XML.PopulateComboFromXML(document, objScreenCombo,comboXML,true);
				}
				catch(ex)
				{
					//Error Occured set to false
					blnSuccess = false;
				}
				
				if(!blnSuccess)
					scScreenFunctions.SetScreenToReadOnly(frmScreen);
			}
		}

		XML = null;		
	}

	function RetrieveContextData()
	{
		m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
		m_sContext = scScreenFunctions.GetContextParameter(window,"idProcessContext",null);
		m_sReadOnly	= scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
		m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","APP0001");
		m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","2");
		m_sCustomer1Number = scScreenFunctions.GetContextParameter(window,"idCustomerNumber1",null);
		m_sCustomer1Version = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber1",null);
		m_sCustomer2Number = scScreenFunctions.GetContextParameter(window,"idCustomerNumber2",null);
		m_sCustomer2Version = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber2",null);
		m_sCustomer3Number = scScreenFunctions.GetContextParameter(window,"idCustomerNumber3",null);
		m_sCustomer3Version = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber3",null);
		m_sCustomer4Number = scScreenFunctions.GetContextParameter(window,"idCustomerNumber4",null);
		m_sCustomer4Version = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber4",null);
		m_sCustomer5Number = scScreenFunctions.GetContextParameter(window,"idCustomerNumber5",null);
		m_sCustomer5Version = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber5",null);
		m_sProcessInd = scScreenFunctions.GetContextParameter(window,"idProcessingIndicator",null); //JR BM0271

		m_sCurrentCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber","1007");
		m_sCurrentCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber","1");


		if(m_sMetaAction == "FromDC031" || m_sMetaAction == "FromDC033" || m_sContext == "CompletenessCheck")
		{
			//Get the current customer
			m_sCurrentCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber",null);
			m_sCurrentCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber",null);
		}	
	}


	function frmScreen.btnAdd.onclick()
	{
		if(updateScreenData())
		{
			/*SET METAACTION CONTEXT DATA TO ADD*/
			scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
			frmToDC031.submit();
		}
		else
		{
			alert('Fail to update screen data')
		}
	}	

	function frmScreen.btnEdit.onclick()
	{
		
		if(updateScreenData())
		{
			/*for edit action set METAACTION AND SEQUENCENUMBER*/
			ListXML.ActiveTag = null;
			ListXML.CreateTagList("VERIFICATION");
					
			<% /* Get the index of the selected row */ %>
			var nRowSelected = scVerificationListTable.getOffset() + scVerificationListTable.getRowSelected();

			if(ListXML.SelectTagListItem(XMLArray[nRowSelected-1]) == true)
			{
				scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
				scScreenFunctions.SetContextParameter(window,"idVerificationSeqNo",ListXML.GetTagText("VERIFICATIONSEQUENCENUMBER"));
				frmToDC031.submit();
				
			}
		}
		else
		{
			alert('Fail to update screen data');
		}
		
	}

	function btnSubmit.onclick()
	{
		if(updateScreenData())
		{
			frmToDC030.submit();
		}
	}

	function frmScreen.btnDelete.onclick()
	{
		var bConfirm = confirm("Would you like to delete a verification record?");
		var operationDelete = 2;
		
		if (bConfirm == true)
		{
			<% /* Get Selected Row Index */ %>
			var nRowSelected = scVerificationListTable.getOffset() + scVerificationListTable.getRowSelected();
			<%/* Set selected row as active node*/%>
			ListXML.SelectTagListItem(XMLArray[nRowSelected-1]);

			<% /* Set up the delete verification XML */%>
			
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			if(XML!=null || XML!='undefined')
			{
				XML.CreateRequestTag(window,operationDelete);
				XML.CreateActiveTag("VERIFICATION");
				
				<% /* BMIDS864  Remove ApplicationNumber and ApplicationFactFindNumber
				XML.CreateTag("APPLICATIONNUMBER",	m_sApplicationNumber);
				XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);   */ %>
				XML.CreateTag("CUSTOMERNUMBER", m_sCurrentCustomerNumber);
				XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCurrentCustomerVersionNumber);
				XML.CreateTag("VERIFICATIONSEQUENCENUMBER", ListXML.GetTagText("VERIFICATIONSEQUENCENUMBER"));
				XML.CreateTag("VERIFICATIONSEQUENCENUMBER", ListXML.GetTagText("VERIFICATIONTYPE"));
				XML.CreateTag("VERIFICATIONSEQUENCENUMBER", ListXML.GetTagText("IDENTIFICATIONTYPE"));
				XML.CreateTag("VERIFICATIONSEQUENCENUMBER", ListXML.GetTagText("ISSUER"));
				
				switch (ScreenRules())
				{
					case 1: 
						<%/*WARING SECTION*/%>
					case 0: 
						<%/*REQUEST DELETE OPERATION*/%>
						XML.RunASP(document, "DeleteVerification.asp");
						break;
					default: 
						<%/*ANY OTHER SPECIFIED CASES ARE TREAT AS ERRORS*/%>
						XML.SetErrorResponse();
				}

					
			<%/*IF RECORD DELETED SUCCESSFULLY, REMOVE TAG FROM XML OUTPUT HOLD ON THE SCREEN*/%>
			if(XML.IsResponseOK() == true)
			{					
				ListXML.RemoveActiveTag();
				scVerificationListTable.RowDeleted();
							
				<%/*IF THE LAST ENTRY WAS DELETED, ADJEST TABLE CURRENT ROW*/%>
				if(ListXML.ActiveTagList.length < nRowSelected)
				{
					nRowSelected = nRowSelected - 1;
				}
							
				<%/*IF STILL MORE ROWS EXIST IN THE TABLE, SELECT A ROW,IF NOT DISABLE APPROPRIATE ACTION BUTTONS*/%>
				if(nRowSelected > 0)
				{
					scVerificationListTable.setRowSelected(nRowSelected - scVerificationListTable.getOffset());
				}
				else
				{
					frmScreen.btnEdit.disabled = true;
					frmScreen.btnDelete.disabled = true;
					scScreenFunctions.SetFocusToFirstField(frmScreen);
				}
			}
				/*CLEAN UP*/
				XML = null;		
			}
		}
	}

	function getVerificationDescriptionByID(verificationXML,VerificationListID)
	{
		var sReturn = "";
		
		var bfound=false;
		
		//alert(verificationXML.xml);
		var msXMLDOM = new ActiveXObject("microsoft.XMLDOM");
		msXMLDOM.loadXML(verificationXML.xml);
		
		var colListEntry = msXMLDOM.selectNodes("LISTNAME/LISTENTRY");
		if(colListEntry!=null && colListEntry.length>0)
		{
		
			for(var i=0; i<colListEntry.length;i++)
			{
				if(colListEntry.item(i).selectSingleNode("VALUEID").text == VerificationListID)
				{
					sReturn = colListEntry.item(i).selectSingleNode("VALUENAME").text;
					return sReturn;	
				}
				
			}
		}
		
		return sReturn;
	}

	<% /* EP16 */ %>
	function nextAEOption (combo)
	{
		var targetOption = -1;
		var validationType;
		for (var i=0; i < combo.options.length; i++)
		{
			validationType = combo.options[i].getAttribute("ValidationType0");
			if (validationType != null && validationType == "AE")
			{
				targetOption = i;
				break;
			}	
			validationType = combo.options[i].getAttribute("ValidationType1");
			if (validationType != null && validationType == "AE")
			{
				targetOption = i;
				break;
			}
		}	
		return targetOption;
	}

</script>

</BODY>
</HTML>
