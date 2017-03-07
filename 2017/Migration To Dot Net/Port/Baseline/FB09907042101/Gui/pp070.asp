<%@ LANGUAGE="JSCRIPT" %>
<HTML>
<%
/*
Workfile:      pp070.asp
Copyright:     Copyright © 2006 Marlborough Stirling

Description:   Payee Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:
LH		02/10/2006	EP1141		Country is now returned as text and not as it's corresponding numeric code 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<HEAD>
<meta NAME="GENERATOR" Content="MSHTML 5.00.2014.210"  >
<LINK href="stylesheet.css" rel=STYLESHEET type=text/css>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
</HEAD>

<BODY>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Scriptlets */ %>
<span id="spnListScroll">
	<span style="LEFT: 310px; POSITION: absolute; TOP: 260px">
		<OBJECT data=scTableListScroll.asp id=scScrollTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 300px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>


<% /* FORMS */ %>
<form id="frmToPP100" method="post" action="PP100.asp" STYLE="DISPLAY: none"></form>
<form id="frmToMN060" method="post" action="MN060.asp" STYLE="DISPLAY: none"></form>
<form id="frmToMN070" method="post" action="MN070.asp" STYLE="DISPLAY: none"></form>
<form id="frmToPP050" method="post" action="PP050.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div style="HEIGHT: 230px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Payee list</strong>
		<% /* MV - 05/08/2002 - BMIDS0294 
		<font face="MS Sans Serif" size="1">
			<strong>Payee list</strong>
		</font> */ %>
	</span>
	<span id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 24px">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	
				<td width="20%" class="TableHead">Payee Name</td>	
				<td width="50%" class="TableHead">Payee Address</td>	
				<td width="20%" class="TableHead">Bank Account No.</td>
				<td class="TableHead">Notes</td></tr>
			<tr id="row01">		
				<td width="20%" class="TableTopLeft">&nbsp;</td>		
				<td width="50%" class="TableTopCenter">&nbsp;</td>		
				<td width="20%" class="TableTopCenter">&nbsp;</td>		
				<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		
				<td width="20%" class="TableLeft">&nbsp;</td>		
				<td width="50%" class="TableCenter">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		
				<td width="20%" class="TableLeft">&nbsp;</td>		
				<td width="50%" class="TableCenter">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		
				<td width="20%" class="TableLeft">&nbsp;</td>		
				<td width="50%" class="TableCenter">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		
				<td width="20%" class="TableLeft">&nbsp;</td>		
				<td width="50%" class="TableCenter">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		
				<td width="20%" class="TableLeft">&nbsp;</td>		
				<td width="50%" class="TableCenter">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		
				<td width="20%" class="TableLeft">&nbsp;</td>		
				<td width="50%" class="TableCenter">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		
				<td width="20%" class="TableLeft">&nbsp;</td>		
				<td width="50%" class="TableCenter">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		
				<td width="20%" class="TableLeft">&nbsp;</td>		
				<td width="50%" class="TableCenter">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		
				<td width="20%" class="TableBottomLeft">&nbsp;</td>		
				<td width="50%" class="TableBottomCenter">&nbsp;</td>		
				<td width="20%" class="TableBottomCenter">&nbsp;</td>		
				<td class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</span>
	
	<span id="spnButtons" style="LEFT: 4px; POSITION: absolute; TOP: 200px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
			<input id="btnAdd" value="Add" type="button" style="WIDTH: 70px" class="msgButton">
		</span>

		<span style="LEFT: 75px; POSITION: absolute; TOP: 0px">
			<input id="btnEdit" value="Edit" type="button" style="WIDTH: 70px" class="msgButton">
		</span> 
		
		<span style="LEFT: 150px; POSITION: absolute; TOP: 0px">
			<input id="btnDelete" value="Delete" type="button" style="WIDTH: 70px" class="msgButton">
		</span> 
				
	</span>
</div>
</form>

<% /* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 340px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/pp070Attribs.asp" -->

<%/* CODE */ %>
<script language="JScript">
<!--
var m_sApplicationNumber, m_sAFFNumber;
var m_iPayeeSetupRole, m_iUserRole ;

var m_sMetaAction ; 
var scScreenFunctions ;
var m_iTableLength = 10;
var PayeeHistoryXML = null;
var m_blnReadOnly = false;
var m_PaymentStatus = "";
var m_PaymentType = "";
var PaymentsForPayeeListXML = null;
var m_sReadOnly = ""; //BS BM0271 16/04/03

function RetrieveContextData()
{
<%	
 /*
	// TEST
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","C00078387");
	scScreenFunctions.SetContextParameter(window,"idRole","50");
	
	// END TEST
 */
%>	
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", "");
	m_sAFFNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", "");
	m_iUserRole = scScreenFunctions.GetContextParameter(window, "idRole", "0");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window, "idReadOnly", "0"); //BS BM0271 16/04/03
}

/* Events */

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Payee Details","PP070",scScreenFunctions)
	
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);
	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	RetrieveContextData();
	
	if(!AcceptedQuote())
	{
		alert("Unable to access Payees as there is no accepted quotation for this application.");
		frmToMN070.submit();
	}
	else
	{
		if(!GetParameterValues())
		{
			alert('Error retrieving setup roles');
			DisableEditButtons();
			return ;
		}
		else if(m_iUserRole < m_iPayeeSetupRole) DisableEditButtons();
	
		PopulateScreen();
		m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "PP070");
		//JR BM0271
		if(m_blnReadOnly) 
		{
			frmScreen.btnAdd.disabled = true;
			frmScreen.btnDelete.disabled = true;
		}
		//BMIDS0468
		if(scScreenFunctions.IsCancelDeclineStage(window))
		{
			frmScreen.btnAdd.disabled=true;
			frmScreen.btnEdit.disabled=true;
			frmScreen.btnDelete.disabled=true;
		}
		
		//End
	}
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function AcceptedQuote()
{
	var AppXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AppXML.CreateRequestTag(window, null)
	AppXML.CreateActiveTag("APPLICATION");
	AppXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AppXML.RunASP(document,"GetApplicationData.asp");
	if(AppXML.IsResponseOK())
	{
		AppXML.SelectTag(null, "APPLICATIONFACTFIND");
		if(AppXML.GetTagText("ACCEPTEDQUOTENUMBER")!= "")
			return true;
	}

	return false;
}

function frmScreen.btnAdd.onclick()
{
	if(m_iUserRole < m_iPayeeSetupRole)
	{
		window.alert("You do not have the authority to edit a payment.") ;
		return ;
	}
	
	// Set values in context(Add) and, call PP100.asp 
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Create");
	scScreenFunctions.SetContextParameter(window,"idXML","");
	<% /* PSC 15/11/2002 BMIDS00294 */ %>
	scScreenFunctions.SetContextParameter(window,"idPayeeReadOnly", "0");
	frmToPP100.submit();
}

function frmScreen.btnDelete.onclick()
{
	var iCurrentRow = scScrollTable.getRowSelected();
	var aPaymentsStatus = new Array();
	var sPayeeSeqNo ;
	
	if(iCurrentRow == -1)
	{
		window.alert("Select a payee for deletion.") ;
		return ;
	}

	if(m_iUserRole < m_iPayeeSetupRole)
	{
		window.alert("You do not have the authority to delete a payee.") ;
		return ;
	}
	
	sPayeeSeqNo = tblTable.rows(iCurrentRow).getAttribute("PayeeHistorySeqNo") ;
	aPaymentsStatus = CheckPaymentsExist(sPayeeSeqNo);
	if(aPaymentsStatus[0])
	{
		alert('Payee details cannot be deleted. Existing payments have been scheduled for this payee.');
		return ;
	}
	else
	{
		if(!aPaymentsStatus[1]) 
		{
			alert('Error retrieving payee details.');
			return ;
		}
	}
	<% /* Delete the Payee selected */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window, "DELETEPAYEEHISTORY")
	XML.CreateActiveTag("PAYEEHISTORY");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.SetAttribute("PAYEEHISTORYSEQNO", sPayeeSeqNo);
	
	// 	XML.RunASP(document,"PaymentProcessingRequest.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"PaymentProcessingRequest.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	
	if(!XML.IsResponseOK())	return false ;
	else PopulateScreen(); 
}

function frmScreen.btnEdit.onclick()
{
	var iTagListCount, sXML ;
	var aPaymentsStatus = new Array();
	var ValidationList = new Array(1);
	var iCurrentRow = scScrollTable.getRowSelected();
	var xmlNode, xmlAttribNode
	
	var sTotalAmount1 = "0";
	var sTotalAmount2 = "0"
	
	if(iCurrentRow == -1)
	{
		window.alert("Select a payee for editing.") ;
		return ;
	}
	
	if(m_iUserRole < m_iPayeeSetupRole)
	{
		window.alert("You do not have the authority to edit a payment.") ;
		return ;
	}
	
	aPaymentsStatus = CheckPaymentsExist(tblTable.rows(iCurrentRow).getAttribute("PayeeHistorySeqNo"));
	if(aPaymentsStatus[0])
	{
		
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
		scScreenFunctions.SetContextParameter(window,"idPayeeReadOnly", "1"); //read Only
		
		var sCondition = ".//DISBURSEMENTPAYMENT[@PAYMENTSTATUS='98']" ;
		var xmlNodeList = PaymentsForPayeeListXML.XMLDocument.selectNodes(sCondition) ;
		
		if(xmlNodeList.length > 0)
		{
			for( var iCount = 0 ; iCount < xmlNodeList.length ; iCount ++)
			{
				xmlNode = xmlNodeList(iCount) ;
				PaymentsForPayeeListXML.ActiveTag.removeChild(xmlNode);
			}
		}
		
		var xmlNodeList = PaymentsForPayeeListXML.XMLDocument.selectNodes("//DISBURSEMENTPAYMENT") ;
		if (xmlNodeList.length == 0) 
		{
			scScreenFunctions.SetContextParameter(window,"idPayeeReadOnly", "0"); //Enabled
		}
		else
		{
			for( var iCount = 0 ; iCount < xmlNodeList.length ; iCount ++)
			{
			
				xmlNode = xmlNodeList(iCount) ;
				
				if (xmlNode.attributes.getNamedItem("PAYMENTSTATUS") != null)
				{
					xmlAttribNode = xmlNode.attributes.getNamedItem("PAYMENTSTATUS") ;
					m_PaymentStatus = xmlAttribNode.value ;
					<% /* PSC 15/11/2002 BMIDS00294 */ %>
					var ValidationList = new Array("P","R","I","INP","IF");
					var m_bPaymentStatus = XML.IsInComboValidationList(document,"PaymentStatus",m_PaymentStatus,ValidationList)
					
					if (m_bPaymentStatus)
					{
						xmlAttribNode = xmlNode.attributes.getNamedItem("NETPAYMENTAMOUNT") ;
						sAmount = xmlAttribNode.value ;
						sTotalAmount1 = parseFloat(sTotalAmount1) + parseFloat(sAmount);
					}
							
				}
				
				if (xmlNode.attributes.getNamedItem("PAYMENTTYPE") != null)
				{
					xmlAttribNode = xmlNode.attributes.getNamedItem("PAYMENTTYPE") ;
					m_PaymentType = xmlAttribNode.value ;
					
					var ValidationList = new Array("N");
					var m_bPaymentType = XML.IsInComboValidationList(document,"PaymentType",m_PaymentType,ValidationList)
					
					if (m_bPaymentType) 
					{
						xmlAttribNode = xmlNode.attributes.getNamedItem("NETPAYMENTAMOUNT") ;
						sAmount = xmlAttribNode.value ;
						sTotalAmount2 = parseFloat(sTotalAmount2) + parseFloat(sAmount);
					}
				}
			}
			
			if (parseFloat(sTotalAmount1) == Math.abs(parseFloat(sTotalAmount2))) 
				scScreenFunctions.SetContextParameter(window,"idPayeeReadOnly", "0"); //Enabled
			else
				scScreenFunctions.SetContextParameter(window,"idPayeeReadOnly", "1"); //Disabled
			}
		}
		else
		{
			if(!aPaymentsStatus[1]) 
			{
				alert('Error retrieving payments details.');
				return ;
			}
			else
			{
				<% /* SYS4536 - Set context to editable. */ %>
				scScreenFunctions.SetContextParameter(window,"idPayeeReadOnly", "0");
			}
		}
	
		sXML = '<PAYEEHISTORYDATA><PAYEEHISTORYSEQNO>' +  tblTable.rows(iCurrentRow).getAttribute("PayeeHistorySeqNo");	
		sXML += '</PAYEEHISTORYSEQNO>'
		
		// Set values in context and call PP100.asp
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
		iTagListCount = tblTable.rows(iCurrentRow).getAttribute("TagListItemCount");
		PayeeHistoryXML.SelectTagListItem(iTagListCount);
	
		sXML += PayeeHistoryXML.ActiveTag.xml ;
		sXML += '</PAYEEHISTORYDATA>'
	
		scScreenFunctions.SetContextParameter(window,"idXML", sXML);
		frmToPP100.submit();
}

function spnTable.ondblclick()
{
	if (scScrollTable.getRowSelected()!= null) 
	{
		//MBIDS0468
		//Bug fixed
		if(!frmScreen.btnEdit.disabled)
		{
			frmScreen.btnEdit.onclick();
		}
		
	}
}

function btnSubmit.onclick()
{
	m_sReturnScreenId =  scScreenFunctions.GetContextParameter(window,"idReturnScreenId","");
	if ( m_sReturnScreenId == "PP050.asp") 
		frmToPP050.submit();
	else
		frmToMN070.submit();
}

<% /* Functions */ %>
function PopulateScreen()
{
	PayeeHistoryXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
	PayeeHistoryXML.CreateRequestTag(window,"FINDPAYEEHISTORYLIST");
	PayeeHistoryXML.CreateActiveTag("PAYEEHISTORY");
	PayeeHistoryXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	PayeeHistoryXML.SetAttribute("_COMBOLOOKUP_","1");
	// 	PayeeHistoryXML.RunASP(document,"PaymentProcessingRequest.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			PayeeHistoryXML.RunASP(document,"PaymentProcessingRequest.asp");
			break;
		default: // Error
			PayeeHistoryXML.SetErrorResponse();
		}

	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = PayeeHistoryXML.CheckResponse(ErrorTypes);
	<% /* PSC 15/11/2002 BMIDS00294 */ %>
	if(ErrorReturn[0] == true || ErrorReturn[1] == ErrorTypes[0] )
	{
		PopulateListbox(0);
		<% /* PSC 15/11/2002 BMIDS00294 */ %>
		EnableEditButtons();
	}
	else
	{	<% /* Disable Add and Edit buttons */ %>
		DisableEditButtons();
		return ;	
	}
}

function PopulateListbox(nStart)
{
	PayeeHistoryXML.CreateTagList("PAYEEHISTORY");
	var iNumberOfRows = PayeeHistoryXML.ActiveTagList.length;

	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfRows);
	ShowList(nStart);
	if (iNumberOfRows > 0) 
		scScrollTable.setRowSelected(1);	
		
}

function ShowList(nStart)
{
	var iCount;
	var sPayeeName, sAccountNo, sNotes, sPayeeHistorySeqNo, sNotes ;
	var sAddressLine ;

	function getAddressLine()
	{
		var sTemp, sAddress ;
		
		sTemp		= '' ;
		sAddress	= '';
		
		function AddToAddressLine(sTemp, bCommaRequired)
		{
			if(sTemp!='')
			{	
				if(sAddress == '') sAddress = sAddress + sTemp ; 
				else
				{
					if (bCommaRequired) sAddress = sAddress + ', ' + sTemp ; 
					else sAddress = sAddress + ' ' + sTemp ; 
				}
			}
		}

		AddToAddressLine(PayeeHistoryXML.GetTagAttribute("ADDRESS","BUILDINGORHOUSENUMBER"), false);
		AddToAddressLine(PayeeHistoryXML.GetTagAttribute("ADDRESS","BUILDINGORHOUSENAME"), false);		
		AddToAddressLine(PayeeHistoryXML.GetTagAttribute("ADDRESS","FLATNUMBER"), true);
		AddToAddressLine(PayeeHistoryXML.GetTagAttribute("ADDRESS","STREET"), false);
		AddToAddressLine(PayeeHistoryXML.GetTagAttribute("ADDRESS","TOWN"), true);
		AddToAddressLine(PayeeHistoryXML.GetTagAttribute("ADDRESS","DISTRICT"), true);
		AddToAddressLine(PayeeHistoryXML.GetTagAttribute("ADDRESS","COUNTY"), true);
		AddToAddressLine(PayeeHistoryXML.GetTagAttribute("ADDRESS","COUNTRY_TEXT"), true);
		AddToAddressLine(PayeeHistoryXML.GetTagAttribute("ADDRESS","POSTCODE"), true);	
	
		return sAddress ;
	}
	
	scScrollTable.clear();		
	for (iCount = 0; iCount < PayeeHistoryXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		PayeeHistoryXML.SelectTagListItem(iCount + nStart);

		sBankName			= PayeeHistoryXML.GetAttribute("BANKNAME");	 	
		sAccountNo			= PayeeHistoryXML.GetAttribute("ACCOUNTNUMBER");	 
		sNotes				= PayeeHistoryXML.GetAttribute("NOTES");	 	
		sPayeeHistorySeqNo	= PayeeHistoryXML.GetAttribute("PAYEEHISTORYSEQNO");

		sPayeeName			= PayeeHistoryXML.GetTagAttribute("THIRDPARTY","COMPANYNAME");	
		sAddressLine		= getAddressLine();
		
		// Add to the table
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0), sPayeeName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1), sAddressLine);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2), sAccountNo);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3), sNotes);
		
		tblTable.rows(iCount+1).setAttribute("PayeeHistorySeqNo", sPayeeHistorySeqNo);		
		tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount + nStart);
	}
}

function CheckPaymentsExist(sPayeeHistorySeqNo)
{	
	<% /*  Check whether any payment records exist for this payee. If so return false else true. */ %>
	PaymentsForPayeeListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var XMLActiveTag = PaymentsForPayeeListXML.CreateRequestTag(window,"FINDPAYMENTSFORPAYEELIST");
	
	PaymentsForPayeeListXML.CreateActiveTag("PAYEEHISTORY");
	PaymentsForPayeeListXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	PaymentsForPayeeListXML.SetAttribute("PAYEEHISTORYSEQNO", sPayeeHistorySeqNo);
	// 	PaymentsForPayeeListXML.RunASP(document, "PaymentProcessingRequest.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			PaymentsForPayeeListXML.RunASP(document, "PaymentProcessingRequest.asp");
			break;
		default: // Error
			PaymentsForPayeeListXML.SetErrorResponse();
		}

	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = PaymentsForPayeeListXML.CheckResponse(ErrorTypes);
	var aReturn	= new Array();
	
	if(ErrorReturn[0]) 
	{
		aReturn[0] = true ; 
		aReturn[1] = false ;  
		
		return aReturn ;
	}
	else
	{	
		<% /* some error occured */ %>
		aReturn[0] = false ;  
		
		if(ErrorReturn[1]==ErrorTypes[0]) 
			<% /*  error occured is 'RecordNotFound' */ %>
			aReturn[1] = true ; 
		else 
			<% /* error occured is not 'RecordNotFound' */ %>
			aReturn[1] = false ;  
		
		return aReturn ;
	}	
}

function GetParameterValues()
{
	<% /* This is a Function to retrieve the Global Parameter Values from database  */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var blnReturn = false;	
	
	<% /* Preparing XML Request string  */ %>
	var XMLActiveTag = XML.CreateRequestTag(window,"SEARCH");
	XML.CreateActiveTag("GLOBALPARAMETER");
	XML.CreateTag("NAME","PPROCPAYEESETUPROLE");
		
	XML.RunASP(document,"FindCurrentParameterList.asp");

	if (XML.IsResponseOK()== true)
	{	
		blnReturn = true;
		if(XML.SelectTag(null,"GLOBALPARAMETERLIST") != null)
		{
			XML.CreateTagList("GLOBALPARAMETER");
			m_iPayeeSetupRole = XML.GetTagText("AMOUNT");			
		}		
	}
	else blnReturn = false;

	XML = null; 
	return blnReturn;
}

function DisableEditButtons()
{
	frmScreen.btnAdd.disabled = true ;
	frmScreen.btnEdit.disabled = true ;
	frmScreen.btnDelete.disabled = true ;
}

<% /* PSC 15/11/2002 BMIDS00294 - Start */ %>
function spnTable.onclick()
{
	if(m_sReadOnly != "1") //BS BM0271 16/04/03
	{
		EnableEditButtons();
	}
	//If it is readonly enable edit button except  cancel and decline stages
	//bmids0468
	//In readonly mode all buttons should be disabled except edit
	if(m_blnReadOnly)
	{
		frmScreen.btnEdit.disabled=false;
		frmScreen.btnAdd.disabled=true;
		frmScreen.btnDelete.disabled=true;
	}
	if(scScreenFunctions.IsCancelDeclineStage(window))
	{
		//BMIDS0468
		DisableEditButtons();
	}
}

function EnableEditButtons()
{	
	<% /* Incorrect Role */ %>
	if(m_iUserRole < m_iPayeeSetupRole)
	{
		DisableEditButtons();
	}
	else
	{
		<% /* Nothing Selected in List box */ %>
		var iCurrentRow = scScrollTable.getRowSelected();
		
		if (iCurrentRow == -1)
		{
			frmScreen.btnEdit.disabled = true;
			frmScreen.btnDelete.disabled = true;
		}
		else
		{
			var aPaymentsStatus = new Array();
			var sPayeeSeqNo;
	
			frmScreen.btnEdit.disabled = false;
			frmScreen.btnDelete.disabled = false;
			frmScreen.btnAdd.disabled = false

			sPayeeSeqNo = tblTable.rows(iCurrentRow).getAttribute("PayeeHistorySeqNo") ;
			aPaymentsStatus = CheckPaymentsExist(sPayeeSeqNo);
			if(aPaymentsStatus[0])
			{
				frmScreen.btnDelete.disabled = true;
			}
		}
	}
}
<% /* PSC 15/11/2002 BMIDS00294 - End */ %>
-->
</script>
</BODY>
</HTML>



