<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
	<% /* 
Workfile:      DC016.asp
Copyright:     Copyright © 

Description:   Packaging Valuation Split
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR		22/02/2007	EP2_1272 Created.
SR		08/03/2007	EP2_1832 - Added functionality to freee the screen.
SR		19/03/2007	EP2_1828 - Added functionality of Critical Data Check on saving
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
	<head>
		<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4" />
		<link href="stylesheet.css" rel="stylesheet" type="text/css" />
		<title></title>
	</head>
	<body>
		<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden"
			tabIndex="-1" type="text/x-scriptlet" width="1" viewastext>
		</object>
		<script src="validation.js" language="JScript"></script>
		<%/* FORMS */%>
		<form id="frmToDC010" method="post" action="DC010.asp" STYLE="DISPLAY: none">
		</form>
		<span id="spnToLastField" tabindex="0"></span>
		<span id="spnListScroll">
			<span style="LEFT: 300px; POSITION: absolute; TOP: 261px">
				<object data="scTableListScroll.asp" id="scScrollTable" style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px"
					tabindex="-1" type="text/x-scriptlet" viewastext>
				</object>
			</span>
		</span>
	<form id="frmScreen" mark validate="onchange" year4>
		<div id="divIntroducerFeeList" style="HEIGHT: 240px; LEFT: 9px; POSITION: absolute; TOP: 60px; WIDTH: 607px"
			class="msgGroup">
			<span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel">
				<strong>Introducer Fees</strong>
			</span>
			<div id="divIntroducerFeesTable" style="HEIGHT: 190px; LEFT: 7px; POSITION: absolute; TOP: 36px; WIDTH: 595px;">
				<span id="spnIntroducerFees" style="LEFT: 4px; POSITION: absolute; TOP: 4px">
					<table id="tblIntroducerFees" width="590" border="0" cellspacing="0" cellpadding="0" class="msgTable"
						style="LEFT: 0px; TOP: 1px">
						<tr id="rowTitles">
							<td width="40%" class="TableHead">Fee Type</td>
							<td width="20%" class="TableHead">Amount</td>
							<td width="20%" class="TableHead">Refund Amount</td>
							<td width="20%" class="TableHead">Rebate Amount</td>
						</tr>
						<tr id="row01">
							<td width="40%" class="TableTopLeft">&nbsp;</td>
							<td width="20%" class="TableTopCenter">&nbsp;</td>
							<td width="20%" class="TableTopCenter">&nbsp;</td>
							<td width="20%" class="TableTopRight">&nbsp;</td>
						</tr>
						<tr id="row02">
							<td width="40%" class="TableLeft">&nbsp;</td>
							<td width="20%" class="TableCenter">&nbsp;</td>
							<td width="20%" class="TableCenter">&nbsp;</td>
							<td width="20%" class="TableRight">&nbsp;</td>
						</tr>
						<tr id="row03">
							<td width="40%" class="TableLeft">&nbsp;</td>
							<td width="20%" class="TableCenter">&nbsp;</td>
							<td width="20%" class="TableCenter">&nbsp;</td>
							<td width="20%" class="TableRight">&nbsp;</td>
						</tr>
						<tr id="row04">
							<td width="40%" class="TableLeft">&nbsp;</td>
							<td width="20%" class="TableCenter">&nbsp;</td>
							<td width="20%" class="TableCenter">&nbsp;</td>
							<td width="20%" class="TableRight">&nbsp;</td>
						</tr>
						<tr id="row05">
							<td width="40%" class="TableLeft">&nbsp;</td>
							<td width="20%" class="TableCenter">&nbsp;</td>
							<td width="20%" class="TableCenter">&nbsp;</td>
							<td width="20%" class="TableRight">&nbsp;</td>
						</tr>
						<tr id="row06">
							<td width="40%" class="TableLeft">&nbsp;</td>
							<td width="20%" class="TableCenter">&nbsp;</td>
							<td width="20%" class="TableCenter">&nbsp;</td>
							<td width="20%" class="TableRight">&nbsp;</td>
						</tr>
						<tr id="row07">
							<td width="40%" class="TableLeft">&nbsp;</td>
							<td width="20%" class="TableCenter">&nbsp;</td>
							<td width="20%" class="TableCenter">&nbsp;</td>
							<td width="20%" class="TableRight">&nbsp;</td>
						</tr>
						<tr id="row08">
							<td width="40%" class="TableLeft">&nbsp;</td>
							<td width="20%" class="TableCenter">&nbsp;</td>
							<td width="20%" class="TableCenter">&nbsp;</td>
							<td width="20%" class="TableRight">&nbsp;</td>
						</tr>
						<tr id="row09">
							<td width="40%" class="TableLeft">&nbsp;</td>
							<td width="20%" class="TableCenter">&nbsp;</td>
							<td width="20%" class="TableCenter">&nbsp;</td>
							<td width="20%" class="TableRight">&nbsp;</td>
						</tr>
						<tr id="row10">
							<td width="40%" class="TableBottomLeft">&nbsp;</td>
							<td width="20%" class="TableBottomCenter">&nbsp;</td>
							<td width="20%" class="TableBottomCenter">&nbsp;</td>
							<td width="20%" class="TableBottomRight">&nbsp;</td>
						</tr>
					</table>
				</span>
				<span id="spnButtons" style="LEFT: 4px; POSITION: absolute; TOP: 165px">
					<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
						<input id="btnAddFee" value="Add" type="button" style="WIDTH: 60px" class="msgButton">
					</span>
					<span style="LEFT: 64px; POSITION: absolute; TOP: 0px">
						<input id="btnEditFee" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
					</span>
					<span style="LEFT: 128px; POSITION: absolute; TOP: 0px">
						<input id="btnDeleteFee" value="Delete" type="button" style="WIDTH: 60px" class="msgButton">
					</span>
				</span>
			</div>
		</div>
		<div id="divEditIntroducerFee" style="HEIGHT: 100px; LEFT: 9px; POSITION: absolute; TOP: 310px; WIDTH: 607px"
			class="msgGroup">
			<span style="LEFT:11px; POSITION: absolute; TOP: 10px; WIDTH: 590px">
				<span style="LEFT:0px;WIDTH:190px" class="msgLabel">Fee Type</span>
				<span style="LEFT:200px;WIDTH:125px" class="msgLabel">Fee Amount</span>
				<span style="LEFT:331px;WIDTH:125px" class="msgLabel">Refund Amount</span>
				<span style="LEFT:462px;WIDTH:125px" class="msgLabel">Rebate Amount</span>
			</span>
			<span style="LEFT:11px; POSITION: absolute; TOP: 30px; WIDTH: 590px; HEIGHT:28px">
				<span style="LEFT:0px; POSITION:absolute; TOP:3px">
					<select id="cboFeeType" style="WIDTH: 185px" class="msgCombo" menusafe="true"></select>
				</span>
				<span style="LEFT:195px; POSITION:absolute; TOP:3px"">
					<input id="txtFeeAmount" maxlength="7" style="LEFT:0px;WIDTH:100px" class="msgTxt">
				</span>
				<span style="LEFT:320px; POSITION:absolute; TOP:3px"">
					<input id="txtRefundAmount" maxlength="7" style="LEFT:0px;WIDTH:100px" class="msgTxt">
				</span>
				<span style="LEFT:450px; POSITION:absolute; TOP:3px"">
					<input id="txtRebateAmount" maxlength="7" style="LEFT:0px;WIDTH:100px" class="msgTxt">
				</span>
			</span>
			<span id="spnEditConfirm" style="LEFT: 4px; POSITION: absolute; TOP: 68px">
				<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
					<input id="btnConfirm" value="Confirm" type="button" style="WIDTH: 60px" class="msgButton">
				</span>
				<span style="LEFT: 64px; POSITION: absolute; TOP: 0px">
					<input id="btnUndo" value="Undo" type="button" style="WIDTH: 60px" class="msgButton">
				</span>
			</span>
		</div>
	</form>
		<div id="msgButtons" style="HEIGHT: 19px; LEFT: 12px; POSITION: absolute; TOP: 450px; WIDTH: 600px"><!-- #include FILE="msgButtons.asp" -->
		</div>
		<span id="spnToFirstField" tabindex="0"></span>
		<!-- #include FILE="fw030.asp" -->
		<!--  #include FILE="attribs/DC016attribs.asp" -->
		<script language="JScript" type="text/javascript">
<!--
var scScreenFunctions;
var scClientScreenFunctions;
var m_sContext = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sReadOnly = null;
var m_blnReadOnly = false;
var sButtonList = new Array("Submit","Cancel");
var m_IntroducerFeeXML = null;
var m_iTableLength = 10;
var m_EditSectionDirty = false;
var m_ScreenChanged = false;
var m_sEditSectionAction = "";
var m_xmlOneOffCostCombo = null;  <% /* stores all the combovalues of OneOffCost with validationType='I'  */ %>
var m_sPackagerValuationFeeType = "" ;
var m_sPackagerValuationAdminFeeType = "" ;
var m_sIdXML2= "";

function window.onload()
{
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	FW030SetTitles("Introducer Fees","DC016",scScreenFunctions);
	RetrieveContextData();
	
	SetMasks();
	Validation_Init();
	PopulateScreen();
	ShowMainButtons(sButtonList);
	if (m_sReadOnly == "1") scScreenFunctions.SetScreenToReadOnly(frmScreen);
	
	frmScreen.btnEditFee.disabled = true;
	frmScreen.btnDeleteFee.disabled = true;
	disableEditFeeSection();
	
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC016");
	if (m_blnReadOnly == true)
	{
		frmScreen.btnAddFee.disabled = true;	
		m_sReadOnly = "1"; 
	}
}

function RetrieveContextData()
{
	m_sContext = scScreenFunctions.GetContextParameter(window,"idProcessContext",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber"); 
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly");
	m_sIdXML2 = scScreenFunctions.GetContextParameter(window,"idXML2",null);
	scScreenFunctions.SetContextParameter(window,"idXML2", "");
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function PopulateScreen()
{
	if(getIntroducerFee()) populateIntroducerFee();
	if(! populateCombos()) return false;
}

function populateCombos()
{
	var XMLCombo = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("OneOffCost");
	
	if(XMLCombo.GetComboLists(document,sGroupList))
	{
		m_xmlOneOffCostCombo = removeNonIntroduceFeeTypes(XMLCombo.GetComboListXML("OneOffCost"));
		<% /* Get FeeType valueids for PackagerValuationFee and PackagerValuationAdminFee*/ %>
		var arrComboValueIds;
		arrComboValueIds=XMLCombo.GetComboIdsForValidation("OneOffCost", "TPV", m_xmlOneOffCostCombo, document);
		m_sPackagerValuationFeeType = arrComboValueIds[0];
		
		arrComboValueIds=XMLCombo.GetComboIdsForValidation("OneOffCost", "TPVA", m_xmlOneOffCostCombo, document);
		m_sPackagerValuationAdminFeeType = arrComboValueIds[0];
		
		<%/* populate the combo with filtered values */%>
		if(populateFeeTypeCombo(m_xmlOneOffCostCombo)) return true;
		else return false;
	}
	return false;
	
	function removeNonIntroduceFeeTypes(xmlOneOffCost)
	{
		var XMLTempDoc = new ActiveXObject("microsoft.XMLDOM");
		XMLTempDoc.loadXML(xmlOneOffCost.xml);
		var xmlListEntryList = XMLTempDoc.selectNodes("LISTNAME/LISTENTRY");
		var xmlListEntry;
		var xmlValidationList;
		var xmlValidation 
		var bValidationTypeFound = false ;

		for(var iCount=xmlListEntryList.length-1; iCount >=0; iCount--)
		{
			xmlListEntry = xmlListEntryList.item(iCount);
			xmlValidationList = xmlListEntry.selectNodes("VALIDATIONTYPELIST/VALIDATIONTYPE");
			bValidationTypeFound = false;
			for(var j=0; j < xmlValidationList.length && !bValidationTypeFound ; j++)
			{
				xmlValidation = xmlValidationList.item(j);
				if(xmlValidation.text == "I") bValidationTypeFound = true;
			}
			
			if(!bValidationTypeFound) xmlListEntry.parentNode.removeChild(xmlListEntry);	
		}		
		return XMLTempDoc;
	}
}

function populateFeeTypeCombo(xmlComboValues)
{	
	if(m_IntroducerFeeXML.PopulateComboFromXML(document, frmScreen.cboFeeType, xmlComboValues, true))
		return true ;
	else return false;
}

function getIntroducerFee()
{
	m_IntroducerFeeXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_IntroducerFeeXML.CreateRequestTag(window, null);
	m_IntroducerFeeXML.SetAttribute("ENTITY_REF", "INTRODUCERFEE");
	m_IntroducerFeeXML.SetAttribute("CRUD_OP", "READ");
	
	m_IntroducerFeeXML.CreateActiveTag("INTRODUCERFEE");
	m_IntroducerFeeXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	m_IntroducerFeeXML.SetAttribute("COMBOLOOKUP","Y");
	
	m_IntroducerFeeXML.RunASP(document,"omCRUDIf.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = m_IntroducerFeeXML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		return true ;
	}
	else return false;
}

function populateIntroducerFee()
{
	m_IntroducerFeeXML.CreateTagList("INTRODUCERFEE[not (@CRUD_OP) or not (@CRUD_OP='DELETE')]");
	var iNoOffees = m_IntroducerFeeXML.ActiveTagList.length;
	if(iNoOffees > 0)
	{
		scScrollTable.initialiseTable(tblIntroducerFees, 0, "", ShowList, m_iTableLength, iNoOffees);
		ShowList(0);
	}
}

function ShowList(nStart)
{
	var sFeeTypeText = "";
	var sFeeType = "";
	var sAmount = "";
	var sRefundAmount = "";
	var sRebateAmount = "";
	var sCRUD_OP = "";
	
	scScrollTable.clear();
	for (var iCount = 0; iCount < m_IntroducerFeeXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		m_IntroducerFeeXML.SelectTagListItem(iCount+nStart);
		sFeeTypeText = m_IntroducerFeeXML.GetAttribute("FEETYPE_TEXT");
		sFeeType = m_IntroducerFeeXML.GetAttribute("FEETYPE");
		sAmount = m_IntroducerFeeXML.GetAttribute("FEEAMOUNT");
		sRefundAmount =  m_IntroducerFeeXML.GetAttribute("REFUNDAMOUNT");
		sRebateAmount =  m_IntroducerFeeXML.GetAttribute("REBATEAMOUNT");
		sCRUD_OP = 	m_IntroducerFeeXML.GetAttribute("CRUD_OP");		
		
		scScreenFunctions.SizeTextToField(tblIntroducerFees.rows(iCount+1).cells(0),sFeeTypeText);
		scScreenFunctions.SizeTextToField(tblIntroducerFees.rows(iCount+1).cells(1),sAmount);
		scScreenFunctions.SizeTextToField(tblIntroducerFees.rows(iCount+1).cells(2),sRefundAmount);
		scScreenFunctions.SizeTextToField(tblIntroducerFees.rows(iCount+1).cells(3),sRebateAmount);
		tblIntroducerFees.rows(iCount+1).setAttribute("TagListItem", (iCount + nStart));
		tblIntroducerFees.rows(iCount+1).setAttribute("CRUD_OP", sCRUD_OP);
		tblIntroducerFees.rows(iCount+1).setAttribute("FEETYPE", sFeeType);	
		tblIntroducerFees.rows(iCount+1).setAttribute("FEEAMOUNT", sAmount);
		tblIntroducerFees.rows(iCount+1).setAttribute("REFUNDAMOUNT", sRefundAmount);
		tblIntroducerFees.rows(iCount+1).setAttribute("REBATEAMOUNT", sRebateAmount);
	}
}

function spnIntroducerFees.onclick()
{	
	if (scScrollTable.isTableDisabled()) return ;
	if (scScrollTable.getRowSelected() != -1 && m_blnReadOnly == false)
	{		
		frmScreen.btnEditFee.disabled = false ;
		frmScreen.btnDeleteFee.disabled = false ;		
	}
}

function frmScreen.cboFeeType.onchange()
{
	if(frmScreen.cboFeeType.selectedIndex != 0) frmScreen.txtFeeAmount.disabled = false;
	if(scScreenFunctions.IsOptionValidationType(frmScreen.cboFeeType, frmScreen.cboFeeType.selectedIndex, 'R'))
	{
		frmScreen.txtRefundAmount.disabled = false;
	}
	else
	{
		frmScreen.txtRefundAmount.value = "";
		frmScreen.txtRefundAmount.disabled = true;
	}
	if(scScreenFunctions.IsOptionValidationType(frmScreen.cboFeeType, frmScreen.cboFeeType.selectedIndex, 'REB'))
	{
		frmScreen.txtRebateAmount.disabled = false;
	}
	else
	{
		frmScreen.txtRebateAmount.value = "" ;
		frmScreen.txtRebateAmount.disabled = true;
	}
	
	setEditSectionDirty();
}

function frmScreen.txtFeeAmount.onchange()
{
	setEditSectionDirty();
}

function frmScreen.txtRefundAmount.onchange()
{
	var sRefundAmount = frmScreen.txtRefundAmount.value ;
	if(sRefundAmount != "")
	{
		if(parseInt(sRefundAmount) > parseInt(frmScreen.txtFeeAmount.value))
		{
			alert("Refund Amount cannot be greater than Fee Amount");
			return false;
		}
	}
	setEditSectionDirty();
}

function frmScreen.txtRebateAmount.onchange()
{
	var sRebateAmount = frmScreen.txtRebateAmount.value ;
	if(sRebateAmount != "")
	{
		if(parseInt(sRebateAmount) > parseInt(frmScreen.txtFeeAmount.value))
		{
			alert("Rebate Amount cannot be greater than Fee Amount");
			return false;
		}
	}
	setEditSectionDirty();	
}

function frmScreen.btnAddFee.onclick()
{
	m_sEditSectionAction = "Add";
	clearEditFeeSection();
	enableEditFeeSection();
	
	var XMLTempDoc = new ActiveXObject("microsoft.XMLDOM");
	XMLTempDoc.loadXML(m_xmlOneOffCostCombo.xml);
	buildFeeTypesForAddFee(XMLTempDoc);
	
	while(frmScreen.cboFeeType.options.length > 0) frmScreen.cboFeeType.remove(0); <% /* Clear existing values */ %>	
	
	populateFeeTypeCombo(XMLTempDoc);
	
	function buildFeeTypesForAddFee(XMLTempDoc)
	{ <% /* Build XML with all FeeTypes that can be added. The result is used to populate combo FeeType.
			The number of records, of the same FeeType, allowed to be added is defined as a validation type starting with '@'
		 */ %>
		var xmlListEntryList = XMLTempDoc.selectNodes("LISTNAME/LISTENTRY");
		var xmlListEntry ;
		var xmlValidationList, xmlValidation ;
		var sValidationType ;
		var bValidationTypeFound = false ;
		var iAllowedNoOfFeeRecords, sFeeType ;			
		
		for(var i= 0; i<xmlListEntryList.length; ++i)
		{
			xmlListEntry = xmlListEntryList.item(i);
			xmlValidationList = xmlListEntry.selectNodes("VALIDATIONTYPELIST/VALIDATIONTYPE");
			iAllowedNoOfFeeRecords = 0;
			sFeeType = "";
			bValidationTypeFound = false ;
			for(var j=0; j < xmlValidationList.length && !bValidationTypeFound ; j++)
			{
				xmlValidation = xmlValidationList.item(j);
				sValidationType = xmlValidation.text;
				sFeeType = xmlListEntry.selectSingleNode("VALUEID").text;
				if(sValidationType.substring(0,1) == "@")				
				{
					iAllowedNoOfFeeRecords = parseInt(sValidationType.substring(1));
					bValidationTypeFound = true ;					
				}				
			}
			if(!bValidationTypeFound) iAllowedNoOfFeeRecords = 1;
			else if(iAllowedNoOfFeeRecords ==0) iAllowedNoOfFeeRecords = 1;

			if(m_IntroducerFeeXML.XMLDocument.selectNodes("//INTRODUCERFEE[not (@CRUD_OP) or not (@CRUD_OP='DELETE')][@FEETYPE='"+ sFeeType + "']").length >= iAllowedNoOfFeeRecords)
			{
				xmlListEntry.parentNode.removeChild(xmlListEntry);				
			}
		}		
	}
}

function frmScreen.btnEditFee.onclick()
{
	m_sEditSectionAction = "Edit";
	
	while(frmScreen.cboFeeType.options.length > 0) frmScreen.cboFeeType.remove(0); <% /* Clear existing values */ %>	
	populateFeeTypeCombo(m_xmlOneOffCostCombo);
	
	var iCount	= scScrollTable.getRowSelected() ;
	frmScreen.cboFeeType.value = tblIntroducerFees.rows(iCount).getAttribute("FEETYPE");
	frmScreen.txtFeeAmount.value = tblIntroducerFees.rows(iCount).getAttribute("FEEAMOUNT");
	frmScreen.txtRefundAmount.value	= tblIntroducerFees.rows(iCount).getAttribute("REFUNDAMOUNT");
	frmScreen.txtRebateAmount.value	= tblIntroducerFees.rows(iCount).getAttribute("REBATEAMOUNT");	
	
	enableEditFeeSection();
	disabledFeeListSection();
	frmScreen.cboFeeType.disabled = true ;
}

function frmScreen.btnDeleteFee.onclick()
{
	if(confirm("Please confirm you wish to delete this fee type"))
	{
		var iCount = scScrollTable.getRowSelected();
		<% /* Assuming that ActiveTagList is not reset after populating the list box - select the corresponding node in XML */ %>
		var iRowId = tblIntroducerFees.rows(iCount).getAttribute("TagListItem"); 
		m_IntroducerFeeXML.SelectTagListItem(iRowId);
		var sCrudOp = m_IntroducerFeeXML.GetAttribute("CRUD_OP");
		if(sCrudOp =="CREATE")
		{  <% /* record was NOT commited to database- so remove the node from XML */ %>
			m_IntroducerFeeXML.ActiveTag.parentNode.removeChild (m_IntroducerFeeXML.ActiveTag);
		}
		else
		{ <% /* record was already commited to database- so set it for delete */ %>
			m_IntroducerFeeXML.SetAttribute("CRUD_OP", "DELETE");
		}
			
		m_ScreenChanged = true ;
		m_IntroducerFeeXML.ActiveTag = null ;
		scScrollTable.clear();
		populateIntroducerFee();
	}
	
	if(scScrollTable.getRowSelected() != -1) 
	{
		frmScreen.btnEditFee.disabled = false;
		frmScreen.btnDeleteFee.disabled = false;
	}
	else 
	{
		frmScreen.btnEditFee.disabled = true;
		frmScreen.btnDeleteFee.disabled = true;
	}
}

function frmScreen.btnUndo.onclick()
{
	disableEditFeeSection();
	setEditSectionClean();
	enableFeeListSection();
}

function thisFeeCanAdded()
{ <% /* If the current fee being added is 'Package Valuation Admin fee', only allow this if 'Packager Valuation Fee' 
		was already added to the case */ %>
	if(frmScreen.cboFeeType.value == m_sPackagerValuationAdminFeeType)
	{
		if(m_IntroducerFeeXML.XMLDocument.selectNodes("//INTRODUCERFEE[not (@CRUD_OP) or not (@CRUD_OP='DELETE')][@FEETYPE='"+ m_sPackagerValuationFeeType + "']").length == 0)
		{
			alert('You must have a Packager Valuation Fee before you can add a Packager Valuation Admin Fee to the application');
			return false
		}	
	}
	return true ;
}

function frmScreen.btnConfirm.onclick()
{
	if(m_sEditSectionAction == "Add" && !thisFeeCanAdded()) return false;
	
	if(m_sEditSectionAction == "Edit")
	{
		var iCount = scScrollTable.getRowSelected();
		<% /* Assuming that ActiveTagList is not reset after populating the list box - select the corresponding node in XML */ %>
		var iRowId = tblIntroducerFees.rows(iCount).getAttribute("TagListItem"); 
		m_IntroducerFeeXML.SelectTagListItem(iRowId);
		m_IntroducerFeeXML.SetAttribute("FEEAMOUNT", frmScreen.txtFeeAmount.value);
		m_IntroducerFeeXML.SetAttribute("REFUNDAMOUNT", frmScreen.txtRefundAmount.value);
		m_IntroducerFeeXML.SetAttribute("REBATEAMOUNT", frmScreen.txtRebateAmount.value);
		<% /* Set the record for update unless this is newly created and not commited yet.  */ %>
		var sCurdOp = m_IntroducerFeeXML.GetAttribute("CRUD_OP");
		if(sCurdOp != "CREATE") m_IntroducerFeeXML.SetAttribute("CRUD_OP", "UPDATE");
	}
	else
	{ <% /* Add mode - add a new row in XML  */ %>
		//m_IntroducerFeeXML.ActiveTag = null ;
		m_IntroducerFeeXML.SelectTag(null, "RESPONSE");
		m_IntroducerFeeXML.CreateActiveTag("INTRODUCERFEE");
		m_IntroducerFeeXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		m_IntroducerFeeXML.SetAttribute("FEETYPE", frmScreen.cboFeeType.value);
		m_IntroducerFeeXML.SetAttribute("FEETYPE_TEXT", frmScreen.cboFeeType.options.item(frmScreen.cboFeeType.selectedIndex).text);		
		m_IntroducerFeeXML.SetAttribute("FEEAMOUNT", frmScreen.txtFeeAmount.value);
		m_IntroducerFeeXML.SetAttribute("REFUNDAMOUNT", frmScreen.txtRefundAmount.value);
		m_IntroducerFeeXML.SetAttribute("REBATEAMOUNT", frmScreen.txtRebateAmount.value);
		
		m_IntroducerFeeXML.SetAttribute("CRUD_OP", "CREATE");
	}

	disableEditFeeSection();
	setEditSectionClean();
	m_ScreenChanged = true ;
	
	m_IntroducerFeeXML.ActiveTag = null ;
	populateIntroducerFee();
	enableFeeListSection();
}

function btnSubmit.onclick()
{   <% /* SR 19/03/2007 : EP2_1828 - Added critical data check on saving */	 %>
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "FromDC016");
	scScreenFunctions.SetContextParameter(window,"idXML2", m_sIdXML2);
	if(m_ScreenChanged)
	{
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window, null);
		XML.SetAttribute("OPERATION","CriticalDataCheck");
		XML.SetAttribute("ENTITY_REF", "INTRODUCERFEE");
		XML.SetAttribute("CRUD_OP", "iUpdate");	
		XML.AddXMLBlock(m_IntroducerFeeXML.SelectTag(null, "RESPONSE"));
		
		XML.SelectTag(null,"REQUEST");
		XML.CreateActiveTag("CRITICALDATACONTEXT");
		XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		XML.SetAttribute("SOURCEAPPLICATION","Omiga");
		XML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
		XML.SetAttribute("ACTIVITYINSTANCE","1");
		XML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
		XML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
		XML.SetAttribute("COMPONENT","omCRUD.omCRUDBO");
		XML.SetAttribute("METHOD","OmRequest");

		window.status = "Critical Data Check - please wait";		
		XML.RunASP(document,"OmigaTMBO.asp");
		window.status = "";
		
		if (!XML.IsResponseOK()) return false ;
	}
	
	XML = null;
	if(m_sContext == "CompletenessCheck") frmToGN300.submit();
	else frmToDC010.submit();
}

function btnCancel.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "FromDC016");
	scScreenFunctions.SetContextParameter(window,"idXML2", m_sIdXML2);
	frmToDC010.submit();
}

function setEditSectionDirty()
{
	m_EditSectionDirty = true ;
	frmScreen.btnConfirm.disabled = false;
}

function setEditSectionClean()
{
	m_EditSectionDirty = false ;
}

function disabledFeeListSection()
{
	frmScreen.btnEditFee.disabled = true;
	frmScreen.btnAddFee.disabled = true;
	frmScreen.btnDeleteFee.disabled = true;
	scScrollTable.DisableTable();
	DisableMainButton("Submit");
}

function enableFeeListSection()
{
	scScrollTable.EnableTable();
	frmScreen.btnAddFee.disabled = false;
	
	if(scScrollTable.getRowSelected() != -1) 
	{
		frmScreen.btnEditFee.disabled = false;
		frmScreen.btnDeleteFee.disabled = false;
	}
	else 
	{
		frmScreen.btnEditFee.disabled = true;
		frmScreen.btnDeleteFee.disabled = true;
	}
	EnableMainButton("Submit");
}

function enableEditFeeSection()
{
	scScreenFunctions.EnableCollection(divEditIntroducerFee);
	disabledFeeListSection();
	if(frmScreen.cboFeeType.selectedIndex == 0) 
	{
		frmScreen.txtFeeAmount.disabled = true;
		frmScreen.txtRefundAmount.disabled = true;
		frmScreen.txtRebateAmount.disabled = true;
	}
	else
	{
		frmScreen.txtFeeAmount.disabled = false;
		if(scScreenFunctions.IsOptionValidationType(frmScreen.cboFeeType, frmScreen.cboFeeType.selectedIndex, 'R'))
		{
			frmScreen.txtRefundAmount.disabled = false;
		}
		else frmScreen.txtRefundAmount.disabled = true;
		
		if(scScreenFunctions.IsOptionValidationType(frmScreen.cboFeeType, frmScreen.cboFeeType.selectedIndex, 'REB'))
		{
			frmScreen.txtRebateAmount.disabled = false;
		}
		else frmScreen.txtRebateAmount.disabled = true; 
	}
	frmScreen.btnUndo.disabled = false ;
}

function disableEditFeeSection()
{
	scScreenFunctions.DisableCollection(divEditIntroducerFee);	
	frmScreen.btnConfirm.disabled = true ;
	frmScreen.btnUndo.disabled = true ;
}
function clearEditFeeSection()
{
	frmScreen.cboFeeType.value = "" ;
	frmScreen.txtFeeAmount.value = "";
	frmScreen.txtRefundAmount.value = "";
	frmScreen.txtRebateAmount.value = "";
}


-->
		</script>
	</body>
</html>
