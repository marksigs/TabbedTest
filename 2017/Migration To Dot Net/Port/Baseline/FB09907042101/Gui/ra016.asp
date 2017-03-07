<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      ra016.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Risk Assessment Review
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
INR	 23/11/2005	MAR645			Created from RA015
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Case Assessment Override/Review  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT data=scClientFunctions.asp height=1 id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scTable.htm height=1 id=scTable style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
<span style="LEFT: 16px; POSITION: absolute; TOP: 204px">
	<OBJECT data=scTableListScroll.asp id=scScrollTable style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
</span>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen"  style="VISIBILITY: VISIBLE" validate   ="onchange" mark>
<div id="divBackground" style="HEIGHT: 420px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 390px" class="msgGroup">

	<span style="LEFT: 4px; POSITION: absolute; TOP: 14px" class="msgLabel">
		UserId
		<span style="LEFT: 40px; POSITION: absolute; TOP: -3px">
			<input id="txtUser" maxlength="20" style="POSITION: absolute; WIDTH: 110px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 180px; POSITION: absolute; TOP: 14px" class="msgLabel">
		Password
		<span style="LEFT: 105px; POSITION: absolute; TOP: -3px">
			<input id="txtPassword" type=password maxlength="15" style="POSITION: absolute; WIDTH: 70px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 180px; POSITION: absolute; TOP: 40px" class="msgLabel">
		Approval Reference
		<span style="LEFT: 105px; POSITION: absolute; TOP: -3px">
			<input id="txtReference" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Override Reasons
	</span>
	<div id="divDispTable" style = "HEIGHT: 200px; LEFT: 4px; POSITION: absolute; TOP: 76px; WIDTH: 380px" class="msgGroup">
	<span id="spnTable">
		<table id="tblTable" width="380" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	
				<td width="15%" class="TableHead">User ID</td>
				<td width="30%" class="TableHead">Date/Time</td>
				<td width="55%" class="TableHead">Override Reason</td>
			</tr>
			<tr id="row01">		
				<td  class="TableTopLeft">&nbsp;</td>		
				<td  class="TableTopCenter">&nbsp;</td>			
				<td  class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row02">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>				
				<td  class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row03">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>				
				<td  class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row04">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>				
				<td  class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row05">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>				
				<td  class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row06">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>				
				<td  class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row07">		
				<td  class="TableBottomLeft">&nbsp;</td>		
				<td  class="TableBottomCenter">&nbsp;</td>			
				<td  class="TableBottomRight">&nbsp;</td>
			</tr>
		</table>
	</span>
	<span style="LEFT: 1px; POSITION: absolute; TOP: 145px" class="msgLabel">
		Override Reason
		<span style="LEFT: 0px; POSITION: absolute; TOP: 16px">
			<select id="cboOverrideReason" name="OverrideReason" style="WIDTH: 380px" class="msgCombo">
			</select>
		</span>
	</span>

	<span style="LEFT: 1px; POSITION: absolute; TOP: 200px" class="msgLabel">
		Override Reason Notes
		<span style="LEFT: 0px; POSITION: absolute; TOP: 16px">
			<textarea id="txtOverrideReasonNotes" rows="5" style="WIDTH: 380px; POSITION: absolute" class="msgTxt"></textarea>
			</select>
		</span>
	</span>

	<span style="LEFT: 1px; POSITION: absolute; TOP: 310px">
		<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
		
		<span style="POSITION: absolute; TOP: 0px"  >
			<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
	</span>
</div>
</form>

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/ra016attribs.asp" -->
<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sUserId = null;
var m_sUnitId = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sStageNumber = null;
var m_nRASequenceNumber=null;
var m_nRARuleNumber=null;
var m_XML=null;
var m_bReadOnly=null;
var m_XMLReasons=null;
var m_iRowTot = null;
var m_arrRequestAttributes = null;
var m_BaseNonPopupWindow = null;
var m_ComboXML = null;
var m_iTableLength = 7;
var m_recordType = null;
var m_sReasons = null;
var m_sCreditCheckGuid = null;
var m_sReasonCode = null;

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;

function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	m_sUserId= sArgArray[0];	
	m_sUnitId= sArgArray[1];	
	m_sApplicationNumber = sArgArray[2];
	m_sApplicationFactFindNumber =sArgArray[3];
	m_sStageNumber =sArgArray[4];
	m_nRASequenceNumber= sArgArray[5];
	m_nRARuleNumber= sArgArray[6];
	m_bReadOnly= sArgArray[7];
	m_arrRequestAttributes=sArgArray[8];
	m_recordType=sArgArray[10];
	m_sReasons=sArgArray[11];

	<% /* This screen receives almost all of its data directly */ %>
	m_XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_XML.LoadXML(sArgArray[9]);
	
	scScreenFunctions.ShowCollection(frmScreen);
	if (m_recordType == "CA")
		m_XML.SelectTag(null,"RISKASSESSMENTRULEOVERRIDE");
	else
	{
		m_XML.SelectTag(null,"CREDITCHECKREASONCODE");
		m_sCreditCheckGuid = m_XML.GetTagText("CREDITCHECKGUID")
		m_sReasonCode = m_XML.GetTagText("REASONCODE")
	}

	GetPreviousOverrideReasons();

	m_ComboXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	PopulateCombos();

	PopulateScreen();
	<% /* Masks will be set for screen input only */ %>
	Validation_Init();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
	
}

function frmScreen.btnOK.onclick ()
{
	if(frmScreen.onsubmit())
		{
			if (m_recordType == "CA")
			{
				CreateRuleOverride();
			}
			else
			{
				CreateCCRuleOverride();
			}
		}
}

function frmScreen.btnCancel.onclick ()
{
	window.close ();
}

function PopulateScreen()
{
	var xmlTag;
	var txtReferenceSelect;
	var createTagListSelect;
	var overrideReasonNotes;
	var overrideReason;
	var selUser;

	with (frmScreen)
	{
		if (m_recordType == "CA")
		{
			txtReferenceSelect = "RAOVERRIDEAPPROVALREFERENCE";
			createTagListSelect = "RESPONSE/RISKASSESSMENTRULEOVERRIDE";
			overrideReasonNotes = "RAOVERRIDEREASON";
			overrideReason = "RAOVERRIDEREASON";
			selUser = "USERID";
		}	
		else
		{
			txtReferenceSelect = "SMOVERRIDEAPPROVALREFERENCE";
			createTagListSelect = "CREDITCHECKREASONHISTORY/APPLICATIONCREDITCHECK/CREDITCHECK/CREDITCHECKREASONCODE"
			overrideReasonNotes = "SMOVERRIDEREASON";
			overrideReason = "SMOVERRIDEREASON";
			selUser = "OVERRIDEUSERID";
		}

		<%/*If there is a reference then it is a review, so set fields to read-only 
		    and use the stored userid. Otherwise it is an Override, so input is required
		  OVERRIDEUSERID */ %>
		if ( m_XML.GetTagText(selUser)!="" || m_bReadOnly) 
		{
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtUser");
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtPassword");
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtReference");
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtOverrideReasonNotes");
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboOverrideReason");
			txtUser.value = m_XML.GetTagText(selUser);
			frmScreen.txtReference.value = m_XMLReasons.GetTagText(txtReferenceSelect);
			frmScreen.btnOK.disabled =true;
			frmScreen.btnCancel.focus;
			m_XMLReasons.ActiveTag = null;
			m_XMLReasons.CreateTagList(createTagListSelect);
			var iOverrideReasons = m_XMLReasons.ActiveTagList.length;
			scScrollTable.initialiseTable(tblTable, 0, "", PopulateTable, m_iTableLength, iOverrideReasons); 
			PopulateTable(0);
		}
		else
		{
			txtUser.value = m_sUserId;
			<% /* Only set masks if the screen accepts input */ %>
			SetMasks();
			txtPassword.value = "";
			txtOverrideReasonNotes.value =m_XML.GetTagText(overrideReasonNotes);
			cboOverrideReason.value = m_XML.GetTagText(overrideReason);
		
			m_XMLReasons.ActiveTag = null;
			m_XMLReasons.CreateTagList(createTagListSelect);
			var iOverrideReasons = m_XMLReasons.ActiveTagList.length;
			scScrollTable.initialiseTable(tblTable, 0, "", PopulateTable, m_iTableLength, iOverrideReasons); 
			PopulateTable(0);
		}
	}
}

function GetPreviousOverrideReasons()
{

	if (m_recordType == "CA")
	{
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTagFromArray(m_arrRequestAttributes,"Search");

		XML.CreateActiveTag("RISKASSESSMENTRULEOVERRIDE");
		XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		XML.CreateTag("STAGENUMBER",m_sStageNumber);
		XML.CreateTag("RISKASSESSMENTSEQUENCENUMBER",m_nRASequenceNumber );
		XML.CreateTag("RARULENUMBER",m_nRARuleNumber );

		XML.RunASP(document,"GetPreviousOverrideReasons.asp");

		if(XML.IsResponseOK())
		{
			m_XMLReasons = XML;
		}
		else
		{
			m_XMLReasons = null;
		}
	}
	else
	{
		//Get all previous CreditCheckReasonCode records with the
		//m_sReasonCode we are interested in
		m_XMLReasons = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_XMLReasons.LoadXML(m_sReasons);
		m_XMLReasons.SelectTag(null,"CREDITCHECKREASONHISTORY");
	}

}

function CreateRuleOverride()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_arrRequestAttributes,"Create");
	XML.CreateActiveTag("RISKASSESSMENTRULEOVERRIDE");
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	XML.CreateTag("STAGENUMBER",m_sStageNumber);
	XML.CreateTag("RISKASSESSMENTSEQUENCENUMBER",m_nRASequenceNumber );
	XML.CreateTag("RARULENUMBER",m_nRARuleNumber );
	XML.CreateTag("RAOVERRIDEAPPROVALREFERENCE",frmScreen.txtReference.value  );
	XML.CreateTag("RAOVERRIDEREASON",frmScreen.txtOverrideReasonNotes.value );
	XML.CreateTag("RAOVERRIDEREASONCODE", frmScreen.cboOverrideReason.value );
	XML.CreateTag("USERID",frmScreen.txtUser.value );
	XML.CreateTag("USERPASSWORD",frmScreen.txtPassword.value );
	XML.CreateTag("UNITID",m_sUnitId );

	// 	XML.RunASP(document,"CreateRuleOverride.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"CreateRuleOverride.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}


	var ErrorTypes = new Array("NOTAUTHORISED");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)
		{
		window.close ();
		}
	else if (ErrorReturn[1] == ErrorTypes[0])
	{
		alert("User Id and password cannot be authenticated, please try again.");
	}
	else alert(ErrorReturn[2]);
	
}


function CreateCCRuleOverride()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_arrRequestAttributes,"Create");
	XML.CreateActiveTag("SMRULEOVERRIDE");
	XML.CreateTag("CREDITCHECKGUID",m_sCreditCheckGuid );
	XML.CreateTag("REASONCODE",m_sReasonCode );
	XML.CreateTag("USERID",frmScreen.txtUser.value );
	//XML.CreateTag("SMOVERRIDEREASONRESULT",frmScreen.txtOverrideReasonNotes.value );
	XML.CreateTag("SMOVERRIDEREASON",frmScreen.txtOverrideReasonNotes.value );
	XML.CreateTag("SMOVERRIDEAPPROVALREFERENCE",frmScreen.txtReference.value  );
	XML.CreateTag("SMOVERRIDEREASONCODE", frmScreen.cboOverrideReason.value );
	XML.CreateTag("USERPASSWORD",frmScreen.txtPassword.value );
	XML.CreateTag("UNITID",m_sUnitId );

	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"CreateCCRuleOverride.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}


	var ErrorTypes = new Array("NOTAUTHORISED");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)
		{
		window.close ();
		}
	else if (ErrorReturn[1] == ErrorTypes[0])
	{
		alert("User Id and password cannot be authenticated, please try again.");
	}
	else alert(ErrorReturn[2]);
	
}

function PopulateCombos()
{
	var sGroupList, groupName;
	if (m_recordType == "CA")
	{
		groupName = "RiskAssessmentOverrideReason"
	}
	else
	{
		groupName = "StrategyManagerOverrideReason";
	}
	sGroupList = new Array(groupName);
	if(m_ComboXML.GetComboLists(document,sGroupList))
	{
		m_ComboXML.PopulateCombo(document,frmScreen.cboOverrideReason,groupName,true);
	}

}

function PopulateTable(nStart)
{
	var iCount;
	var selDate, selReasonCode, selUser;
	var iOverrideReasons = m_XMLReasons.ActiveTagList.length;
	if (m_recordType == "CA")
	{
		selDate = "RAOVERRIDEDATETIME";
		selReasonCode = "RAOVERRIDEREASONCODE";
		selUser = "USERID";
	}
	else
	{
		selDate = "SMOVERRIDEDATETIME";
		selReasonCode = "SMOVERRIDEREASONCODE"
		selUser = "OVERRIDEUSERID";
	}

	for (iCount = 0; (iCount < iOverrideReasons) && (iCount < 7); iCount++)
	{
		m_XMLReasons.SelectTagListItem(iCount + nStart);
		
		var sUserID = m_XMLReasons.GetTagText(selUser);
		var sDate = m_XMLReasons.GetTagText(selDate);
		var sOverrideReason = m_XMLReasons.GetTagText(selReasonCode);
		
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),sUserID);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),sDate);
		
		m_ComboXML.SelectSingleNode("//RESPONSE/LIST/LISTNAME/LISTENTRY[VALUEID = '" + sOverrideReason + "']"); 
			if (m_ComboXML.ActiveTag != null) 
				{ 	
					scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),m_ComboXML.GetTagText("VALUENAME")); 	
				}
		tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount);
	}
}
//DB End

/*
DB BMIDS01065 - No longer needed.
DB BMIDS00895 - Populate table with each record when user clicks on review.
function PopulateReviewTable(nStart)
{
	var iCount;
	var iReviewOverrideReasons = m_XML.ActiveTagList.length;

	for (iCount = 0; (iCount < iReviewOverrideReasons) && (iCount < 7); iCount++)
	{
		m_XML.SelectTagListItem(iCount + nStart);
		
		var sUserID = m_XML.GetTagText("USERID");
		var sDate = m_XML.GetTagText("RAOVERRIDEDATETIME");
		var sReviewOverrideReason = m_XML.GetTagText("RAOVERRIDEREASONCODE");
		
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),sUserID);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),sDate);
		
		m_ComboXML.SelectSingleNode("//RESPONSE/LIST/LISTNAME/LISTENTRY[VALUEID = '" + sReviewOverrideReason + "']"); 
			if (m_ComboXML.ActiveTag != null) 
				{ 	
					scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),m_ComboXML.GetTagText("VALUENAME")); 	
				}
		tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount);
	}
}	
DB END
*/

function spnTable.onclick()
{
	PopulateApprovalReference();
}

function PopulateApprovalReference()
{
	var txtReferenceSelect;
	var createTagListSelect;
	var overrideReasonNotes;
	var overrideReason;

	if (m_recordType == "CA")
	{
		txtReferenceSelect = "RAOVERRIDEAPPROVALREFERENCE";
		createTagListSelect = "RESPONSE/RISKASSESSMENTRULEOVERRIDE";
		overrideReason = "RAOVERRIDEREASON";
	}	
	else
	{
		txtReferenceSelect = "SMOVERRIDEAPPROVALREFERENCE";
		createTagListSelect = "CREDITCHECKREASONHISTORY/APPLICATIONCREDITCHECK/CREDITCHECK/CREDITCHECKREASONCODE"
		overrideReason = "SMOVERRIDEREASON";
	}
		
	m_XMLReasons.ActiveTag = null;
	m_XMLReasons.CreateTagList(createTagListSelect);
	
    var nRowSelected = scScrollTable.getRowSelectedIndex();

	if(m_XMLReasons.SelectTagListItem(nRowSelected-1) == true)
	{
		frmScreen.txtReference.value = m_XMLReasons.GetTagText(txtReferenceSelect);
		frmScreen.txtOverrideReasonNotes.value =   m_XMLReasons.GetTagText(overrideReason);
	}
}

-->
</script>
</DIV>
</body>
</html>


