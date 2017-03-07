<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      ra015.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Risk Assessment Review
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MH	10/04/00	Created
BG		17/05/00	SYS0752 Removed Tooltips
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

BMIDS History

Prog	Date		Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
DB		24/10/02	BMIDS00712 Added new combo and table for Override Reason.
DB		11/11/02	BMIDS00895 Changed max length of approval ref + bug fix from above.
DB		25/11/02	BMIDS01065 Further change to new override table.
MV		29/01/2003	BM0302	Added the ReasonNotes TextBox 
MC		20/04/2004	BMIDS517	white space padded to the title text
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
MF		01/08/2005	MAR019		Change title text
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
	
	<!-- DB BMIDS00712 Replaced with a table and combo box.
	<span style="TOP: 66px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Previous override reasons
		<span style="TOP: 16px; LEFT: 0px; POSITION: ABSOLUTE">
			<textarea id="txtPreviousOverridereasons" rows="5" style="WIDTH: 380px; POSITION: absolute" class="msgTxt"></textarea>
		</span>
	</span>

	<span style="TOP: 164px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Override reason
		<span style="TOP: 16px; LEFT: 0px; POSITION: ABSOLUTE">
			<textarea id="txtOverridereason" rows="5" style="WIDTH: 380px; POSITION: absolute" class="msgTxt"></textarea>
		</span>
	</span>
	DB End-->
	
	<!--DB BMIDS00712 - New table added to display override reasons -->
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
	
	<!--DB BMIDS00712 - New combo added to display override reason -->
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
<!-- #include FILE="attribs/ra015attribs.asp" -->
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

	<% /* This screen receives almost all of its data directly */ %>
	m_XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_XML.LoadXML(sArgArray[9]);
	m_XML.SelectTag(null,"RISKASSESSMENTRULEOVERRIDE");
	
	scScreenFunctions.ShowCollection(frmScreen);
	GetPreviousOverrideReasons();
	
	//DB BMIDS00712
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
		CreateRuleOverride();
		}
}

function frmScreen.btnCancel.onclick ()
{
	window.close ();
}

function PopulateScreen()
{
	var xmlTag;
	
	with (frmScreen)
	{
		//scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtPreviousOverridereasons");
		
		<%/*If there is a reference then it is a review, so set fields to read-only 
		    and use the stored userid. Otherwise it is an Override, so input is required
		  */ %>
		if ( m_XML.GetTagText("USERID")!="" || m_bReadOnly) 
		{
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtUser");
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtPassword");
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtReference");
			//DB BMIDS00712 - changed from a textbox to a combo.
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtOverrideReasonNotes");
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboOverrideReason");
			txtUser.value = m_XML.GetTagText("USERID");
			frmScreen.txtReference.value = m_XMLReasons.GetTagText("RAOVERRIDEAPPROVALREFERENCE");
			frmScreen.btnOK.disabled =true;
			frmScreen.btnCancel.focus;
			//DB BMIDS00895 - Populate the table when user clicks on review.
			//DB BMIDS01065 - No longer need this xml or the populatereviewtable function
			//m_XML.ActiveTag = null;
			//m_XML.CreateTagList("RISKASSESSMENTRULE/RISKASSESSMENTRULEOVERRIDE");
			//var iReviewOverrideReasons = m_XML.ActiveTagList.length;
			//scScrollTable.initialiseTable(tblTable, 0, "", PopulateReviewTable, m_iTableLength, iReviewOverrideReasons); 
			//PopulateReviewTable(0);

			m_XMLReasons.ActiveTag = null;
			m_XMLReasons.CreateTagList("RESPONSE/RISKASSESSMENTRULEOVERRIDE");
			var iOverrideReasons = m_XMLReasons.ActiveTagList.length;
			scScrollTable.initialiseTable(tblTable, 0, "", PopulateTable, m_iTableLength, iOverrideReasons); 
			PopulateTable(0);
			//DB END
		}
		else
		{
			txtUser.value = m_sUserId;
			<% /* Only set masks if the screen accepts input */ %>
			SetMasks();
			txtPassword.value = "";
			//txtReference.value = m_XML.GetTagText("RAOVERRIDEAPPROVALREFERENCE");
			//txtPreviousOverridereasons.value=m_XMLReasons.GetTagText("PREVIOUSOVERRIDEREASONS");
			//DB BMIDS00712 - Changed to a combo
			txtOverrideReasonNotes.value =m_XML.GetTagText("RAOVERRIDEREASON");
			cboOverrideReason.value = m_XML.GetTagText("RAOVERRIDEREASON");
		
			//DB BMIDS00712
			m_XMLReasons.ActiveTag = null;
			m_XMLReasons.CreateTagList("RESPONSE/RISKASSESSMENTRULEOVERRIDE");
			var iOverrideReasons = m_XMLReasons.ActiveTagList.length;
			scScrollTable.initialiseTable(tblTable, 0, "", PopulateTable, m_iTableLength, iOverrideReasons); 
			PopulateTable(0);
		}
	}
}

function GetPreviousOverrideReasons()
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
	//DB BMIDS00712 - changed textbox to combo.
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

//DB BMIDS00712 - New method added to populate combo.
function PopulateCombos()
{
	var sGroupList = new Array("RiskAssessmentOverrideReason");
	if(m_ComboXML.GetComboLists(document,sGroupList))
	{
		m_ComboXML.PopulateCombo(document,frmScreen.cboOverrideReason,"RiskAssessmentOverrideReason",true);
	}
}
//DB End

//DB BMIDS00712 - New method added to populate table.
function PopulateTable(nStart)
{
	var iCount;
	var iOverrideReasons = m_XMLReasons.ActiveTagList.length;

	for (iCount = 0; (iCount < iOverrideReasons) && (iCount < 7); iCount++)
	{
		m_XMLReasons.SelectTagListItem(iCount + nStart);
		
		var sUserID = m_XMLReasons.GetTagText("USERID");
		var sDate = m_XMLReasons.GetTagText("RAOVERRIDEDATETIME");
		var sOverrideReason = m_XMLReasons.GetTagText("RAOVERRIDEREASONCODE");
		
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

//DB BMIDS01065 - Added to populate the approval reference field when an override reason is clicked.
function spnTable.onclick()
{
	PopulateApprovalReference();
}

function PopulateApprovalReference()
{
	m_XMLReasons.ActiveTag = null;
	m_XMLReasons.CreateTagList("RESPONSE/RISKASSESSMENTRULEOVERRIDE");
	
    var nRowSelected = scScrollTable.getRowSelectedIndex();

	if(m_XMLReasons.SelectTagListItem(nRowSelected-1) == true)
	{
		frmScreen.txtReference.value = m_XMLReasons.GetTagText("RAOVERRIDEAPPROVALREFERENCE");
		frmScreen.txtOverrideReasonNotes.value =   m_XMLReasons.GetTagText("RAOVERRIDEREASON");
	}
}
//DB END

-->
</script>
</DIV>
</body>
</html>


