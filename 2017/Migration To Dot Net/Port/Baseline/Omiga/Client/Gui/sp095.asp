<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      SP095.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Cheque Reprint Payment Reprocessing 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AW		18/07/2002	BM029 Sceen created
GD		11/10/2002	BMIDS00514 - Fix to : 'Reprint cheques fails when more than 10 selected'.
GD		07/11/2002	BMIDS00681 - Do btnConfirm processing, when the 'X' is clicked to close the window.
GD		08/11/2002	BMIDS00559 - Only enable btnReprint if 1 or more reprints are selected.
GD		12/11/2002	BMIDS00869 - Allow double-click to toggle reprint status.
GD		21/11/2002	BMIDS01045
*/

%>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Reprint Cheques <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>
<% /* Scriptlets - remove any which are not required */ %>
<script src="validation.js" language="JScript"></script>

<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scTable.htm height=1 id=scTable 
style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 type=text/x-scriptlet 
VIEWASTEXT></OBJECT>
<span id="spnTableListScroll" style="LEFT: 290px; POSITION: absolute; TOP: 240px; VISIBILITY: hidden">
<OBJECT data=scTableListScroll.asp id=scTableList 
style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 
type=text/x-scriptlet VIEWASTEXT></OBJECT>
</span>

<% /* Specify Forms Here */ %>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
<% /* Start of form */ %>

<form id="frmScreen"  STYLE="VISIBILITY: hidden" year4 mark validate ="onchange">
	<div id="divBackground" style="HEIGHT: 400px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 604px" class="msgGroup">
		<span id="spnPaymentSanctionList" style="LEFT: 4px; POSITION: absolute; TOP: 40px">
			<table id="tblReprintCheques" width="580" border="0" cellspacing="0" cellpadding="0" class="msgTable">				

				<tr id="rowTitles"><td width="18%" class="TableHead">Application Number</td><td width="15%" class="TableHead">Payment Sequence Number</td><td width="18%" class="TableHead">Cheque Number</td><td width="20%" class="TableHead">Payment Amount</td><td width="25%" class="TableHead">Reprint Yes/No</td></tr>				
				<tr id="row01"><td width="18%" class="TableTopLeft">&nbsp;</td><td width="15%" class="TableTopCenter">&nbsp;</td><td width="18%" class="TableTopCenter">&nbsp;</td><td width="20%" class="TableTopCenter">&nbsp;</td><td width="25%" class="TableTopRight">&nbsp;</td></tr>
				<tr id="row02"><td width="18%" class="TableLeft">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="18%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="25%" class="TableRight">&nbsp;</td></tr>
				<tr id="row03"><td width="18%" class="TableLeft">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="18%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="25%" class="TableRight">&nbsp;</td></tr>
				<tr id="row04"><td width="18%" class="TableLeft">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="18%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="25%" class="TableRight">&nbsp;</td></tr>
				<tr id="row05"><td width="18%" class="TableLeft">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="18%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="25%" class="TableRight">&nbsp;</td></tr>
				<tr id="row06"><td width="18%" class="TableLeft">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="18%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="25%" class="TableRight">&nbsp;</td></tr>
				<tr id="row07"><td width="18%" class="TableLeft">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="18%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="25%" class="TableRight">&nbsp;</td></tr>
				<tr id="row08"><td width="18%" class="TableLeft">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="18%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="25%" class="TableRight">&nbsp;</td></tr>
				<tr id="row09"><td width="18%" class="TableLeft">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="18%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="25%" class="TableRight">&nbsp;</td></tr>
				<tr id="row10"><td width="18%" class="TableBottomLeft">&nbsp;</td><td width="15%" class="TableBottomCenter">&nbsp;</td><td width="18%" class="TableBottomCenter">&nbsp;</td><td width="20%" class="TableBottomCenter">&nbsp;</td><td width="25%" class="TableBottomRight">&nbsp;</td></tr>
				<!---->

			</table>
		</span>
		<span id="spnInput" style="FONT-WEIGHT: normal; LEFT: 10px; POSITION: absolute; TOP: 260px" class="msgLabel">
			</span><!-- Buttons -->
			<span style="LEFT: 240px; POSITION: absolute; TOP: 300px">
				<input id="btnReprintStatus" value="Alter Reprint Status" type="button" 
					style="WIDTH: 130px" class="msgButton">
			</span>
			<span id="spnButtons" style="LEFT: 40px; POSITION: absolute; TOP: 355px">
				<span style="LEFT: 1px; POSITION: absolute; TOP: 0px">
					<input id="btnConfirm" value="Confirm" type="button" 
						style="WIDTH: 90px" class="msgButton">
				
			<span style="HEIGHT: 20px; LEFT: 130px; POSITION: absolute; TOP: 0px; WIDTH: 260px" class="msgLabel">
					Click Confirm button if all cheques have been printed successfully, otherwise indicate those which need re-printing and click reprint button
			</span>
			<span style="LEFT: 420px; POSITION: absolute; TOP: 0px">
				<input id="btnReprint" value="Reprint" type="button" 
					style="WIDTH: 100px" class="msgButton">
			</span>
		</span>
	</div>
</form>


<% /* End of form */ %>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>
<% /* Specify Code Here */ %>
<script language="JScript">
<!--

var m_sMetaAction = null;
//GD Temp DEBUG
//var m_iTableLength = 10;
var m_iTableLength = 10;
var m_iTotalRecords = 0;
var XMLReprint = null;
var scScreenFunctions;
var m_sReadOnly = null;
var m_sContext;
var m_SelectedPaymentMethods;
var m_ProcessChequesFlag;
var m_SelectedRows = null;
var m_LastRowSelected = null;
var m_nAmount = 0;
var m_nRows = 0;
var	m_ArraySelectedRows = null;
var m_sUserID = "";
var m_sUnitID = "";
//Added 2 new flags to handle 'X' being clicked.
var blnConfirmProcessingDone;
var blnReprintButtonClicked;

function window.onload()
{
	scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	blnConfirmProcessingDone = false;
	blnReprintButtonClicked = false;
	InitialiseScreen();
	RetrieveData();
	GetPayments();
	ToggleReprintButton();
	<% /* scScreenFunctions.HideCollection(divStatus);  */ %>
	scScreenFunctions.ShowCollection(spnTableListScroll);
	scScreenFunctions.ShowCollection(frmScreen);
	window.returnValue = null;
}	

function InitialiseScreen()
{	
	frmScreen.btnReprint.disabled = false;
	frmScreen.btnConfirm.disabled = false;	
	//scScreenFunctions.HideCollection(frmScreen.btnReprintCheques)
}		

function RetrieveData()
{
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];

	var sParameters	= sArguments[4];
	
	m_ArraySelectedRows = sParameters[1];
	m_sUserID = sParameters[2];
	m_sUnitID = sParameters[3];
}

function GetPayments()
{
	scTableList.clear();
	
	XMLReprint = new scXMLFunctions.XMLObject();
	XMLReprint.LoadXML(m_ArraySelectedRows);		
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XMLReprint.CheckResponse(ErrorTypes);
	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		if(ErrorReturn[0] == true)
		{
			<%/* SR 06/07/01 : SYS2412 - Get FactFindData for all the applications returned from the 
							   above method  */
			%>	
			if (XMLReprint.CreateTagList("PAYMENTRECORD") != null) m_iTotalRecords = XMLReprint.ActiveTagList.length;
			else m_iTotalRecords = 0;
						
			scTableList.initialiseTable(tblReprintCheques, 0, "", Showlist, m_iTableLength, m_iTotalRecords)
			scTableList.EnableMultiSelectTable();
			Showlist(0);
		}
		else	
		{
			alert("No records have been found. Please amend your search criteria");  
		}
	}			

	ErrorTypes = null;
	ErrorReturn = null;
}

function Showlist(iOffset)
{
	var iCount ;
	var iRowCount = 1;
	var varApplicationNumber, varAmount, varPaymentSequenceNum, varChequeNum, varReprint, iReprint
		
	for (iCount=0; iCount < m_iTotalRecords && iCount < m_iTableLength;iCount++)
	{			
		if (XMLReprint.SelectTagListItem(iCount+iOffset) == true)
		{
			XMLReprint.SetAttribute("TABLEROW", iRowCount + iOffset);
			varApplicationNumber = XMLReprint.GetAttribute("APPLICATIONNUMBER");
			varPaymentSequenceNum = XMLReprint.GetAttribute("PAYMENTSEQUENCENUMBER");
			varChequeNum = XMLReprint.GetAttribute("CHEQUENUMBER");	
			//GD BMIDS01045varAmount = XMLReprint.GetAttribute("AMOUNT");
			varAmount = XMLReprint.GetAttribute("NETPAYMENTAMOUNT");
			iReprint = XMLReprint.GetAttribute("REPRINTSTATUS");
			varReprint = "No";
			//GD BMIDS00514
			if (iReprint == "1")
			{
				varReprint = "Yes";
			} 
			
			scScreenFunctions.SizeTextToField(tblReprintCheques.rows(iRowCount).cells(0), varApplicationNumber);
			scScreenFunctions.SizeTextToField(tblReprintCheques.rows(iRowCount).cells(1), varPaymentSequenceNum);
			scScreenFunctions.SizeTextToField(tblReprintCheques.rows(iRowCount).cells(2), varChequeNum);
			scScreenFunctions.SizeTextToField(tblReprintCheques.rows(iRowCount).cells(3), varAmount);
			scScreenFunctions.SizeTextToField(tblReprintCheques.rows(iRowCount).cells(4), varReprint);
					
			iRowCount++;
			
			tblReprintCheques.rows(iCount+1).setAttribute("ApplicationNumber", varApplicationNumber);
			tblReprintCheques.rows(iCount+1).setAttribute("SelectedRow");
			tblReprintCheques.rows(iCount+1).setAttribute("PaymentSequenceNumber", varPaymentSequenceNum);
			tblReprintCheques.rows(iCount+1).setAttribute("ChequeNumber", varChequeNum);
			tblReprintCheques.rows(iCount+1).setAttribute("Amount", varAmount);
			tblReprintCheques.rows(iCount+1).setAttribute("ReprintStatus", varReprint);
			tblReprintCheques.rows(iCount+1).setAttribute("LISTSEQ", iCount + iOffset);
			
			XMLReprint.SetAttribute("LISTSEQ", iCount + iOffset);
		}
	}
		
}

function frmScreen.btnReprintStatus.onclick()
{ 
	<%
	//GD BMIGS00514
	%>
	var iRowIndex = 0;
	var sReprintStatus = ""; 
	var iNewStatusCode;
	var aSelectedRows = scTableList.getArrayofRowsSelected();
	var iXMLitem;
	
	if(aSelectedRows.length == 0)
	{
		alert('Select a row.');
		return ;
	}
	
	<% /* on multi-row selection, this would be enabled only status in all the rows is same. */ %>
	//GD BMIDS00514 if(tblReprintCheques.rows(aSelectedRows[0]).cells(4).innerText == 'Yes')
	XMLReprint.SelectTagListItem(aSelectedRows[0]-1)
	if (XMLReprint.GetAttribute("REPRINTSTATUS") == "1")
	{
		sReprintStatus = 'No' ;
		iNewStatusCode = 0;
	}
	else 
	{
		sReprintStatus = 'Yes' ;
		iNewStatusCode = 1;
	}
	var iXMLitem;
	
	
	
	for(iRowIndex = 0 ; iRowIndex < aSelectedRows.length ; ++iRowIndex)
	{	
		<%
		//GD BMIDS00514 START
		//ORIGINAL scScreenFunctions.SizeTextToField(tblReprintCheques.rows(aSelectedRows[iRowIndex]).cells(4), sReprintStatus);
		//ORIGINAL XMLReprint.SelectTagListItem(tblReprintCheques.rows(aSelectedRows[iRowIndex]).getAttribute("LISTSEQ"));
		//ORIGINAL XMLReprint.SetAttribute("REPRINTSTATUS", iNewStatusCode);
		%>
		iXMLitem = aSelectedRows[iRowIndex];
		XMLReprint.SelectTagListItem(iXMLitem-1);
		XMLReprint.SetAttribute("REPRINTSTATUS", iNewStatusCode);
	}
	//Refresh Screen
	//BMIDS00559 START	
	//Showlist(0);
	scTableList.redisplaySelection();
	ToggleReprintButton()
	//BMIDS00559 END	
}
//BMIDS00681 START
function window.onbeforeunload()
{
	//Only show windows dialog, if Confirm HAS NOT BEEN DONE and we're not trying to do a reprint.
	if ((blnConfirmProcessingDone == false) && (blnReprintButtonClicked == false))
	{
		window.event.returnValue = "Leaving this screen will set all payments to 'Paid'.";
	}

}

function window.onunload()
{
	//Only do the confirm processing if Confirm HAS NOT BEEN DONE, AND we're not trying to do a reprint.
	if ((blnConfirmProcessingDone == false) && (blnReprintButtonClicked == false))
	{
		DoConfirmProcessing();
	}
}
//BMIDS00681 END

function spnPaymentSanctionList.onclick()
{
	//CheckRowSelected();
}

function frmScreen.btnConfirm.onclick()
{
var bContinue;

	bContinue = window.confirm("This will set all payments to 'Paid'. Click OK to continue. Click Cancel to stop.");

	if (bContinue)
	{
		DoConfirmProcessing();
	}
	else
	{
		return;
	}
}
//BMIDS00681 START - Place confirm processing in separate method, so window.onunload can call it if clicking 'X'
function DoConfirmProcessing()
{
	var sReturn = new Array();
	sReturn[0] = true;
	window.returnValue = sReturn;
	blnConfirmProcessingDone = true;
	window.close();
}
//BMIDS00681 END	

//BMIDS00559 START

function ToggleReprintButton()
{
	//Local function start
	function CheckForReprints()
	{
		//var blnresult = false;
		var iCount, iReprint;
		
		for (iCount=0; iCount < m_iTotalRecords;iCount++)
		{			
			if (XMLReprint.SelectTagListItem(iCount) == true)
			{
				iReprint = XMLReprint.GetAttribute("REPRINTSTATUS");
				if (iReprint == "1")
				{
					return(true);
				} 
			}
		}
		return(false)
	}	
	//Local function end
	
	
	//Main function body start
	if (CheckForReprints() == true)
	{
		//Enable Reprint button
		frmScreen.btnReprint.disabled = false;
	} else
	{
		//Disable Reprint button
		frmScreen.btnReprint.disabled = true;
	
	}
	//main fuunction body end
}

//BMIDS00559 END
function frmScreen.btnReprint.onclick()
{
	blnReprintButtonClicked = true;
	var sReturn = new Array();
	sReturn[0] = false;
	sReturn[1] = XMLReprint.XMLDocument.xml; // The XML string
	window.returnValue = sReturn;
	window.close();
}
//BMIDS00869 START
function spnPaymentSanctionList.ondblclick()
{

	if (tblReprintCheques.rows(scTableList.getRowSelected()).getAttribute("PaymentSequenceNumber") != null)
	{
		frmScreen.btnReprintStatus.onclick();

	}
}
//BMIDS00869 END
-->
</script>
</SPAN>
</body>
</html>