<%@ LANGUAGE="JSCRIPT" %>
<html>
<comment>
Workfile:      DC223.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Deposit Details (Popup)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		08/03/2000	Created
AY		31/03/00	scScreenFunctions change
AY		12/04/00	SYS0328 - Dynamic currency display
IW		03/05/00	SYS0625 - Remove Prompts, Add Decimal Places.
IW		10/05/00	SYS0722 Resolved ReadOnly Screen issues
BG		17/05/00	SYS0752 Removed Tooltips
BG		09/06/00	SYS0625 Amount fields now allow and save decimal place
LD		23/05/02	SYS4727 Use cached versions of frame functions
MC		19/04/2004	BMIDS517 White space added to the title text
GHun	28/07/2004	BMIDS823 (BBG498)PadToPrecision should never be called directly, but should 
					be called indirectly via the RoundValue function
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
</comment>
<head>
	<META HTTP-EQUIV="expires" CONTENT="Wed, 01 Jan 2003 01:01:00 GMT">
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>DC223 - Deposit Details <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>

<form id="frmScreen" year4 validate="onchange" mark>
<div style="LEFT: 10px; WIDTH: 545px; POSITION: absolute; TOP: 10px; HEIGHT: 208px" class="msgGroup">
	<TABLE STYLE="MARGIN-LEFT: 10px" class="msgLabel">
		<TR class="msgLabelHead" >
			<TH>
				Source
			</TH>
			<TH>
				Please Specify
			</TH>
			<TH>
				Amount
			</TH>
		</TR>
		<TR>
			<TD>
				<select id="cboSource1"     
      style="WIDTH: 200px" class="msgCombo" onchange="SourceChanged()" 
     ></select>
			</TD>
			<TD>
				<input id="txtSourceReason1" maxlength="40" style="WIDTH: 200px" class="msgTxt">			
			</TD>
			<TD>
				<label class="msgCurrency">£</label>
				<input id="txtAmount1" maxlength="9" style="WIDTH: 100px" class="msgTxt" onchange="CalculateTotals()">			
			</TD>
		</TR>
		<TR>
			<TD>
				<select id="cboSource2" style="WIDTH: 200px" class="msgCombo" onchange="SourceChanged()" ></select>
			</TD>
			<TD>
				<input id="txtSourceReason2" maxlength="40" style="WIDTH: 200px;" class="msgTxt">			
			</TD>
			<TD>
				<label class="msgCurrency">£</label>
				<input id="txtAmount2" maxlength="9" style="WIDTH: 100px;" class="msgTxt" onchange="CalculateTotals()">			
			</TD>
		</TR>
		<TR>
			<TD>
				<select id="cboSource3"     
      style="WIDTH: 200px" class="msgCombo" onchange="SourceChanged()" 
     ></select>
			</TD>
			<TD>
				<input id="txtSourceReason3" maxlength="40" style="WIDTH: 200px; " class="msgTxt">			
			</TD>
			<TD>
				<label  class="msgCurrency">£</label>
				<input id="txtAmount3" maxlength="9" style="WIDTH: 100px;" class="msgTxt" onchange="CalculateTotals()">			
			</TD>
		</TR>
		
		<TR>
			<TD>
				<select id="cboSource4"     
      style="WIDTH: 200px" class="msgCombo" onchange="SourceChanged()" 
     ></select>
			</TD>
			<TD>
				<input id="txtSourceReason4" maxlength="40" style="WIDTH: 200px;" class="msgTxt">			
			</TD>
			<TD>
				<label class="msgCurrency">£</label>
				<input id="txtAmount4" maxlength="9" style="WIDTH: 100px; " class="msgTxt" onchange="CalculateTotals()">			
			</TD>
		</TR>
		<TR>
			<TD>
				<select id="cboSource5"     
      style="WIDTH: 200px" class="msgCombo" onchange="SourceChanged()" 
     ></select>
			</TD>
			<TD>
				<input id="txtSourceReason5" maxlength="40" style="WIDTH: 200px; " class="msgTxt">			
			</TD>
			<TD>
				<label class="msgCurrency">£</label>
				<input id="txtAmount5" maxlength="9" style="WIDTH: 100px; " class="msgTxt" onchange="CalculateTotals()">			
			</TD>
		</TR>
		<TR>
			<TD>
				<select id="cboSource6" style="WIDTH: 200px" class="msgCombo" onchange="SourceChanged()" ></select>
			</TD>
			<TD>
				<input id="txtSourceReason6" maxlength="40" style="WIDTH: 200px;" class="msgTxt">			
			</TD>
			<TD>
				<label class="msgCurrency">£</label>
				<input id="txtAmount6" maxlength="9" style="WIDTH: 100px;" class="msgTxt" onchange="CalculateTotals()">			
			</TD>
		</TR>
		<TR class="msgLabelHead">
			
			<TD colspan=2 align=right>
				Total Deposit
			</TD>
			<TD HEIGHT=30>
				<label class="msgCurrency">£</label>
				<input id="txtTotalDeposit" maxlength="20" style="WIDTH: 100px" class="msgReadOnly" readonly tabindex=-1>
			</TD>
		</TR>
	</TABLE><!--
	<span style="LEFT: 8px; POSITION: absolute; TOP: 20px" class="msgLabel">Source</span>
	<span style="LEFT: 228px; POSITION: absolute; TOP: 20px" class="msgLabel">Amount</span>

	<span id="spnDeposit1" style="TOP: 41px; LEFT: 4px; POSITION: ABSOLUTE">
		<select id="cboSource1"  style="WIDTH: 200px; TOP: 0px; LEFT: 4px; POSITION: ABSOLUTE" class="msgCombo" onchange="SourceChanged()"></select>

		<span style="TOP: 0px; LEFT: 224px; POSITION: absolute">
			<label style="TOP: 3px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtAmount1" maxlength="9" style="POSITION: absolute; WIDTH: 100px" class="msgTxt" onchange="CalculateTotals()">
		</span> 
	</span>

	<span id="spnDeposit2" style="TOP: 65px; LEFT: 4px; POSITION: ABSOLUTE">
		<select id="cboSource2"  style="WIDTH: 200px; TOP: 0px; LEFT: 4px; POSITION: ABSOLUTE" class="msgCombo" onchange="SourceChanged()"></select>

		<span style="TOP: 0px; LEFT: 224px; POSITION: absolute">
			<label style="TOP: 3px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtAmount2"  maxlength="9" style="POSITION: absolute; WIDTH: 100px" class="msgTxt" onchange="CalculateTotals()">
		</span> 
	</span>

	<span id="spnDeposit3" style="TOP: 89px; LEFT: 4px; POSITION: ABSOLUTE">
		<select id="cboSource3" style="WIDTH: 200px; TOP: 0px; LEFT: 4px; POSITION: ABSOLUTE" class="msgCombo" onchange="SourceChanged()"></select>

		<span style="TOP: 0px; LEFT: 224px; POSITION: absolute">
			<label style="TOP: 3px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtAmount3" maxlength="9" style="POSITION: absolute; WIDTH: 100px" class="msgTxt" onchange="CalculateTotals()">
		</span> 
	</span>

	<span id="spnDeposit4" style="TOP: 113px; LEFT: 4px; POSITION: ABSOLUTE">
		<select id="cboSource4"  style="WIDTH: 200px; TOP: 0px; LEFT: 4px; POSITION: ABSOLUTE" class="msgCombo" onchange="SourceChanged()"></select>

		<span style="TOP: 0px; LEFT: 224px; POSITION: absolute">
			<label style="TOP: 3px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtAmount4"  maxlength="9" style="POSITION: absolute; WIDTH: 100px" class="msgTxt" onchange="CalculateTotals()">
		</span> 
	</span>

	<span id="spnDeposit5" style="TOP: 137px; LEFT: 4px; POSITION: ABSOLUTE">
		<select id="cboSource5"  style="WIDTH: 200px; TOP: 0px; LEFT: 4px; POSITION: ABSOLUTE" class="msgCombo" onchange="SourceChanged()"></select>

		<span style="TOP: 0px; LEFT: 224px; POSITION: absolute">
			<label style="TOP: 3px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtAmount5"  maxlength="9" style="POSITION: absolute; WIDTH: 100px" class="msgTxt" onchange="CalculateTotals()">
		</span> 
	</span>

	<span id="spnDeposit6" style="TOP: 161px; LEFT: 4px; POSITION: ABSOLUTE">
		<select id="cboSource6"  style="WIDTH: 200px; TOP: 0px; LEFT: 4px; POSITION: ABSOLUTE" class="msgCombo" onchange="SourceChanged()"></select>

		<span style="TOP: 0px; LEFT: 224px; POSITION: absolute">
			<label style="TOP: 3px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtAmount6"  maxlength="9" style="POSITION: absolute; WIDTH: 100px" class="msgTxt" onchange="CalculateTotals()">
		</span> 
	</span>
	
	<span style="LEFT: 140px; POSITION: absolute; TOP: 193px" class="msgLabel">
		Total Deposit
		<span style="TOP: 0px; LEFT: 88px; POSITION: absolute">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtTotalDeposit" maxlength="20" style="POSITION: absolute; WIDTH: 100px; TOP: -3px" class="msgReadOnly" readonly tabindex=-1>
		</span> 
	</span>-->
</div>

<span style="LEFT: 10px; POSITION: absolute; TOP: 240px">
	<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
	<span style="LEFT: 70px; POSITION: absolute; TOP: 0px">
		<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
</span>

</form><!-- #include FILE="attribs/DC223attribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var scScreenFunctions;

var DepositXML = null;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_BaseNonPopupWindow = null;

/* EVENTS */

function frmScreen.btnCancel.onclick()
{
	DepositXML = null;
	window.close();
}

function frmScreen.btnOK.onclick()
{
	if (CommitChanges())
	{
		DepositXML = null;
		window.close();
	}
}

function window.onload()
{
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters	= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	DepositXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	DepositXML.LoadXML(sParameters[0]);
	m_sReadOnly	= sParameters[1];
	m_sApplicationNumber = sParameters[2];
	m_sApplicationFactFindNumber = sParameters[3];
	scScreenFunctions.SetThisCurrency(frmScreen,sParameters[4]);

	SetMasks();
	Validation_Init();	
	Initialise();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

/* FUNCTIONS */


function CalculateTotals()
{
	var dTotal = 0;
	var dNumber = 0;

	dNumber = parseFloat(frmScreen.txtAmount1.value);
	if (!isNaN(dNumber)) dTotal = dTotal + dNumber;

	dNumber = parseFloat(frmScreen.txtAmount2.value);
	if (!isNaN(dNumber)) dTotal = dTotal + dNumber;

	dNumber = parseFloat(frmScreen.txtAmount3.value);
	if (!isNaN(dNumber)) dTotal = dTotal + dNumber;

	dNumber = parseFloat(frmScreen.txtAmount4.value);
	if (!isNaN(dNumber)) dTotal = dTotal + dNumber;

	dNumber = parseFloat(frmScreen.txtAmount5.value);
	if (!isNaN(dNumber)) dTotal = dTotal + dNumber;

	dNumber = parseFloat(frmScreen.txtAmount6.value);
	if (!isNaN(dNumber)) dTotal = dTotal + dNumber;

	<% /* BMIDS823 PadToPrecision should never be called directly, but should be called
	indirectly via the RoundValue function
	dTotal = m_BaseNonPopupWindow.top.frames[1].document.all.scMathFunctions.PadToPrecision(dTotal,2);  */ %>
	dTotal = m_BaseNonPopupWindow.top.frames[1].document.all.scMathFunctions.RoundValue(dTotal, 2);
	
	frmScreen.txtTotalDeposit.value = dTotal;
}

function CommitChanges()
{
	var bSuccess = true;

	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
				bSuccess = SaveDeposit();
		}
		else
			bSuccess = false;
	return(bSuccess);
}

function Initialise()
// Initialises the screen
{
	PopulateCombos();
	PopulateScreen();
	if(m_sReadOnly == "1")
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function PopulateCombos()
// Populates all combos on the screen
{
	var blnSuccess = true;
	ComboXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("DepositSource");

	if(ComboXML.GetComboLists(document,sGroupList))
	{
		// Populate Country combo
		blnSuccess = ComboXML.PopulateCombo(document,frmScreen.cboSource1,"DepositSource",true);
		blnSuccess = blnSuccess && ComboXML.PopulateCombo(document,frmScreen.cboSource2,"DepositSource",true);
		blnSuccess = blnSuccess && ComboXML.PopulateCombo(document,frmScreen.cboSource3,"DepositSource",true);
		blnSuccess = blnSuccess && ComboXML.PopulateCombo(document,frmScreen.cboSource4,"DepositSource",true);
		blnSuccess = blnSuccess && ComboXML.PopulateCombo(document,frmScreen.cboSource5,"DepositSource",true);
		blnSuccess = blnSuccess && ComboXML.PopulateCombo(document,frmScreen.cboSource6,"DepositSource",true);

		if(!blnSuccess) scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
}

function PopulateScreen()
// Populates the screen with details of the item selected in 
{
	DepositXML.CreateTagList("NEWPROPERTYDEPOSIT");

	for (var iCount = 0; (iCount < DepositXML.ActiveTagList.length) && (iCount < 10); iCount++)
	{
		DepositXML.SelectTagListItem(iCount);
		frmScreen.all("cboSource" + (iCount + 1)).value = DepositXML.GetTagText("SOURCEOFFUNDING");
		frmScreen.all("txtAmount" + (iCount + 1)).value = DepositXML.GetTagText("AMOUNT");
		//BMIDS746
		frmScreen.all("txtSourceReason" + (iCount + 1)).value =DepositXML.GetTagText("SOURCEREASON");
		
	}

	SourceChanged();
	CalculateTotals();
}

function SaveDeposit()
{
	var sReturn = new Array();
	if (IsChanged())
	{
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		var XMLListTag = XML.CreateActiveTag("NEWPROPERTYDEPOSITLIST");

		for (var iIndex = 1; iIndex <= 6; iIndex++)
		{
			XML.ActiveTag = XMLListTag;

			if (frmScreen.all("cboSource" + iIndex).selectedIndex > 0)
			{
				XML.CreateActiveTag("NEWPROPERTYDEPOSIT");
				XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
				XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
				XML.CreateTag("SOURCEOFFUNDING", frmScreen.all("cboSource" + iIndex).value);
				XML.CreateTag("SOURCEREASON", frmScreen.all("txtSourceReason" + iIndex).value);
				XML.CreateTag("AMOUNT", frmScreen.all("txtAmount" + iIndex).value);
			}
		}

		DepositXML = XML.XMLDocument.xml;
	}
	
	sReturn[0] = IsChanged();
	sReturn[1] = DepositXML;
	window.returnValue	= sReturn;
	window.close();
}

function SourceChanged()
{
	function SetDisabling(nIndex)
	{
		if (frmScreen.all("cboSource" + nIndex).selectedIndex == 0)
		{
			frmScreen.all("txtAmount" + nIndex).value = "";
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtAmount" + nIndex);
			frmScreen.all("txtAmount" + nIndex).disabled = true;
		}
		else
		{
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtAmount" + nIndex);
			frmScreen.all("txtAmount" + nIndex).disabled = false;
		}
	}

	SetDisabling(1);
	SetDisabling(2);
	SetDisabling(3);
	SetDisabling(4);
	SetDisabling(5);
	SetDisabling(6);

	CalculateTotals();
}
-->
</script>
</body>
</html>


