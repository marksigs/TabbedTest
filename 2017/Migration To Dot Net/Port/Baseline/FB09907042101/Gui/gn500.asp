<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      gn500.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Currency Calculator Pop Up
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
PF		6/4/01		SYS 2252 - Created
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>

<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
	<META HTTP-EQUIV="expires" CONTENT="Wed, 01 Jan 2003 01:01:00 GMT">
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>Currency Calculator  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<%/*<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="TOP: 250px; LEFT: 210px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scTable" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>*/%>

<% /* LANGUAGE=javascript  Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: visible" mark validate="onchange">
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 225px; WIDTH: 360px; POSITION: ABSOLUTE" class="msgGroup">

<span style="TOP: 15px; LEFT:10px; POSITION:ABSOLUTE" class="msgLabel">
		<label style="TOP:0px; LEFT:10px; POSITION:absolute" class="msgCurrency">Amount to convert </label> 
		<span style="TOP: 0px; LEFT:160px; POSITION:ABSOLUTE">
		<input id="txtConvertAmt" maxlength="15"  style="TOP: 0px; WIDTH:100px; POSITION:ABSOLUTE" class="msgTxt"> 
		</span>
		</span>

      <span style="TOP: 45px; LEFT: 10px; POSITION: ABSOLUTE; width: 96px; height: 19px" class="msgLabel">
      <label style="TOP:0px; LEFT:10px; POSITION: ABSOLUTE" class="msgCurrency">From Currency </label>
      <span style="TOP: 0px; LEFT: 160px; POSITION: ABSOLUTE; width: 100px; height: 19px"> 
      <select id="cboFromCurr" style="WIDTH: 170px" class="msgCombo"></select>
      </span>
      </span>
 
      <span style="TOP: 75px; LEFT: 10px; POSITION: ABSOLUTE; width: 96px; height: 19px" class="msgLabel">
      <label style="TOP:0px; LEFT:10px; POSITION:absolute" class="msgCurrency">To Currency </label>
      <span style="TOP: 0px; LEFT: 160px; POSITION: ABSOLUTE; width: 100px; height: 19px"> 
      <select id="cboToCurr" style="WIDTH: 170px" class="msgCombo"></select>
      </span>
      </span>

      <span style="TOP: 105px; LEFT:10px; POSITION:ABSOLUTE" class="msgLabel">
      <label style="TOP:0px; LEFT:10px; POSITION: ABSOLUTE" class="msgCurrency">Conversion Rate</label>
      <span style="TOP: 0px; LEFT:160px; POSITION:ABSOLUTE">
      <input id="txtConvertRate" maxlength="15"  style="TOP: 0px; WIDTH:100px; POSITION:ABSOLUTE" class="msgTxt">
      </span>
      </span>
      
      <span style="TOP: 145px; LEFT: 10px; POSITION: ABSOLUTE; width: 100px; height: 30px">
      <label style="TOP:5px; LEFT:10px; POSITION:absolute" class="msgCurrency">Result</label>
      <input id="txtResult" maxlength="15"  style="TOP: 0px; WIDTH:100px; LEFT:160px; POSITION:ABSOLUTE" class="msgTxt">
      </span>
      
      <span style="TOP: 190px; LEFT:170px; POSITION:ABSOLUTE; height: 30px">
      <input id="btnConvert" value="Convert" type="button" style="WIDTH: 60px" class="msgButton">
      </span>
      
      <span style="TOP: 190px; LEFT:280px; POSITION:ABSOLUTE; height: 30px">
      <input id="btnClose" value="Close" type="button" style="WIDTH:60px" class="msgButton" onclick="return frmScreen.btnClose.onclick()">
      </span>  
      
      </div>
</form>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/gn500Attribs.asp" -->

<script language="JScript">
<!--
var XML;
var scScreenFunctions;
var m_sUserId;
var m_sUnitId;
var m_sMachineId;
var m_sDistributionChannelId;
var m_sBaseCurrIndex = -1;
var m_sBaseCurrLocation = -1;
var m_BaseNonPopupWindow = null;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	SetMasks();
	RetrieveData();
	Validation_Init();
	frmScreen.txtConvertAmt.focus();
	PopulateFromCombo();
	frmScreen.btnConvert.disabled = true;
	frmScreen.txtConvertRate.disabled = true;
	frmScreen.txtResult.disabled = true;
	frmScreen.cboToCurr.disabled = true;

	//Set up the <SELECT> list item in the ToCurr combo
	var oOption = document.createElement("OPTION");
	oOption.text = "<SELECT>";
	frmScreen.cboToCurr.add(oOption);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function SetMasks()
{
	frmScreen.txtConvertAmt.setAttribute("filter", "[0-9.]");
	frmScreen.txtConvertAmt.setAttribute("amount", ".");
}

function PopulateFromCombo()

	// Populates the From Combos with details of the currencies returned
	{
		XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.RunASP(document,"FindCurrencyList.asp");
		XML.CreateTagList("CURRENCY");
		
		var oOption = document.createElement("OPTION");
		oOption.text = "<SELECT>";
		frmScreen.cboFromCurr.add(oOption);
	
		for (var nCurr=0; nCurr<XML.ActiveTagList.length; nCurr++)
			{
				XML.SelectTagListItem(nCurr);
				var combocurrcode = XML.GetTagText("CURRENCYCODE");
				var combocurrname = XML.GetTagText("CURRENCYNAME");
				var a_convertRate = XML.GetTagText("CONVERSIONRATE");
				var a_roundPrecision = XML.GetTagText("ROUNDINGPRECISION");
				var a_roundDirection = XML.GetTagText("ROUNDINGDIRECTION");
				var a_BaseInd = XML.GetTagText("BASECURRENCYIND");
				//Set up the combo entries
				var oOption = document.createElement("OPTION");
				oOption.value = combocurrcode;
				oOption.text = combocurrname;
				oOption.setAttribute("CONVERSIONRATE",a_convertRate)
				oOption.setAttribute("ROUNDINGPRECISION",a_roundPrecision)
				oOption.setAttribute("ROUNDINGDIRECTION",a_roundDirection)
				oOption.setAttribute("BASECURRENCYIND",a_BaseInd)
				frmScreen.cboFromCurr.add(oOption);
				if (a_BaseInd == "1")
					m_sBaseCurrIndex = (frmScreen.cboFromCurr.length - 1);
			}
		if (m_sBaseCurrIndex == -1)
			{
				alert("No Base currency found.  Please contact the Systems Supervisor.");
				frmScreen.txtConvertAmt.disabled = true;
				frmScreen.cboFromCurr.disabled = true;
				frmScreen.cboToCurr.disabled = true;
				frmScreen.btnConvert.disabled = true;
			}
	}

function PopulateToCombo()
// Populates the To Combo with details of the currencies returned
// Adjusts the lists depending on the selection made in cboFromCurr
{
	var indexcount = frmScreen.cboFromCurr.options.length;

	//Clear the Combo first
	for (var nCount=0; nCount<indexcount; nCount++)
		{
			var oOption = frmScreen.cboToCurr.options(nCount);
			frmScreen.cboToCurr.remove(oOption);
		}

	//Check that a currency has been selected in cboToCurr
	if (frmScreen.cboFromCurr.selectedIndex == "0")
			{
				//Set cboToCurr to <SELECT> and disable until next time
				var oOption = document.createElement("OPTION");
				oOption.text = "<SELECT>";
				frmScreen.cboToCurr.add(oOption);
				frmScreen.cboToCurr.disabled = true;
				frmScreen.txtConvertRate.value = "";
			}
	else	
		if (frmScreen.cboFromCurr.selectedIndex != m_sBaseCurrIndex)
			{
				//populate ToCombo with the BaseCurrency
				//No attributes assigned to combo entries cos I dont think we need them when it's BaseCurrency
				var oOption = document.createElement("OPTION");
				oOption.value = frmScreen.cboFromCurr.item(m_sBaseCurrIndex).value;
				oOption.text = frmScreen.cboFromCurr.item(m_sBaseCurrIndex).text;
				frmScreen.cboToCurr.add(oOption);
				frmScreen.cboToCurr.disabled = true;
				m_sBaseCurrLocation = 2;
			}
		else 
			{	//alert("cboFromCurr is Base")
				//Set up the <SELECT> list item
				var oOption = document.createElement("OPTION");
				oOption.text = "<SELECT>";
				frmScreen.cboToCurr.add(oOption);
				
				//Create a loop which sets up all the other currencies EXCEPT Base.
				var oLength = frmScreen.cboFromCurr.options.length;
					for (var nCurr=1; nCurr<oLength; nCurr++)
						{
							if (nCurr != m_sBaseCurrIndex)
							{
							var a_convertRate = frmScreen.cboFromCurr.item(nCurr).getAttribute("CONVERSIONRATE");
							var a_roundPrecision = frmScreen.cboFromCurr.item(nCurr).getAttribute("ROUNDINGPRECISION");
							var a_roundDirection = frmScreen.cboFromCurr.item(nCurr).getAttribute("ROUNDINGDIRECTION");
							var a_BaseInd = frmScreen.cboFromCurr.item(nCurr).getAttribute("BASECURRENCYIND");

							//Set up the combo entries
							var oOption = document.createElement("OPTION");
							oOption.value = frmScreen.cboFromCurr.item(nCurr).value;
							oOption.text = frmScreen.cboFromCurr.item(nCurr).text;
							oOption.setAttribute("CONVERSIONRATE",a_convertRate)
							oOption.setAttribute("ROUNDINGPRECISION",a_roundPrecision)
							oOption.setAttribute("ROUNDINGDIRECTION",a_roundDirection)
							oOption.setAttribute("BASECURRENCYIND",a_BaseInd)
							frmScreen.cboToCurr.add(oOption);
							}
						}
			frmScreen.cboToCurr.disabled = false;
			m_sBaseCurrLocation = 1;
			PopulateConvertRate();
			}
}

function RetrieveData()
{		
 	var sArguments			= window.dialogArguments;
	//window size and position
	window.dialogTop		= sArguments[0]; 
	window.dialogLeft		= sArguments[1];
	window.dialogWidth		= sArguments[2];
	window.dialogHeight		= sArguments[3];
	var RecordArguments		= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sAttributeArray=RecordArguments;			
	m_sUserId = RecordArguments[0]; // UserId
	m_sUnitId = RecordArguments[1]; // UnitId
	m_sMachineId = RecordArguments[2]; // User Name
	m_sChannelId = RecordArguments[3]; // Unit Name	
}

function frmScreen.btnClose.onclick()
{
	window.close();
}

function frmScreen.btnConvert.onclick()
{
	var IndexFromSelected = frmScreen.cboFromCurr.selectedIndex;
	var IndexToSelected = frmScreen.cboToCurr.selectedIndex;
	var ConvertAmt = frmScreen.txtConvertAmt.value;
	var ConvertRate = frmScreen.txtConvertRate.value;

	if (m_sBaseCurrLocation != 1)
		{
			var ConvertPrec = frmScreen.cboFromCurr.item(IndexFromSelected).getAttribute("ROUNDINGPRECISION");
			var ConvertDirection = frmScreen.cboFromCurr.item(IndexFromSelected).getAttribute("ROUNDINGDIRECTION");
			var ConvertResult = ConvertAmt / ConvertRate;
			if (ConvertDirection == "30")
				{
					frmScreen.txtResult.value = m_BaseNonPopupWindow.top.frames[1].document.all.scMathFunctions.RoundValue(ConvertResult, ConvertPrec);
				}
			else frmScreen.txtResult.value = m_BaseNonPopupWindow.top.frames[1].document.all.scMathFunctions.RoundWithDirection(ConvertResult, ConvertPrec, ConvertDirection);
		}
	else
		{	
			var ConvertPrec = frmScreen.cboToCurr.item(IndexToSelected).getAttribute("ROUNDINGPRECISION");
			var ConvertDirection = frmScreen.cboToCurr.item(IndexToSelected).getAttribute("ROUNDINGDIRECTION");
			var ConvertResult = ConvertAmt * ConvertRate;
			if (ConvertDirection == "30")
				{
					frmScreen.txtResult.value = m_BaseNonPopupWindow.top.frames[1].document.all.scMathFunctions.RoundValue(ConvertResult, ConvertPrec);
				}
			else frmScreen.txtResult.value = m_BaseNonPopupWindow.top.frames[1].document.all.scMathFunctions.RoundWithDirection(ConvertResult, ConvertPrec, ConvertDirection);
		}
}

function ValidateCurrency(oComboName, CurrencySelection)
	{
		var validCurrency = 1;
		//Checks that all required values are present for the selected currency
		if (oComboName.selectedIndex != "0")
			{
				if (oComboName.item(CurrencySelection).getAttribute("BASECURRENCYIND") != "1")
					{
						if (oComboName.item(CurrencySelection).getAttribute("CONVERSIONRATE") == "")		
							{
								validCurrency = 0;
							}
						else if (oComboName.item(CurrencySelection).getAttribute("ROUNDINGPRECISION") == "")		
									{
										validCurrency = 0;
									}
						else if (oComboName.item(CurrencySelection).getAttribute("ROUNDINGDIRECTION") == "")		
									{
										validCurrency = 0;
									}
						else
									{
										validCurrency = 1;
									}
					}
			}
			
		if (validCurrency == 0)
			{
				alert ("This Currency has not been properly maintained. Please contact the Systems Supervisor");
				frmScreen.txtResult.value = "";
				frmScreen.btnConvert.disabled = true;
			}	
		else
			{
				frmScreen.btnConvert.disabled = false;
			}
	}


function frmScreen.cboFromCurr.onchange() 
	{
		EnableConvert();
		frmScreen.txtResult.value = "";
		var cboIndexSelected = frmScreen.cboFromCurr.selectedIndex;
		ValidateCurrency(frmScreen.cboFromCurr, cboIndexSelected);
		PopulateToCombo();
		PopulateConvertRate();
	}

function frmScreen.cboToCurr.onchange()
	{
		EnableConvert();
		frmScreen.txtResult.value = "";
		var cboIndexSelected = frmScreen.cboToCurr.selectedIndex;
		ValidateCurrency(frmScreen.cboToCurr, cboIndexSelected);
		PopulateConvertRate();
	}

function PopulateConvertRate()
	{
		if (m_sBaseCurrLocation == "1")
			{
				var cboIndexSelected = frmScreen.cboToCurr.selectedIndex;
				frmScreen.txtConvertRate.value = frmScreen.cboToCurr.item(cboIndexSelected).getAttribute("CONVERSIONRATE");
			}
		else
			{
				var cboIndexSelected = frmScreen.cboFromCurr.selectedIndex;
				frmScreen.txtConvertRate.value = frmScreen.cboFromCurr.item(cboIndexSelected).getAttribute("CONVERSIONRATE");
			}
	
		if (frmScreen.txtConvertRate.value == "null")
			{
				frmScreen.txtConvertRate.value = "";
			}
	}

function EnableConvert()
	{
		if (frmScreen.cboFromCurr.selectedIndex != "0")
			{
				if (frmScreen.cboToCurr.selectedIndex != "0")
					{
							frmScreen.btnConvert.disabled = false;
					}
			}
		else if (frmScreen.cboToCurr.selectedIndex != "0")
			{
				if	(frmScreen.cboFromCurr.selectedIndex != "0")
					{
							frmScreen.btnConvert.disabled = false;
					}
			}
		else frmScreen.btnConvert.disabled = true;	
	}



-->

</script>

</body>
</html>



