<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<% /*
Workfile:		cm120.asp
Copyright:		Copyright © 2000 Marlborough Stirling

Description:	Further Filtering
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		29/03/00	scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
AP		15/08/00	SYS1056 Set 'All Products Without Checks' option = Disabled
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:
Prog	Date		Description

AW		11/06/2002	BM011	Special groups combo
GD		19/06/2002	BMIDS00077 - Upgrade to Core 7.0.2
MO		01/08/2002	BMIDS00281 - Fixed bug, to capture the IE default onsubmit event when ENTER is pressed.
MV		22/08/2002	BMIDS00355	IE 5.5 upgrade Errors - Modified the HTML layout
ASu		11/09/2002	BMIDS00413 - Resize and position items within pop up in order to correct layout. 
MO		07/11/2002	BMIDS00717 - Change default to all products without checks
MDC		15/11/2002	BMIDS00939  CC015 Ported Products
DB		28/03/2003	BM0009 - SpecificMortgageProduct loses its checked status when re-entering screen.
HMA     14/10/2003  BMIDS650   - Correct 'all products with checks' filtering
MC		20/04/2004	BMIDS517	White space padded to the title text
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Mars History:
Prog	Date		Description

GHun	15/06/2006	MAR1849 Removed extra space from ComboGroup name
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specifc History:
AShaw	07/11/2006	EP2_8	New Flexible options.
							Pass in params to disable options.
							Add code to disable Special Groups.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<html>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"/>
	<title>Further Filtering  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body><!-- Form Manager Object - Controls Soft Coded Field Attributes -->
<!-- Screen Functions Object - see comments within scScreenFunctions.htm for details -->
<%/*
GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2
<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT><!-- XML Functions Object - see comments within scXMLFunctions.htm for details -->
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>*/%>
<!-- Specify Forms Here -->
<!-- Specify Screen Layout Here - the visibility is set to hidden making the screen display more graceful-->
<form id="frmScreen" style ="VISIBILITY: visible" validate="onchange" mark >
<div id="divBackground" style="HEIGHT: 380px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 461px" class="msgGroup">
	<span style="HEIGHT: 80px; LEFT: 20px; POSITION: absolute; TOP: 20px; WIDTH: 420px" class="msgGroupLight">
		<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
			<strong>Filtering search using...</strong>
		</span>
		<% /* MO - 07/11/2002 - BMIDS00717 - Changed default */ %>
		<span style="LEFT: 10px; WIDTH: 180px; POSITION: absolute; TOP: 30px" class="msgLabel">
			<label for="optAllProductsWithChecks" class="msgLabel">All Products With Checks?</label>
			<input id="optAllProductsWithChecks" name="Filtering" type="radio" value="0" checked
					 style="LEFT: 160px; POSITION: absolute; TOP: -3px">
		</span>
		<span style="LEFT: 10px; WIDTH: 180px; POSITION: absolute; TOP: 50px" class="msgLabel">
			<label for="optAllProductsWithoutChecks" class="msgLabel">All Products Without Checks?</label>
			<input id="optAllProductsWithoutChecks" name="Filtering" type="radio" value="1" style="LEFT: 160px; POSITION: absolute; TOP: -3px">
		</span>
	</span>
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	<span style="TOP: 110px; LEFT: 20px; HEIGHT: 250px; WIDTH: 420px; POSITION: ABSOLUTE" class="msgGroupLight">
		<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			<strong>Further Filtering options...</strong>
		</span>
		<span id="spnFilterOptionCheckboxs"	style="TOP: 30px; LEFT: 10px; WIDTH: 155px; POSITION: absolute">
				<label for="chkShowAllProductsByGroup" class="msgLabel">All Products By Group?</label>
				<input id="chkShowAllProductsByGroup" name="ShowAllProductsByGroup" 
					type="checkbox" value="1" style="LEFT: 155px; POSITION: absolute">
			
			<span style="TOP: 30px; LEFT: 0px; WIDTH: 155px; POSITION: absolute">
				<label for="chkShowProductsByCode" class="msgLabel">Specific Mortgage Product?</label>
				<input id="chkShowProductsByCode" name="chkShowProductsByCode" 
					type="checkbox" value="1" style="LEFT: 155px; POSITION:absolute">
			</span>
			<span style="TOP: 60px; LEFT: 0px; WIDTH: 155px; POSITION: ABSOLUTE">
				<label for="chkShowDiscountedProducts" class="msgLabel">Discounted Products?</label>
				<input id="chkShowDiscountedProducts" name="ShowDiscountedProducts" 
					type="checkbox" value="1" style="LEFT: 155px; POSITION: ABSOLUTE">
			</span>
			<span style="TOP: 90px; LEFT: 0px; WIDTH: 155px; POSITION: ABSOLUTE">
				<label for="chkShowFixedRateProducts"  class="msgLabel">Fixed Rate Products?</label>
				<input id="chkShowFixedRateProducts" name="ShowFixedRateProducts" 
					type="checkbox" value="1" style="LEFT: 155px; POSITION: ABSOLUTE">
			</span>
			<span style="TOP: 120px; LEFT: 0px; WIDTH: 155px; POSITION: ABSOLUTE">
				<label for="chkShowStandardVarProducts" class="msgLabel">Standard Variable Products?</label>
				<input id="chkShowStandardVarProducts" name="ShowStandardVarProducts" 
					type="checkbox" value="1" style="LEFT: 155px; POSITION: ABSOLUTE">
			</span>
			<span style="TOP: 150px; LEFT: 0px; WIDTH: 155px; POSITION: ABSOLUTE">
				<label for="chkShowCappedFloorProducts" class="msgLabel">Capped/Floored Products?</label>
				<input id="chkShowCappedFloorProducts" name="ShowCappedFloorProducts" 
					type="checkbox" value="1" style="LEFT: 155px; POSITION: ABSOLUTE">
			</span>
			<!-- BMIDS00939 MDC 15/11/2002 -->
			<span style="TOP: 180px; LEFT: 0px; WIDTH: 155px; POSITION: ABSOLUTE">
				<label for="chkShowPortedProducts" class="msgLabel">Portable Products?</label>
				<input id="chkShowPortedProducts" name="ShowPortedProducts" 
					type="checkbox" value="1" style="LEFT: 155px; POSITION: ABSOLUTE">
			</span>
			<!-- AShaw 25/10/06 EP2_8 -->
			<span style="TOP: 210px; LEFT: 0px; WIDTH: 155px; POSITION: ABSOLUTE">
				<label for="chkFlexibleFilter" class="msgLabel">Felxible/Non-Flexible Products?</label>
				<input id="chkFlexibleFilter" name="FlexibleFilter" 
					type="checkbox" value="1" style="LEFT: 155px; POSITION: ABSOLUTE">
			</span>
			<!-- BMIDS00939 MDC 15/11/2002 - End -->
		</span>
		<span id="spnFilterOptionValues" style="TOP: 35px; LEFT: 210px; POSITION: ABSOLUTE" class="msgLabel">
			<span style="LEFT: 0px; POSITION: absolute; TOP: 0px" class="msgLabel">
				Special Group
				<span style="LEFT: 100px; POSITION: absolute; TOP: 0px">
					<select id="cboSpecialGroup" name="SpecialGroup" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgCombo"></select>
				</span>
			</span>
			<span style="TOP: 30px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
<!-- SG 29/05/02 SYS4767 START-->
				Product Code
				<span style="TOP: 0px; LEFT: 100px; POSITION: ABSOLUTE">
					<input id="txtProductCode" name="ProductCode"  
							maxlength="6" style="TOP: -3px; WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
				</span>
			</span>
			<span style="TOP: 60px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
<!-- SG 29/05/02 SYS4767 END-->
				Discounted Period
				<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
					<select id="cboDiscountedPeriod" name="DiscountedPeriod" style="WIDTH: 100px" class="msgCombo"></select>
				</span>
			</span>
<!-- SG 29/05/02 SYS4767 START-->
			<span style="TOP: 90px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
<!-- SG 29/05/02 SYS4767 END-->
				Fixed Rate Period
				<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
					<select id="cboFixedRatePeriod" name="FixedRatePeriod" style="WIDTH: 100px" class="msgCombo"></select>
				</span>
			</span>
<!-- SG 29/05/02 SYS4767 START-->
			<span style="TOP: 150px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
<!-- SG 29/05/02 SYS4767 END-->
				Capped/Floored <br>Period
				<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
					<select id="cboCappedFlooredPeriod" name="CappedFlooredPeriod" style="WIDTH: 100px" class="msgCombo"></select>
				</span>
			</span>
<!-- AShaw 25/10/06 EP2_8 -->
			<span style="LEFT: 0px; WIDTH: 60px; POSITION: absolute; TOP: 210px" class="msgLabel">
				<label for="optFlexible" class="msgLabel">Flexible</label>
				<input id="optFlexible" name="Flexible" type="radio" value="1" checked
					 style="LEFT: 60px; POSITION: absolute; TOP: -3px" disabled>
			</span>
			<span style="LEFT: 100px; WIDTH: 80px; POSITION: absolute; TOP: 210px" class="msgLabel">
				<label for="optNonFlexible" class="msgLabel">Non-Flexible</label>
				<input id="optNonFlexible" name="Flexible" type="radio" value="0" style="LEFT: 80px; POSITION: absolute; TOP: -3px" disabled>
			</span>
		</span>
	</span>
<!-- SG 29/05/02 SYS4767 START-->
	<span style="TOP: 375px; LEFT: 10px; POSITION: ABSOLUTE">
<!-- SG 29/05/02 SYS4767 END-->
		<span style="TOP: 10px; LEFT: 0px; POSITION: ABSOLUTE">
			<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="TOP: 10px; LEFT: 65px; POSITION: ABSOLUTE">
			<input id="btnCancel" value="Cancel"type="button" style="WIDTH: 60px" class="msgButton">
		</span>
	</span>




</div>
</form>
<!-- File containing field attributes (remove if not required) -->
<!--  #include FILE="attribs/cm120attribs.asp" -->
<!-- Specify Code Here -->
<script language="JScript" type="text/javascript">
<!--
var m_sApplicationMode		= null;
var	m_sUserId			= null;
var	m_sUnitId			= null;
var	m_sUserType			= null;
var m_sFurtherFiltering	= null;
var m_bOkButtonPressed  = false; <%/*MO		1/08/2002	BMIDS00281*/%>
var scScreenFunctions;
<%/*GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2*/%>
var m_BaseNonPopupWindow = null;
var m_CalculatedFilter = null;  // EP2_8 - New Filter.

function window.onload()
{
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();%>
	RetrieveData();
	SetMasks();
	GetComboLists();
	PopulateScreen();
	Validation_Init();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	window.returnValue = null;
}
function RetrieveData()
{
	var sArguments = window.dialogArguments;
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters		= sArguments[4];
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	m_BaseNonPopupWindow = sArguments[5];
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sApplicationMode		= sParameters[0];
	m_sUserId			= sParameters[1];
	m_sUnitId			= sParameters[2];
	m_sUserType			= sParameters[3];
	m_sFurtherFiltering	= sParameters[4];	// XML String
}
function PopulateScreen()
{
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	<%//var XMLFurtherFiltering = new scXMLFunctions.XMLObject();%>
	var XMLFurtherFiltering = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	// APS 01/12/99 So that we get frames around our combo controls
	scScreenFunctions.ShowCollection(frmScreen);
	XMLFurtherFiltering.LoadXML(m_sFurtherFiltering);
	if (XMLFurtherFiltering.SelectTag(null, "FURTHERFILTERING") != null)
	{
		// EP2_8 - Calculated filter value
		m_CalculatedFilter = XMLFurtherFiltering.GetTagText("CALCULATEDFILTER");
		if (m_CalculatedFilter != "-1") 
		{
			scScreenFunctions.SetCheckBoxValue(frmScreen, frmScreen.chkFlexibleFilter.id, true);
			scScreenFunctions.SetRadioGroupState(frmScreen, "Flexible" , "R");
			scScreenFunctions.SetRadioGroupValue(frmScreen, "Flexible", m_CalculatedFilter);
			scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.chkFlexibleFilter.id);
		}
		
		//	AW	06/06/02	BM012
		var sFilterSearch = XMLFurtherFiltering.GetTagText("ALLPRODUCTSWITHCHECKS");
		
		<% /* BMIDS650  The ALLPRODUCTSWITHCHECKS tag holds a 'With checks' flag (1=With   0=Without)
		                With Checks radio value = 0
		                Without Checks radio value = 1 */ %>
		                
		if (sFilterSearch == "1")  //With Checks
		{
			scScreenFunctions.SetRadioGroupValue(frmScreen, "Filtering", "0");
		}
		else                       //sFilterSearch = 0 or null(default to Without)
		{
			scScreenFunctions.SetRadioGroupValue(frmScreen, "Filtering", "1");
		}

		//if (frmScreen.optFurtherFiltering.checked == true)
		//{
			var sProductsByGroup = XMLFurtherFiltering.GetTagText("PRODUCTSBYGROUP");
			scScreenFunctions.SetCheckBoxValue(frmScreen, frmScreen.chkShowAllProductsByGroup.id, sProductsByGroup);
			if (sProductsByGroup == frmScreen.chkShowAllProductsByGroup.value)
			{
				frmScreen.chkShowAllProductsByGroup.checked = true;
				frmScreen.cboSpecialGroup.value = XMLFurtherFiltering.GetTagText("PRODUCTGROUP");
			}
			else scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.cboSpecialGroup.id);
			<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
			//SG 29/05/02 SYS4767 START			
			var sProductOverrideCode = XMLFurtherFiltering.GetTagText("PRODUCTSBYOVERRIDECODE");
			scScreenFunctions.SetCheckBoxValue(frmScreen, frmScreen.chkShowProductsByCode.id, sProductOverrideCode);
			if (sProductOverrideCode == frmScreen.chkShowProductsByCode.value)
			{
				frmScreen.chkShowProductsByCode.checked = true;
				frmScreen.txtProductCode.value = XMLFurtherFiltering.GetTagText("PRODUCTOVERRIDECODE");
				<% /* DB BM0009 - When user re-enters screen this field should still be checked */ %>
				frmScreen.chkShowProductsByCode.onclick();
				<% /* DB End */ %>
			}
			else scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.txtProductCode.id);
			//SG 29/05/02 SYS4767 END
			var sDiscountedProducts = XMLFurtherFiltering.GetTagText("DISCOUNTEDPRODUCTS");
			scScreenFunctions.SetCheckBoxValue(frmScreen, frmScreen.chkShowDiscountedProducts.id, sDiscountedProducts);
			if (sDiscountedProducts == frmScreen.chkShowDiscountedProducts.value)
			{
				frmScreen.chkShowDiscountedProducts.checked = true;
				scScreenFunctions.SetComboOnValidationType(frmScreen, "DiscountedPeriod", XMLFurtherFiltering.GetTagText("DISCOUNTEDPERIOD"));
			}
			else scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.cboDiscountedPeriod.id);

			var sFixedRateProducts = XMLFurtherFiltering.GetTagText("FIXEDRATEPRODUCTS");
			scScreenFunctions.SetCheckBoxValue(frmScreen, frmScreen.chkShowFixedRateProducts.id, sFixedRateProducts);
			
			if (sFixedRateProducts == frmScreen.chkShowFixedRateProducts.value)
			{
				frmScreen.chkShowFixedRateProducts.checked = true;
				scScreenFunctions.SetComboOnValidationType(frmScreen, "FixedRatePeriod", XMLFurtherFiltering.GetTagText("FIXEDRATEPERIOD"));
			}
			else scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.cboFixedRatePeriod.id);

			var sStandardVariableProducts = XMLFurtherFiltering.GetTagText("STANDARDVARIABLEPRODUCTS");
			scScreenFunctions.SetCheckBoxValue(frmScreen, frmScreen.chkShowStandardVarProducts.id, sStandardVariableProducts);
			
			var sCappedFlooredProducts = XMLFurtherFiltering.GetTagText("CAPPEDFLOOREDPRODUCTS");
			scScreenFunctions.SetCheckBoxValue(frmScreen, frmScreen.chkShowCappedFloorProducts.id, sCappedFlooredProducts);
			
			if (sCappedFlooredProducts == frmScreen.chkShowCappedFloorProducts.value)
			{
				frmScreen.chkShowCappedFloorProducts.checked = true;
				scScreenFunctions.SetComboOnValidationType(frmScreen, "CappedFlooredPeriod", XMLFurtherFiltering.GetTagText("CAPPEDFLOOREDPERIOD"));
			}
			else scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.cboCappedFlooredPeriod.id);

			<% /* BMIDS00939 MDC 15/11/2002 */ %>
			var sPortedProducts = XMLFurtherFiltering.GetTagText("MANUALPORTEDLOANIND");
			scScreenFunctions.SetCheckBoxValue(frmScreen, frmScreen.chkShowPortedProducts.id, sPortedProducts);
			<% /* BMIDS00939 MDC 15/11/2002 - End */ %>
		//}
	}
	//	AW	06/06/02	BM012
	/*if (frmScreen.optFurtherFiltering.checked != true)
	{
		scScreenFunctions.DisableCollection(spnFilterOptionCheckboxs);
		scScreenFunctions.DisableCollection(spnFilterOptionValues);
	}*/
	
			<%/* AShaw	25/10/2006	EP2_8 - New logic */%>
			if (m_CalculatedFilter == "-1")
			{
				var sFlexibleProducts = XMLFurtherFiltering.GetTagText("FLEXIBLENONFLEXIBLEPRODUCTS");
				var sFlexible = XMLFurtherFiltering.GetTagText("FLEXIBLEIND");
				scScreenFunctions.SetCheckBoxValue(frmScreen, frmScreen.chkFlexibleFilter.id, sFlexibleProducts);
				if (sFlexible == "0")
				{
					frmScreen.optNonFlexible.checked = true;
					frmScreen.chkFlexibleFilter.onclick();
				}
				else
				{
					frmScreen.optFlexible.checked = true;
					frmScreen.chkFlexibleFilter.onclick();
				}
			}
			
	// Ep2_8 - Disable Special groups for Epsom.
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bEnableSpecialGroup = XML.GetGlobalParameterBoolean(document,"CMFilterBySpecialGroup");			
	if (bEnableSpecialGroup == 0)
	{
		scScreenFunctions.SetCheckBoxValue(frmScreen, frmScreen.chkShowAllProductsByGroup.id, false);
		scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.chkShowAllProductsByGroup.id);
	}

	
	XMLFurtherFiltering = null;
}

//	AW	06/06/02	BM012
function frmScreen.optAllProductsWithChecks.onclick()
{
	//scScreenFunctions.DisableCollection(spnFilterOptionCheckboxs);
	//scScreenFunctions.DisableCollection(spnFilterOptionValues);
}
function frmScreen.optAllProductsWithoutChecks.onclick()
{
	//scScreenFunctions.DisableCollection(spnFilterOptionCheckboxs);
	//scScreenFunctions.DisableCollection(spnFilterOptionValues);
}

function frmScreen.btnCancel.onclick()
{
	window.close();
}
function CheckMandatoryFieldsOK()
{
	var bSuccess = true;
	
	if(frmScreen.chkShowAllProductsByGroup.checked == true &&
	   frmScreen.cboSpecialGroup.value == "")
	{
		bSuccess = false;
		alert("Please enter Special Group");
		frmScreen.cboSpecialGroup.focus();
	}
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	//SG 29/05/02 SYS4767 START
	if(frmScreen.chkShowProductsByCode.checked == true &&
	   frmScreen.txtProductCode.value == "")
	{
		bSuccess = false;
		alert("Please enter Product Code");
		frmScreen.txtProductCode.focus();
	}
	//SG 29/05/02 SYS4767 END
		
	if(frmScreen.chkShowCappedFloorProducts.checked == true &&
	   frmScreen.cboCappedFlooredPeriod.value == "")
	{
		bSuccess = false;
		alert("Please enter Capped/Floored period choice");
		frmScreen.cboCappedFlooredPeriod.focus();
	}
	if(frmScreen.chkShowDiscountedProducts.checked == true &&
	   frmScreen.cboDiscountedPeriod.value == "")
	{
		bSuccess = false;
		alert("Please enter Discounted period choice");
		frmScreen.cboDiscountedPeriod.focus();
	}
	if(frmScreen.chkShowFixedRateProducts.checked == true &&
	   frmScreen.cboFixedRatePeriod.value == "")
	{
		bSuccess = false;
		alert("Please enter Fixed Rate period choice");
		frmScreen.cboFixedRatePeriod.focus();
	}
	return(bSuccess);
}
function frmScreen.btnOK.onclick()
{
	
	<%/*MO		1/08/2002	BMIDS00281*/%>
	m_bOkButtonPressed = true;

	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	<%//var XMLFurtherFiltering = new scXMLFunctions.XMLObject();%>
	var XMLFurtherFiltering = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLFurtherFiltering.CreateActiveTag("FURTHERFILTERING");
	if ((m_sApplicationMode == "Mortgage Calculator") && (frmScreen.optAllProductsWithoutChecks.checked == true))
	{
		// TODO: Error handler code here...
		alert("raise error 205");
	}
	else if(CheckMandatoryFieldsOK())
	{
		if (frmScreen.onsubmit() == true)
		{
			var sFilterSearch = scScreenFunctions.GetRadioGroupValue(frmScreen, "Filtering");
			
			<% /* BMIDS650  Start
				  WITHCHECKS = Radio value 0.   WITHOUTCHECKS = Radio Value 1. 
			      Change the FilterSearch to a WithChecks flag */ %>
			if (sFilterSearch == "0")
			{
				sFilterSearch = "1";
			}
			else
			{
				sFilterSearch = "0";
			}	
			<% /* BMIDS650  End */ %>

			XMLFurtherFiltering.CreateTag("ALLPRODUCTSWITHCHECKS", sFilterSearch);
			//	AW	06/06/02	BM012
			//if (frmScreen.optFurtherFiltering.checked == true)
			//{			var xxxx = frmScreen.cboSpecialGroup.nodeValue
			XMLFurtherFiltering.CreateTag("PRODUCTSBYGROUP", scScreenFunctions.GetCheckBoxValue(frmScreen, frmScreen.chkShowAllProductsByGroup.id));
			XMLFurtherFiltering.CreateTag("PRODUCTGROUP", frmScreen.cboSpecialGroup.value);
			XMLFurtherFiltering.CreateTag("DISCOUNTEDPRODUCTS", scScreenFunctions.GetCheckBoxValue(frmScreen, frmScreen.chkShowDiscountedProducts.id));
			XMLFurtherFiltering.CreateTag("DISCOUNTEDPERIOD", scScreenFunctions.GetComboValidationType(frmScreen, "DiscountedPeriod"));
			XMLFurtherFiltering.CreateTag("FIXEDRATEPRODUCTS", scScreenFunctions.GetCheckBoxValue(frmScreen, frmScreen.chkShowFixedRateProducts.id));
			XMLFurtherFiltering.CreateTag("FIXEDRATEPERIOD", scScreenFunctions.GetComboValidationType(frmScreen,"FixedRatePeriod"));
			XMLFurtherFiltering.CreateTag("STANDARDVARIABLEPRODUCTS", scScreenFunctions.GetCheckBoxValue(frmScreen, frmScreen.chkShowStandardVarProducts.id));
			XMLFurtherFiltering.CreateTag("CAPPEDFLOOREDPRODUCTS", scScreenFunctions.GetCheckBoxValue(frmScreen, frmScreen.chkShowCappedFloorProducts.id));
			XMLFurtherFiltering.CreateTag("CAPPEDFLOOREDPERIOD", scScreenFunctions.GetComboValidationType(frmScreen, "CappedFlooredPeriod"));
			<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
			//SG 29/05/02 SYS4767 START				
			XMLFurtherFiltering.CreateTag("PRODUCTSBYOVERRIDECODE", scScreenFunctions.GetCheckBoxValue(frmScreen, frmScreen.chkShowProductsByCode.id));
			XMLFurtherFiltering.CreateTag("PRODUCTOVERRIDECODE", frmScreen.txtProductCode.value);
			//SG 29/05/02 SYS4767 END
			//}
			<% /* BMIDS00939 MDC 15/11/2002 */ %>
			XMLFurtherFiltering.CreateTag("MANUALPORTEDLOANIND", scScreenFunctions.GetCheckBoxValue(frmScreen, frmScreen.chkShowPortedProducts.id));
			<% /* BMIDS00939 MDC 15/11/2002 - End */ %>

			<%/* AShaw	25/10/2006	EP2_8 - New logic */%>
			XMLFurtherFiltering.CreateTag("CALCULATEDFILTER", m_CalculatedFilter);
			if (frmScreen.chkFlexibleFilter.checked == true)
			{
				XMLFurtherFiltering.CreateTag("FLEXIBLENONFLEXIBLEPRODUCTS", "1");
				if (frmScreen.optFlexible.checked == true)
					XMLFurtherFiltering.CreateTag("FLEXIBLEIND", "1");
				else
				{
					XMLFurtherFiltering.CreateTag("FLEXIBLEIND", "0");
				}
			}
			else
			{
				XMLFurtherFiltering.CreateTag("FLEXIBLENONFLEXIBLEPRODUCTS", "0");
				XMLFurtherFiltering.CreateTag("FLEXIBLEIND", "1");
			}
			

			var sReturn = new Array();	
			sReturn[0] = IsChanged();
			sReturn[1] = XMLFurtherFiltering.XMLDocument.xml;
			window.returnValue	= sReturn;
			XMLFurtherFiltering = null;
			window.close();
		}
	}
	
	<%/*MO		1/08/2002	BMIDS00281*/%>
	m_bOkButtonPressed = false;
	
}
function GetComboLists()
{
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%//var XML = new scXMLFunctions.XMLObject();%>
	var sGroups = new Array("SpecialGroup", "FilteringDiscountPeriod", "FilteringFixedRatePeriod", "FilteringCappedFlooredPeriod");
	var bSuccess = false;
	if (XML.GetComboLists(document, sGroups) == true)
	{
		bSuccess = true;
		bSuccess = bSuccess & XML.PopulateCombo(document, frmScreen.cboSpecialGroup, "SpecialGroup", true);
		bSuccess = bSuccess & XML.PopulateCombo(document, frmScreen.cboDiscountedPeriod, "FilteringDiscountPeriod", true);
		bSuccess = bSuccess & XML.PopulateCombo(document, frmScreen.cboFixedRatePeriod, "FilteringFixedRatePeriod", true);
		bSuccess = bSuccess & XML.PopulateCombo(document, frmScreen.cboCappedFlooredPeriod, "FilteringCappedFlooredPeriod", true);
	}
	if(bSuccess == false)
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
}
<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
//SG 29/05/02 SYS4767 START
/* STB: SYS4219 - Filtering must be done either soley on product code or on  */
/* any of the other criteria available.                                      */
function frmScreen.chkShowProductsByCode.onclick()
{	
	if (frmScreen.chkShowProductsByCode.checked == true)
	{
		// Enable the input for product code.
		scScreenFunctions.SetFieldToWritable(frmScreen, frmScreen.txtProductCode.id);
				
		// Disable all other filters.
		scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.chkShowAllProductsByGroup.id);
		scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.chkShowDiscountedProducts.id);
		scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.chkShowFixedRateProducts.id);
		scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.chkShowStandardVarProducts.id);
		scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.chkShowCappedFloorProducts.id);
		<% /* BMIDS00939 MDC 15/11/2002 */ %>
		scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.chkShowPortedProducts.id);
		<% /* BMIDS00939 MDC 15/11/2002 - End */ %>
		
		// Clear the values of all other filters.
		scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.cboSpecialGroup.id);
		scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.cboDiscountedPeriod.id);
		scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.cboFixedRatePeriod.id);
		scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.cboCappedFlooredPeriod.id);
	}	
	else
	{
		// Disable input for the product code.
		scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.txtProductCode.id);
		
		// Enable all other filters.
		scScreenFunctions.SetFieldToWritable(frmScreen, frmScreen.chkShowAllProductsByGroup.id);
		scScreenFunctions.SetFieldToWritable(frmScreen, frmScreen.chkShowDiscountedProducts.id);
		scScreenFunctions.SetFieldToWritable(frmScreen, frmScreen.chkShowFixedRateProducts.id);
		scScreenFunctions.SetFieldToWritable(frmScreen, frmScreen.chkShowStandardVarProducts.id);
		scScreenFunctions.SetFieldToWritable(frmScreen, frmScreen.chkShowCappedFloorProducts.id);
		<% /* BMIDS00939 MDC 15/11/2002 */ %>
		scScreenFunctions.SetFieldToWritable(frmScreen, frmScreen.chkShowPortedProducts.id);
		<% /* BMIDS00939 MDC 15/11/2002 - End */ %>
	}
}
//SG 29/05/02 SYS4767 END
function frmScreen.chkShowAllProductsByGroup.onclick()
{
	if (frmScreen.chkShowAllProductsByGroup.checked == true)
		scScreenFunctions.SetFieldToWritable(frmScreen, frmScreen.cboSpecialGroup.id);
	else
		scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.cboSpecialGroup.id);
}
function frmScreen.chkShowDiscountedProducts.onclick()
{
	if (frmScreen.chkShowDiscountedProducts.checked == true)
		scScreenFunctions.SetFieldToWritable(frmScreen, frmScreen.cboDiscountedPeriod.id);
	else
		scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.cboDiscountedPeriod.id);
}
function frmScreen.chkShowFixedRateProducts.onclick()
{
	if (frmScreen.chkShowFixedRateProducts.checked == true)
		scScreenFunctions.SetFieldToWritable(frmScreen, frmScreen.cboFixedRatePeriod.id);
	else
		scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.cboFixedRatePeriod.id);
}
function frmScreen.chkShowCappedFloorProducts.onclick()
{
	if (frmScreen.chkShowCappedFloorProducts.checked == true)
		scScreenFunctions.SetFieldToWritable(frmScreen, frmScreen.cboCappedFlooredPeriod.id);
	else
		scScreenFunctions.SetFieldToDisabled(frmScreen, frmScreen.cboCappedFlooredPeriod.id);
}
<%/*MO		1/08/2002	BMIDS00281 - Start */%>
function frmScreen.onsubmit()
{
	if (m_bOkButtonPressed == false){
		return false;
	}
}
<%/*MO		1/08/2002	BMIDS00281 - End */%>


<%/* AShaw	25/10/2006	EP2_8 - New method */%>
function frmScreen.chkFlexibleFilter.onclick()
{
	if (frmScreen.chkFlexibleFilter.checked == true)
		scScreenFunctions.SetRadioGroupState(frmScreen, "Flexible" , "W");
	else
		scScreenFunctions.SetRadioGroupState(frmScreen, "Flexible" , "R");
}


-->
</script>
<script src="validation.js" language="JScript" type="text/javascript"></script>
</body>
</html>
