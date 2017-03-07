<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      DC202.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:	Retype Valuation Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AShaw	28/11/2006	EP2_2 Created screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

*/

%>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>DC202 - Retype Valuation Details  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" mark validate ="onchange">
	<div id="divBackground" style="LEFT: 10px; WIDTH: 400px; POSITION: absolute; TOP: 20px; HEIGHT: 350px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Original Valuation Date
		<span style="LEFT: 175px; POSITION: absolute; TOP: -3px">
			<input id="txtOVDate" maxlength="10" style="POSITION: absolute; WIDTH: 160px" class="msgTxt" NAME="txtOVDate">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 35px" class="msgLabel">
		Original Valuer Company name
		<span style="LEFT: 175px; POSITION: absolute; TOP: -3px">
			<input id="txtOVCoName" maxlength="40" style="POSITION: absolute; WIDTH: 160px" class="msgTxt" NAME="txtOVCoName">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 70px" class="msgLabel">
		Original Valuer Company Address
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 95px" class="msgLabel">
		Postcode
		<span style="LEFT: 175px; POSITION: absolute; TOP: -3px">
			<input id="txtOVCoPostCode" maxlength="8" style="POSITION: absolute; WIDTH: 80px" class="msgTxt" NAME="txtOVCoPostCode">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 120px" class="msgLabel">
		Flat No. / Name
		<span style="LEFT: 175px; POSITION: absolute; TOP: -3px">
			<input id="txtOVCoFlatNo" maxlength="40" style="POSITION: absolute; WIDTH: 80px" class="msgTxt" NAME="txtOVCoFlatNo">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 145px" class="msgLabel">
		House No. and Name
		<span style="LEFT: 175px; POSITION: absolute; TOP: -3px">
			<input id="txtOVCoHouseNo" maxlength="40" style="POSITION: absolute; WIDTH: 30px" class="msgTxt" NAME="txtOVCoHouseNo">
		</span>
		<span style="LEFT: 205px; POSITION: absolute; TOP: -3px">
			<input id="txtOVCoHouseName" maxlength="40" style="POSITION: absolute; WIDTH: 130px" class="msgTxt" NAME="txtOVCoHouseName">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 170px" class="msgLabel">
		Street
		<span style="LEFT: 175px; POSITION: absolute; TOP: -3px">
			<input id="txtOVCoStreet" maxlength="40" style="POSITION: absolute; WIDTH: 160px" class="msgTxt" NAME="txtOVCoStreet">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 195px" class="msgLabel">
		District
		<span style="LEFT: 175px; POSITION: absolute; TOP: -3px">
			<input id="txtOVCoDistrict" maxlength="40" style="POSITION: absolute; WIDTH: 160px" class="msgTxt" NAME="txtOVCoDistrict">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 220px" class="msgLabel">
		Town
		<span style="LEFT: 175px; POSITION: absolute; TOP: -3px">
			<input id="txtOVCoTown" maxlength="40" style="POSITION: absolute; WIDTH: 160px" class="msgTxt" NAME="txtOVCoTown">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 245px" class="msgLabel">
		County
		<span style="LEFT: 175px; POSITION: absolute; TOP: -3px">
			<input id="txtOVCoCounty" maxlength="40" style="POSITION: absolute; WIDTH: 160px" class="msgTxt" NAME="txtOVCoCounty">
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 270px" class="msgLabel">
	Country
		<span style="LEFT: 175px; POSITION: absolute; TOP: -3px">
			<select id="cboOVCoCountry" style="WIDTH: 120px" class="msgCombo" NAME="cboOVCoCountry"></select>
		</span>
	</span>
	
	<span style="TOP: 310px; LEFT: 0px; POSITION: ABSOLUTE">
		<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton" NAME="btnOK">
	</span>
	<span style="TOP: 310px; LEFT: 70px; POSITION: ABSOLUTE">
		<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton" NAME="btnCancel">
	</span>

	</div>
</form>

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/DC202attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var scScreenFunctions;
	
var RetypeValXML = null;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_BaseNonPopupWindow = null;

/* EVENTS */

function frmScreen.btnCancel.onclick()
{
	RetypeValXML = null;
	window.close();
}

function frmScreen.btnOK.onclick()
{
	if (CommitChanges())
	{
		RetypeValXML = null;
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
	RetypeValXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	RetypeValXML.LoadXML(sParameters[0]);
	m_sReadOnly	= sParameters[1];
	m_sApplicationNumber = sParameters[2];
	m_sApplicationFactFindNumber = sParameters[3];

	SetMasks();
	Validation_Init();	
	Initialise();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

/* FUNCTIONS */
function CommitChanges()
{
	var bSuccess = true;

	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
			{
				if (ValidateScreen())
					bSuccess = SaveRetypeVal();
				else
					bSuccess = false
			}
		}
		else
			bSuccess = false;
	return(bSuccess);
}


function Initialise()
// Initialises the screen
{
	PopulateScreen();
	if(m_sReadOnly == "1")
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function PopulateScreen()
// Populates the screen with details of the item selected in 
{
	RetypeValXML.SelectTag(null,"NEWPROPERTY");
	if (RetypeValXML.ActiveTag != null)
		{
			// Fill our Combo.
			ComboXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			var sGroupList = new Array("Country");
			if(ComboXML.GetComboLists(document,sGroupList))
				ComboXML.PopulateCombo(document,frmScreen.cboOVCoCountry,"Country",false);
			// Set default value
			frmScreen.cboOVCoCountry.value = ComboXML.GetComboIdForValidation("Country", "UK", null, document);

			// OK, hit it, Set the Values 
			if (RetypeValXML != null)
			{
				frmScreen.txtOVDate.value = RetypeValXML.GetTagText("DATEOFORIGINALVALUATION");
				frmScreen.txtOVCoName.value = RetypeValXML.GetTagText("ORIGINALVALUERCOMPANYNAME"); 
				frmScreen.txtOVCoPostCode.value = RetypeValXML.GetTagText("ORIGINALVALUERCOMPANYPOSTCODE");
				frmScreen.txtOVCoFlatNo.value =  RetypeValXML.GetTagText("ORIGINALVALUERCOMPANYFLATNONAME"); 
				frmScreen.txtOVCoHouseNo.value = RetypeValXML.GetTagText("ORIGINALVALUERCOMPANYBUILDINGORHOUSENUMBER");
				frmScreen.txtOVCoHouseName.value = RetypeValXML.GetTagText("ORIGINALVALUERCOMPANYBUILDINGORHOUSENAME");
				frmScreen.txtOVCoStreet.value = RetypeValXML.GetTagText("ORIGINALVALUERCOMPANYSTREET"); 
				frmScreen.txtOVCoDistrict.value = RetypeValXML.GetTagText("ORIGINALVALUERCOMPANYDISTRICT");
				frmScreen.txtOVCoTown.value = RetypeValXML.GetTagText("ORIGINALVALUERCOMPANYTOWN"); 
				frmScreen.txtOVCoCounty.value = RetypeValXML.GetTagText("ORIGINALVALUERCOMPANYCOUNTY");
				if (RetypeValXML.GetTagText("ORIGINALVALUERCOMPANYCOUNTRY") != "")
					frmScreen.cboOVCoCountry.value = RetypeValXML.GetTagText("ORIGINALVALUERCOMPANYCOUNTRY"); 
			}
		}
}

function SaveRetypeVal()
{
	var sReturn = new Array();
	if (IsChanged())
	{
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		var roottag = XML.CreateActiveTag("ORIGINALVALUATION");
		XML.ActiveTag = roottag
		XML.CreateActiveTag("RETYPE");
		XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.SetAttribute("DATEOFORIGINALVALUATION", frmScreen.txtOVDate.value);
		XML.SetAttribute("ORIGINALVALUERCOMPANYNAME", frmScreen.txtOVCoName.value);
		XML.SetAttribute("ORIGINALVALUERCOMPANYPOSTCODE", frmScreen.txtOVCoPostCode.value);
		XML.SetAttribute("ORIGINALVALUERCOMPANYFLATNONAME", frmScreen.txtOVCoFlatNo.value);
		XML.SetAttribute("ORIGINALVALUERCOMPANYBUILDINGORHOUSENUMBER", frmScreen.txtOVCoHouseNo.value);
		XML.SetAttribute("ORIGINALVALUERCOMPANYBUILDINGORHOUSENAME", frmScreen.txtOVCoHouseName.value);
		XML.SetAttribute("ORIGINALVALUERCOMPANYSTREET", frmScreen.txtOVCoStreet.value);
		XML.SetAttribute("ORIGINALVALUERCOMPANYDISTRICT", frmScreen.txtOVCoDistrict.value);
		XML.SetAttribute("ORIGINALVALUERCOMPANYTOWN", frmScreen.txtOVCoTown.value);
		XML.SetAttribute("ORIGINALVALUERCOMPANYCOUNTY", frmScreen.txtOVCoCounty.value);
		XML.SetAttribute("ORIGINALVALUERCOMPANYCOUNTRY", frmScreen.cboOVCoCountry.value);

		RetypeValXML = XML.XMLDocument.xml;
	}
	
	sReturn[0] = IsChanged();
	sReturn[1] = RetypeValXML;
	window.returnValue	= sReturn;
	window.close();
}

function ValidateScreen()
{
	<% /* Check that the Valuation date isn't > 3 months ago. */ %>
	var OrigDate = scScreenFunctions.GetDateObjectFromString(frmScreen.txtOVDate.value);
	var CurrentDate = scScreenFunctions.GetAppServerDate(true);
	
	if(GetDiffInMonths(OrigDate, CurrentDate) > 2)
	{
		alert("Retype original valuation date cannot be more than three months ago.");
		frmScreen.txtOVDate.focus();
		return(false);
	}
	else 
		return(true);
	
}

function GetDiffInMonths(FirstDate, SecondDate)
{
// Use these if passing string in.
//	strDate1 = FirstDate.split("/"); 
//	strYear1 = strDate1[2].split(" ");
//	var date1 = new Date(strYear1[0],strDate1[1]-1,strDate1[0])
//	strDate2 = SecondDate.split("/"); 
//	strYear2 = strDate2[2].split(" ");
//	var date2 = new Date(strYear2[0],strDate2[1]-1,strDate2[0])

	var date1 = FirstDate;
	var date2 = SecondDate;
	
	// Use miliseconds difference from baseline
	var iDiffMS = date2.valueOf() - date1.valueOf();
	var dtDiff = new Date(iDiffMS);

	// Calc Difference (Also includes days adjustment - Courtesy Mark Chivers Inc.)
	var nYears  = date2.getFullYear() - date1.getFullYear();
	var TotalMonths = date2.getMonth() - date1.getMonth() + (nYears!=0 ? nYears*12 : 0);
    // Day of month adjustment. If previous DAY > latest DAY.
    if ((date1.getDate() - date2.getDate()) > 0)
    TotalMonths = TotalMonths - 1


	// Return no of months.
	return TotalMonths;
}

function frmScreen.txtOVCoFlatNo.onchange()
{
	// Set viscond for txtOVCoHouseNo and txtOVCoHouseName
	if (frmScreen.txtOVCoFlatNo.value != "")
	{
		frmScreen.txtOVCoHouseNo.removeAttribute("required");
		frmScreen.txtOVCoHouseName.removeAttribute("required");
	}
	else
	{
		frmScreen.txtOVCoHouseNo.setAttribute("required", "true");
		frmScreen.txtOVCoHouseName.setAttribute("required", "true");	
	}
}

function frmScreen.txtOVCoHouseNo.onchange()
{
	// Set viscond for txtOVCoFlatNo and txtOVCoHouseName
	if (frmScreen.txtOVCoHouseNo.value != "")
	{
		frmScreen.txtOVCoFlatNo.removeAttribute("required");
		frmScreen.txtOVCoHouseName.removeAttribute("required");
	}
	else
	{
		frmScreen.txtOVCoFlatNo.setAttribute("required", "true");
		frmScreen.txtOVCoHouseName.setAttribute("required", "true");
	}
}

function frmScreen.txtOVCoHouseName.onchange()
{
	// Set viscond for txtOVCoHouseNo and txtOVCoFlatNo
	if (frmScreen.txtOVCoHouseName.value != "")
	{
		frmScreen.txtOVCoHouseNo.removeAttribute("required");
		frmScreen.txtOVCoFlatNo.removeAttribute("required");
	}
	else
	{
		frmScreen.txtOVCoHouseNo.setAttribute("required", "true");
		frmScreen.txtOVCoFlatNo.setAttribute("required", "true");	
	}
}


-->
</script>
</body>
</html>

