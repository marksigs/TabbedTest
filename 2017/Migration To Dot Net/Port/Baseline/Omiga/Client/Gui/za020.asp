<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /*
Workfile:      za020.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   This frame is used to find a name and address in the
				Omiga 4 names and address directory. This screen may be
				called when more than one name and address has been returned
				on a search from the calling screen, or may be called by the
				calling screen to perform all the searches.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DPol	23/11/99	Created.
AD		02/02/2000	Rework
IW		14/03/2000	SYS0455 - Town Length increased to 40
AD		21/03/2000	SYS0316 - Clear table at each search
AD		21/03/2000	SYS0025 - Table accepts double-clicking selection
AY		03/04/00	scScreenFunctions change
MC		27/04/2000	SYS0513 - Add Sortcode for Bank/Building Society
MH		16/05/2000  SYS0698
BG		17/05/00	SYS0752 Removed Tooltips 
BG		18/05/00	SYS0752 Changed Search Button label to Directory Search
GD		10/03/01	Added functionality to take in different XML from Valuation Screens
JR		18/09/01	Omiplus24 TelephoneNumber changes to include Area/Country Code
MC		08/10/01	SYS2781 - Enable Client Customisation
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

DPF		13/09/2002	APWP1/BM007 - Enable fields if Directory type is 'VAL'
MV		26/09/2002	BMIDS00522	Removed alert('Blah')
ASu		07/10/2002	BMIDS00152	Modification to DirectorySearchVal to include Town
SA		31/10/2002	BMIDS00658	Valuer Type is the 10th param passed in.
PSC		26/11/2002	BMIDS00998	Address and Panel are optional
PSC		03/12/2002	BM0105	    Rework for performance
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>

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
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>ZA020 - Directory Search  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript"></script>

<span id="spnDirectoryListScroll">
	<span style="LEFT: 250px; POSITION: absolute; TOP: 305px">
		<OBJECT data=scTableListScroll.asp id=scDirectoryTable style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>	

<%/* FORMS */%>
<form id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange">
<div id="divBackground" style="HEIGHT: 330px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 589px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Panel Id
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtPanelId" name="PanelId" maxlength="10" style="POSITION: absolute; WIDTH: 200px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 35px" class="msgLabel">
		Company Name
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtCompanyName" name="CompanyName" maxlength="45" style="POSITION: absolute; WIDTH: 150px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 315px; POSITION: absolute; TOP: 35px" class="msgLabel">
		<label id="idTown"></label>
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtTown" name="Town" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 60px" class="msgLabel">
		<label id="idSortCode"></label>
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtSortCode" name="SortCode" maxlength="8" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 315px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Filter
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<select id="cboFilter" name="Filter" style="WIDTH: 150px" class="msgCombo">
			</select>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 100px">
		<input id="btnDirectorySearch" value="Directory Search" type="button" style="WIDTH: 94px" class="msgButton">
	</span>

	<span id="spnDirectoryTable" style="LEFT: 4px; POSITION: absolute; TOP: 130px">
		<table id="tblDirectoryTable" width="580" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="28%" class="TableHead">Name&nbsp;</td>		<td width="11%" class="TableHead">Panel Id&nbsp;</td>	<td width="42%" class="TableHead">Address&nbsp;</td>	<td class="TableHead">Telephone&nbsp;</td></tr>
			<tr id="row01">		<td width="28%" class="TableTopLeft">&nbsp;</td>		<td width="11%" class="TableTopCenter">&nbsp;</td>	<td width="42%" class="TableTopCenter">&nbsp;</td>		<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		<td width="28%" class="TableLeft">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>		<td width="42%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		<td width="28%" class="TableLeft">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>		<td width="42%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		<td width="28%" class="TableLeft">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>		<td width="42%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		<td width="28%" class="TableLeft">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>		<td width="42%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		<td width="28%" class="TableLeft">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>		<td width="42%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		<td width="28%" class="TableLeft">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>		<td width="42%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		<td width="28%" class="TableLeft">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>		<td width="42%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		<td width="28%" class="TableLeft">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>		<td width="42%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		<td width="28%" class="TableBottomLeft">&nbsp;</td>	<td width="11%" class="TableBottomCenter">&nbsp;</td>	<td width="42%" class="TableBottomCenter">&nbsp;</td>	<td class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 295px">
		<input id="btnSelect" value="Select" type="button" style="WIDTH: 60px" class="msgButton">
		<span style="LEFT: 70px; POSITION: absolute; TOP: 0px">
			<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
	</span>
</div>
</form>

<%/* SCRIPT */%>
<!-- #include FILE="includes/directorysearch.asp" -->
<!-- #include FILE="attribs/za020attribs.asp" -->
<!-- #include FILE="customise/za020customise.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = null;
var AttributeArray = null;
var m_XMLResults;
var m_sDirectoryType = null;
var m_iTableLength = 10;
var scScreenFunctions;
var m_blnIsFromValuation = false;
var m_sValuationType;
var m_BaseNonPopupWindow = null;
var m_sValuerType;

function frmScreen.btnCancel.onclick()
{
	window.close();
}

function frmScreen.btnDirectorySearch.onclick()
{
	
	if (!CheckField(frmScreen.txtPanelId,"Panel Id")) return;
	if (!CheckField(frmScreen.txtCompanyName,"Company Name")) return
	if (!CheckField(frmScreen.txtTown, "Town")) return

	m_XMLResults = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	if (m_blnIsFromValuation==false)
	{
		var nRecordsFound = DirectorySearch(m_XMLResults, AttributeArray, frmScreen.txtPanelId.value, frmScreen.txtCompanyName.value,
		frmScreen.txtTown.value, frmScreen.cboFilter.value, m_sDirectoryType, frmScreen.txtSortCode.value);
	} 
	else
	{
		<% /* ASu BMIDS00152 - Start */ %>
		var nRecordsFound = DirectorySearchVal(m_XMLResults, AttributeArray, frmScreen.txtPanelId.value, frmScreen.txtCompanyName.value, frmScreen.txtTown.value);
		<% /* ASu - End */ %>		
	}
	
	if (nRecordsFound > 0)
		PopulateDirectoryList();
	else	// if record not found is returned display message
	{
		<% /* ASu BMIDS00152 - Start */ %>
		scDirectoryTable.clear();
		<% /* ASu - End */ %>		
		alert("Unable to find any matching records on the directory for your search criteria. Please amend the search criteria and/or wildcard your search.");

		frmScreen.btnSelect.disabled = true;
		frmScreen.txtCompanyName.focus();
	}
	XML = null;
}
<% /* ASu BMIDS00152 - Start */ %>
function DirectorySearchVal(XMLPanelValuerList,AttributeArray, sPanelID, sCompanyName, sTown)
<% /* ASu - End */ %>		
{
	var iIndex;
	var nNumFound=0;
		//BUILD-UP REQUEST
		XMLPanelValuerList.CreateRequestTagFromArray(AttributeArray,null);
		XMLPanelValuerList.CreateActiveTag("VA_PANELVALUERLIST");
		XMLPanelValuerList.CreateTag("VALUATIONTYPE", m_sValuationType);
	
		<% /* ASu BMIDS00152 - Start */ %>
		if (m_sValuerType != "")
		{
			XMLPanelValuerList.CreateTag("VALUERTYPE", m_sValuerType);
		}
		<% /* ASu BMIDS00152 - End */ %>
		if (sCompanyName != "")
		{
			XMLPanelValuerList.CreateTag("COMPANYNAME", sCompanyName);
		}
		if (sPanelID != "")
		{
			XMLPanelValuerList.CreateTag("PANELID", sPanelID);
		}
		<% /* ASu BMIDS00152 - Start */ %>
		if (sTown != "")
		{
			XMLPanelValuerList.CreateTag("TOWN", sTown);			
		}
		<% /* ASu BMIDS00152 - End */ %>
		
		// 		XMLPanelValuerList.RunASP(document,"FindPanelValuerList.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XMLPanelValuerList.RunASP(document,"FindPanelValuerList.asp");
				break;
			default: // Error
				XMLPanelValuerList.SetErrorResponse();
			}


		var sErrorArray = new Array("RECORDNOTFOUND");
		var sResponseArray = XMLPanelValuerList.CheckResponse(sErrorArray);
		//CHECK RESPONSE IS OK
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = XMLPanelValuerList.CheckResponse(ErrorTypes);
			
		if (ErrorReturn[0] || (ErrorReturn[1] == ErrorTypes[0])) 
		{
			if (ErrorReturn[1] != ErrorTypes[0])
			{
				//XMLPanelValuerList.ActiveTag = null;
				var tagList = XMLPanelValuerList.CreateTagList("VA_PANELVALUERLIST");
				var nNumFound = tagList.length;
				
			}

		
		} 
	return(nNumFound);


}

function CheckField(refField,sName)
{
	var sField=refField.value ;
	if(sField.indexOf("*") == 0)
	{
		alert("You must enter at least 1 preceding character to wildcard the " + sName);
		refField.focus();
		return false;
	}
	else 
		return true;

}

function frmScreen.btnSelect.onclick()
{
	if(scDirectoryTable.getRowSelected() < 0)
	{
		alert("No Address Selected");
		return;
	}

	// flag a change, as the select button has been pressed and something has been
	// selected in the table.
	FlagChange(true);

	// return an array
	var sReturn = new Array();

	sReturn[0] = IsChanged(); // Has there been a change made

	// Get the index of the selected row
	m_XMLResults.ActiveTag = null;
	//GD
	if (m_blnIsFromValuation==false)
	{
		m_XMLResults.CreateTagList("NAMEANDADDRESSDIRECTORY");
	} else
	{
		m_XMLResults.CreateTagList("VA_PANELVALUERLIST");
	}
	


	var nRowSelected = scDirectoryTable.getOffset() + scDirectoryTable.getRowSelected();

	if(m_XMLResults.SelectTagListItem(nRowSelected-1) == true)
	{
		<% /* PSC 03/12/2002 BM0105 - Start */ %>
		var DetailsXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		var xmlRequest;
		xmlRequest = DetailsXML.CreateRequestTagFromArray(AttributeArray,null);
		DetailsXML.CreateActiveTag("NAMEANDADDRESSDIRECTORY");
		DetailsXML.CreateTag("DIRECTORYGUID",m_XMLResults.GetTagText("DIRECTORYGUID")); 

		switch (ScreenRules())
		{
			case 1: // Warning
			case 0: // OK
					DetailsXML.RunASP(document,"GetDirectoryDetails.asp");
				break;
			default: // Error
				DetailsXML.SetErrorResponse();
		}
		
		if (DetailsXML.IsResponseOK())
		{
			DetailsXML.SelectTag(null, "NAMEANDADDRESSDIRECTORY");
			sReturn[1] = DetailsXML.ActiveTag.xml; // The XML string
			window.returnValue = sReturn;
			window.close();
		}
		<% /* PSC 03/12/2002 BM0105 - End */ %>
	}	
}

function tblDirectoryTable.ondblclick()
{
	tblDirectoryTable.onclick();
	frmScreen.btnSelect.onclick();
}


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	// because this screen is a pop up it needs parameters passed
	// as context data is not available!?
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	AttributeArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	// Setting up the variables from the argument array.....
	// (1-3 contain standard context variables)
	m_sDirectoryType = AttributeArray[4];
	frmScreen.txtCompanyName.value = AttributeArray[5];
	
	<% /* PSC 03/12/2002 BM0105 - Start */ %>
	if (AttributeArray[6] == true)
		m_blnIsFromValuation = true;
	<% /* PSC 03/12/2002 BM0105 - End */ %>

	<% /* ASu BMIDS00152 - Start */ %>
	m_sValuationType = AttributeArray[7];
	//BMIDS00658 Valuer Type is the 10th param passed in
	//m_sValuerType =  AttributeArray[8];
	m_sValuerType =  AttributeArray[10];
	<% /* ASu BMIDS00152 - End */ %>
	
	
	
	SetMasks();

	// MC SYS2564/SYS2781 for client customisation
	Customise();

	Validation_Init();
	GetComboLists();
	frmScreen.cboFilter.selectedIndex = 0;

	scScreenFunctions.ShowCollection(frmScreen);

	m_XMLResults = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		
	frmScreen.btnSelect.disabled = true;
	frmScreen.txtPanelId.focus();

	//DPF 13/9/02 - APWP1/BM007 - amended field settnig decision structures

	if (m_sDirectoryType == 10) // Legal Representative 
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtPanelId");
	}
	else
	{
		// Set focus on company name as Panel Id will be disabled
		frmScreen.txtCompanyName.focus();
	}
	
	if (m_sDirectoryType == 11) // Valuer
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtPanelId");
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtCompanyName");
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtTown");
		frmScreen.txtCompanyName.value = AttributeArray[5];
		frmScreen.txtTown.value = AttributeArray[8];
		frmScreen.txtPanelId.value = AttributeArray[9];
	}
	else
	{
		// Set focus on company name as Panel Id will be disabled
		frmScreen.txtCompanyName.focus();
		<% /* ASu BMIDS00152 - Start */ %>
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtPanelId")
		<% /* ASu BMIDS00152 - End */ %>
	}
		
	<% /* Bank, Building Society or Other Lender */ %>
	if (m_sDirectoryType == 3)
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtSortCode");
	else
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtSortCode");
	
	
	
	// Do a search if companyname supplied
	if (frmScreen.txtCompanyName.value !="")
		frmScreen.btnDirectorySearch.onclick();
			
	// return null if OK is not pressed
	window.returnValue = null;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function GetAddressString()
{
	var sAddressString = null;

	sAddressString = m_XMLResults.GetTagText("POSTCODE");
	sAddressString = sAddressString + ", " + m_XMLResults.GetTagText("FLATNUMBER");
	sAddressString = sAddressString + ", " + m_XMLResults.GetTagText("BUILDINGORHOUSENAME");
	sAddressString = sAddressString + ", " + m_XMLResults.GetTagText("BUILDINGORHOUSENUMBER");
	sAddressString = sAddressString + ", " + m_XMLResults.GetTagText("STREET");
	sAddressString = sAddressString + ", " + m_XMLResults.GetTagText("DISTRICT");
	sAddressString = sAddressString + ", " + m_XMLResults.GetTagText("TOWN");
	sAddressString = sAddressString + ", " + m_XMLResults.GetTagText("COUNTY");

	return sAddressString;
}




function GetComboLists()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("DirectoryFilter");
	var bSuccess = false;

	if(XML.GetComboLists(document,sGroups))
	{
		bSuccess = true;
		bSuccess = XML.PopulateCombo(document,frmScreen.cboFilter,"DirectoryFilter", false);
	}

	if(!bSuccess) scScreenFunctions.SetScreenToReadOnly(frmScreen);
	XML = null;
}

function PopulateDirectoryList()
{
	// create a tag list based on NAMEANDADDRESSDIRECTORY or (GD) VA_PANELVALUERLIST
	
	var TagListNAMEANDADDRESSDIRECTORY, TagListNAMEANDADDRESSDIRECTORYVAL;
	TagListNAMEANDADDRESSDIRECTORY = m_XMLResults.CreateTagList("NAMEANDADDRESSDIRECTORY");
	var iNoOfNameAndAddressDirectorys = m_XMLResults.ActiveTagList.length;
	TagListNAMEANDADDRESSDIRECTORYVAL = m_XMLResults.CreateTagList("VA_PANELVALUERLIST");
	var iNoOfNameAndAddressDirectorysVal = m_XMLResults.ActiveTagList.length;


	//check there are some name and address directories to retrieve
	if ((iNoOfNameAndAddressDirectorys > 0) || (iNoOfNameAndAddressDirectorysVal > 0))
	{
		//GD Decide which list to use
		if (iNoOfNameAndAddressDirectorys > 0)
		{
			TagListNAMEANDADDRESSDIRECTORY = m_XMLResults.CreateTagList("NAMEANDADDRESSDIRECTORY");
		} 
		else
		{
			TagListNAMEANDADDRESSDIRECTORY = m_XMLResults.CreateTagList("VA_PANELVALUERLIST");
			iNoOfNameAndAddressDirectorys = iNoOfNameAndAddressDirectorysVal;
			<% /* ASu BMIDS00152 - Start 
			m_sValuationType = AttributeArray[7];
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTown")
				ASu BMIDS00152 - End */ %>
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboFilter")
		}
		//obtain the data for the list
		// initalise the table
		scDirectoryTable.initialiseTable(tblDirectoryTable, 0, "",ShowDirectoryList,m_iTableLength,iNoOfNameAndAddressDirectorys);
		scDirectoryTable.clear();
		ShowDirectoryList(0)
		// select the top line, enable the select button and set focus on it
		frmScreen.btnSelect.disabled = false;
		frmScreen.btnSelect.focus();
	}
}

function ShowDirectoryList(iStart)
{
	<% /* PSC 03/12/2002 BM0105 - Start */ %>			
	var sPanelId = null;
	var sName = null;
	var sAddress = null;
	var sTelephone = null;
	var sTelephoneExt = null;

	for (var iLoop = 0; iLoop < m_XMLResults.ActiveTagList.length && iLoop < m_iTableLength;
		iLoop++)
	{
		m_XMLResults.SelectTagListItem(iLoop + iStart);
		sName = m_XMLResults.GetTagText("COMPANYNAME");
		sAddress = GetAddressString();

		sTelephone = m_XMLResults.GetTagText("COUNTRYCODE");
		sTelephone = sTelephone + " " + m_XMLResults.GetTagText("AREACODE");
		sTelephone = sTelephone + " " + m_XMLResults.GetTagText("TELENUMBER");
		sTelephoneExt = m_XMLResults.GetTagText("EXTENSIONNUMBER");
		
		if (sTelephoneExt != "")
			sTelephone = sTelephone + " ext. " + sTelephoneExt;	
			
		sPanelId = m_XMLResults.GetTagText("PANELID");
		ShowRow(iLoop+1, sPanelId, sName, sAddress, sTelephone);
	}
	<% /* PSC 03/12/2002 BM0105 - End */ %>
}

function ShowRow(nIndex, sPanelId, sName, sAddress, sTelephone)
{
	scScreenFunctions.SizeTextToField(tblDirectoryTable.rows(nIndex).cells(0),sName);
	scScreenFunctions.SizeTextToField(tblDirectoryTable.rows(nIndex).cells(1),sPanelId);
	scScreenFunctions.SizeTextToField(tblDirectoryTable.rows(nIndex).cells(2),sAddress);
	scScreenFunctions.SizeTextToField(tblDirectoryTable.rows(nIndex).cells(3),sTelephone);
}
-->
</script>
</body>
</html>

