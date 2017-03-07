<%@ LANGUAGE="JSCRIPT" %>
<html>
<comment>
Workfile:      DC031.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Verification Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		23/11/99	Created
JLD		14/12/1999	DC/033 - Customer name field
					DC/034 - moved screen fields about
					DC/035 - If combo is 'R' and 'O' deal with it properly
					DC/037 - sorted routing out 
					DC/041 - Deletion also now catered for.
AD		30/01/2000	Rework
AY		14/02/00	Change to msgButtons button types
AY		30/03/00	New top menu/scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
CL		05/03/01	SYS1920 Read only functionality added
LD		23/05/02	SYS4727 Use cached versions of frame functions

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
</comment>

<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog    Date           Description
DB      14/03/2003     Disabled OK when clicked to ensure that multiple records are
					   not created on the Verification table when it's locked.

MC		18/05/2004		BMIDS747 - SCREEN LAYOUT MODIFIED,ABSOLUTE POSITION SPANS REPLACED
						WITH <TABLE>, VERIFICATION DETAILS SCREEN CHANGE	
MC		04/06/2004		Combos resized. (sandy's email Ref)		
HMA     13/07/2004      BMIDS747  Restrict sixe of text fields		
KRW     14/07/2004      BMIDS747  Allow 150 characters for Reason for using Tier 2 or 3 Identification
HMA     11/08/2004      BMIDS828  Add 'year4' to validate date correctly.
HMA     14/09/2004      BMIDS864  Remove ApplicationNumber and ApplicationFactFindNumber from Verification table
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT data=scClientFunctions.asp height=1 id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript"></script>

<OBJECT data=scTable.htm height=1 id=scTable 
style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 type=text/x-scriptlet 
VIEWASTEXT></OBJECT>
<OBJECT data=scTableListScroll.asp id=scScrollTable 
style="DISPLAY: none; HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" 
tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>

<%/* FORMS */%>
<form id="frmToDC035" method="post" action="DC035.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC031" method="post" action="DC031.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span><!-- Specify Screen Layout Here -->
<form id="frmScreen" mark   validate ="onchange" year4>
<div id="divCustomerName" style="HEIGHT: 40px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<TABLE style="MARGIN-LEFT: 10px; WIDTH: 594px" class="msgLabel">
		<TR>
			<TD width="20%">Customer Name
			</TD>
			<TD width="80%">
				<input id="txtCustName" name="CustomerName" maxlength="70" class="msgTxt" 
      size=36 
      style="HEIGHT: 20px; WIDTH: 244px">
			</TD>
		</TR>
	</TABLE>
</div>
<div id="divVerificationInfo" style="HEIGHT: 258px; LEFT: 10px; POSITION: absolute; TOP: 110px; WIDTH: 604px" class="msgGroup">
	<TABLE class="msgLabel" style="MARGIN-LEFT: 10px">
		<TR>
			<TD width="20%">Verification Type</TD>
			<TD width="80%">
				<SELECT id="idVerificationType" name="VerificationType" class="msgTxt" style="WIDTH: 150px">
					<OPTION value="ElectronicID" selected>Electronic ID</OPTION>
					<OPTION value="Another Type">Verify Type2</OPTION>
				</SELECT>
			</TD>
		</TR>
		<TR>
			<TD>ID Type</TD>
			<TD>
				<SELECT id="idIDType" name="IDType" class="msgTxt" style="WIDTH: 150px" disabled>
					<OPTION value="EV Type" selected>EV Type</OPTION>
					<OPTION value="ID Type2">ID Type2</OPTION>
				</SELECT>
			</TD>
		</TR>
		<TR>
			<TD>Issuer</TD>
			<TD><INPUT id="idIssuer" name="Issuer" type=TextBox readonly maxlength=20 class="msgTxt" style="HEIGHT: 20px; WIDTH: 150px" size=31></INPUT></TD>
		</TR>
		<TR>
			<TD>Reference</TD>
			<TD><INPUT id="idReference" name="Reference" type=TextBox readonly maxlength=20 class="msgTxt" style="HEIGHT: 20px; WIDTH: 150px" size=31></INPUT></TD>
		</TR>
		<TR>
			<TD>Date</TD>
			<TD><INPUT id="idVerificationDate" name="VerificationDate" readonly maxlength=10 type=TextBox class="msgTxt" style="HEIGHT: 20px; WIDTH: 122px" size=15 
     ></INPUT></TD>
		</TR>
		<TR>
			<TD valign=top>Reason for using Tier 2 or 3 Identification</TD>
			<TD><TEXTAREA class=msgTxt cols=25 id=idReasonForT2orT3 name=ReasonForT2orT3 readOnly rows=1 style="HEIGHT: 71px; WIDTH: 325px"></TEXTAREA>
			</TD>
		</TR>
	</TABLE>
</div>		
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 380px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" --> 
</div>


<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/DC031Attribs.asp" -->

<%/* CODE */%>
<script language="JScript">
<!--
	var m_sMetaAction = "";
	var m_sReadOnly = "";
	var m_sCurrentCustomerNumber = "";
	var m_sCurrentCustomerVersion = "";
	var m_sApplicationNumber = "";
	var m_sApplicationFactFindNumber = "";
	var m_sVerificationSeqNo = ""
	var scScreenFunctions;

	var m_sPersonalSeq1 = "";
	var m_sPersonalSeq2 = "";
	var m_sResidencySeq1 = "";
	var m_sResidencySeq2 = "";
	var m_blnReadOnly = false;

	var ListXML=null;

	<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
	var scClientScreenFunctions;

	function window.onload()
	{
		var sButtonList =null;
		
		//scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
		GetRulesData();
		// Make the required buttons available on the bottom of the screen
		// (see msgButtons.asp for details)
		scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
		FW030SetTitles("Verification Details","DC031",scScreenFunctions);

		RetrieveContextData();
		Initialise();
		
		// Added by automated update TW 09 Oct 2002 SYS5115
		SetMasks();
		Validation_Init();
		scScreenFunctions.SetFocusToFirstField(frmScreen);
		
		//m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC031");
		m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC031");
		
		if(m_blnReadOnly==true)
		{
			m_sReadOnly="1";
		}
		
		if(m_blnReadOnly!=true)
		{
			if(m_sMetaAction == "Edit")
			{
				sButtonList = new Array("Submit","Cancel");
			}
			else
			{	
				sButtonList = new Array("Submit","Cancel","Another");
			}
		}
		else
		{
			sButtonList = new Array("Submit");
		}
		
		
		
		/*=[MC] SHOW MAIN BUTTONS*/
		ShowMainButtons(sButtonList);
		
		// Added by automated update TW 09 Oct 2002 SYS5115
		ClientPopulateScreen();
	}


	function cBoolean(vBool)
	{
		var bReturn = new Boolean(false);
		
		if(vBool=="undefined")
		{
			bReturn=false;	
		}
		vBool = vBool.toString().toLowerCase();
		
		if(vBool=="1" || vBool == 1 || vBool=="true")
		{
			bReturn=true;
		}
		return bReturn;
	}

	function PopulateCombos()
	{
		var XMLPersonal = null;
		var XMLResidential = null;
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var sGroupList = new Array("PersonalIDType", "ResidencyIDType");

		if(XML.GetComboLists(document, sGroupList))
		{
			XMLPersonal = XML.GetComboListXML("PersonalIDType");
			XMLResidential = XML.GetComboListXML("ResidencyIDType");

			var blnSuccess = true;
			blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboPersonalType1,XMLPersonal,true);
			blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboPersonalType2,XMLPersonal,true);
			blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboResidencyType1,XMLResidential,true);
			blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboResidencyType2,XMLResidential,true);

			if(blnSuccess == false)
				scScreenFunctions.SetScreenToReadOnly(frmScreen);
		}
		XML = null;		
	}

	function PopulateScreen()
	{
		//set the Customer Name
		var sCustomerName = scScreenFunctions.GetContextCustomerName(window, m_sCurrentCustomerNumber);
		frmScreen.txtCustName.value = sCustomerName;
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCustName");		
		
		if(m_sMetaAction=='Edit')
		{
			ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			ListXML.CreateRequestTag(window,null)
			ListXML.CreateActiveTag("VERIFICATION");
			
			<% /* BMIDS864  Remove ApplicationNumber and ApplicationFactFindNumber from Verification table
			ListXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
			ListXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);   */ %>
			ListXML.CreateTag("CUSTOMERNUMBER", m_sCurrentCustomerNumber);
			ListXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCurrentCustomerVersion);
			
			if(m_sVerificationSeqNo != "")
			{
				//Reminder*******************************
				//Make readonly, Sequence no must exist to retrieve right record
				ListXML.CreateTag("VERIFICATIONSEQUENCENUMBER", m_sVerificationSeqNo);
			}
		
			//Find Verification Record	
			ListXML.RunASP(document, "FindVerificationList.asp");

			var ErrorTypes = new Array("RECORDNOTFOUND");
			var sResponseArray = ListXML.CheckResponse(ErrorTypes);
		
			if (sResponseArray[0] == true || 
				sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
			{
				
				ListXML.ActiveTag = null;
				ListXML.SelectTag(null,"VERIFICATION");
				
				var verificationSeqNo = ListXML.GetTagText("VERIFICATIONSEQUENCENUMBER");
				var verificationType =ListXML.GetTagText("VERIFICATIONTYPE");
				var idType	=  ListXML.GetTagText("IDENTIFICATIONTYPE");
				var issuer =ListXML.GetTagText("ISSUER");
				var reference = ListXML.GetTagText("REFERENCE");
				var verificationDate = ListXML.GetTagText("VERIFICATIONDATE");
				
				<% /* BMIDS864 Remove ApplicationNumber and ApplicationFactFindNumber
				var appNo =ListXML.GetTagText("APPLICATIONNUMBER");
				var appFactFindNo =ListXML.GetTagText("APPLICATIONFACTFINDNUMBER");  */ %>
				
				var custNo =ListXML.GetTagText("CUSTOMERNUMBER");
				var custVersionNo =ListXML.GetTagText("CUSTOMERVERSIONNUMBER");
				var reasonT1T3 = ListXML.GetTagText("OTHERIDENTIFICATION");
			
				if(verificationSeqNo!=null && verificationSeqNo!='')
				{
					frmScreen.idVerificationType.value = verificationType;
					frmScreen.idVerificationType.onchange();
				}
				
				
				frmScreen.idReference.value = reference;
				frmScreen.idVerificationDate.value=verificationDate;
				frmScreen.idReasonForT2orT3.value = reasonT1T3;
				frmScreen.idIssuer.value = issuer;
				if(idType!=null && idType!='')
				{
					frmScreen.idIDType.value = idType;
				}
			}
		}
		ErrorTypes = null;
		ErrorReturn = null;	
	}

	function Initialise()
	{
		
		//[MC]Populate Verification Type Combo BMIDS0747 -CC063
		PopulateComboEx("VerVerificationType",frmScreen.idVerificationType,1);
		removeComboOptions(frmScreen.idIDType,1);
		PopulateScreen();

		if(m_sReadOnly == "1")
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
		else
			frmScreen.idVerificationType.focus();			
	}

	function frmScreen.idVerificationType.onchange()
	{
		var selVerificationType =frmScreen.idVerificationType.options[frmScreen.idVerificationType.selectedIndex].text;
		//*=Remove white space characters from the selected verification type string
		selVerificationType = selVerificationType.replace(/\s+/,"");
		removeComboOptions(frmScreen.idIDType,0);
		
		if(selVerificationType!="<SELECT>")
		{
			clearData();
			enableElements();
			PopulateComboEx("Ver" + selVerificationType + "IDType",frmScreen.idIDType);
		}
		else
		{
			clearData();
			disableElements();
			
		}
	}

	function clearData()
	{
		frmScreen.idIDType.value="";
		frmScreen.idIssuer.value="";
		frmScreen.idReference.value="";
		frmScreen.idReasonForT2orT3.value="";
		frmScreen.idVerificationDate.value="";
	}


	function disableElements()
	{
		//exclude verification type from this list,
		frmScreen.idIDType.disabled=true;
		frmScreen.idIssuer.readOnly=true;
		frmScreen.idReference.readOnly=true;
		frmScreen.idReasonForT2orT3.readOnly=true;
		frmScreen.idVerificationDate.readOnly=true;
	}
	function enableElements()
	{
		frmScreen.idIDType.disabled=false;
		frmScreen.idIssuer.readOnly=false;
		frmScreen.idReference.readOnly=false;
		frmScreen.idReasonForT2orT3.readOnly=false;
		frmScreen.idVerificationDate.readOnly=false;
	}

	function onComboChange(ComboId, SpanId, TxtRefId, TxtOtherId)
	{
		var selIndex = ComboId.selectedIndex;
		if(selIndex != -1 &&
		   scScreenFunctions.IsOptionValidationType(ComboId,selIndex,"O"))
			scScreenFunctions.ShowCollection(SpanId);
		else
		{
			scScreenFunctions.HideCollection(SpanId);
			TxtOtherId.value = "";
		}

		if(selIndex != -1 &&
		   scScreenFunctions.IsOptionValidationType(ComboId,selIndex,"R"))
			//reference type, so enable reference field
			scScreenFunctions.SetFieldToWritable(frmScreen, TxtRefId.id);
		else
		{
			TxtRefId.value = "";
			scScreenFunctions.SetFieldToReadOnly(frmScreen, TxtRefId.id);
		}				
	}

	function spnToFirstField.onfocus()
	{
		scScreenFunctions.SetFocusToFirstField(frmScreen);
	}

	function spnToLastField.onfocus()
	{
		scScreenFunctions.SetFocusToLastField(frmScreen);
	}

	function RetrieveContextData()
	{
		m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
		m_sReadOnly	= scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
		m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","2631");
		m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
		m_sCurrentCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber","1007");
		m_sCurrentCustomerVersion = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber","1");
		m_sVerificationSeqNo= scScreenFunctions.GetContextParameter(window,"idVerificationSeqNo","1");
		
	}

	function AddIDToRequest( XML, tagActiveTag, sSeqNum, sVerificationType,
						     sComboValue, sOtherValue, sRefValue )
	{
		if(sSeqNum != "" || sVerificationType != "")
		{
			//Need to do some sort of update/create/delete. If Both of these are null 
			//then assume that the section has been left blank. ie. Nothing to update 
			//and nothing to create
			XML.ActiveTag = tagActiveTag; 
			XML.CreateActiveTag("VERIFICATION");
			
			<% /* BMIDS864  Remove ApplicationNumber and ApplicationFactFindNumber 
			XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
			XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber); */ %>
			
			XML.CreateTag("CUSTOMERNUMBER", m_sCurrentCustomerNumber);
			XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCurrentCustomerVersion);
			if(sSeqNum != "")
				XML.CreateTag("VERIFICATIONSEQUENCENUMBER", sSeqNum);
			if(sVerificationType != "")
			{
				// not doing a delete, so include the non-primary fields
				XML.CreateTag("VERIFICATIONTYPE", sVerificationType);
				XML.CreateTag("IDENTIFICATIONTYPE", sComboValue);
				XML.CreateTag("OTHERIDENTIFICATION", sOtherValue);
				XML.CreateTag("REFERENCE", sRefValue);		
			}
		}
	}

	function AddVerificationToRequest(XML, tagActiveTag, sSeqNum, sVerificationType, sIDType,Reference,Issuer,VerificationDate,ReasonT1T3)
	{
		if(sSeqNum != "" || sVerificationType != "")
		{
			//Need to do some sort of update/create/delete. If Both of these are null 
			//then assume that the section has been left blank. ie. Nothing to update 
			//and nothing to create
			XML.ActiveTag = tagActiveTag; 
			
			<% /* BMIDS864  Remove ApplicationNumber and ApplicationFactFindNumber 
			XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
			XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);  */ %>
			
			XML.CreateTag("CUSTOMERNUMBER", m_sCurrentCustomerNumber);
			XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCurrentCustomerVersion);
			if(sSeqNum != "")
				XML.CreateTag("VERIFICATIONSEQUENCENUMBER", sSeqNum);
			if(sVerificationType != "")
			{
				// not doing a delete, so include the non-primary fields
				XML.CreateTag("VERIFICATIONTYPE", sVerificationType);
				XML.CreateTag("IDENTIFICATIONTYPE", sIDType);
				XML.CreateTag("OTHERIDENTIFICATION", ReasonT1T3);
				XML.CreateTag("REFERENCE", Reference);
				XML.CreateTag("ISSUER", Issuer);
				XML.CreateTag("VERIFICATIONDATE", VerificationDate);
				
			}
		}
	}


	function btnSubmit.onclick()
	{
		var bSuccess = frmScreen.onsubmit();
		/*=
			[MC]If Verification list item not selected or data not inputed, cancel the action 
				and reroute to DC035
		*/
		if(!doesVerificationSelected())
		{
			cancelRequestedAction();
			return;
		}
		
		//save any changed data
		if (m_sReadOnly != "1" && bSuccess) 
		{
			if (IsChanged() == true )
			{
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				
				if(m_sMetaAction=="Edit")
				{
					//Verification Update Request is 3
					XML.CreateRequestTag(window,3);			
				}
				else
				{
					//Null represents a create new
					XML.CreateRequestTag(window,null);
				}
				
				var tagVERIFICATION = XML.CreateActiveTag("VERIFICATION");
				var sType = "";
				var sVerificationType="";
				
				if(!frmScreen.idVerificationType.selectedIndex<=0)
				{
					sVerificationType = frmScreen.idVerificationType.value;
				}
				
				var sIDType = "";
				if(!frmScreen.idIDType.selectedIndex<=0)
				{
					sIDType = frmScreen.idIDType.value;
				}
				
				var reasonT1T3 = frmScreen.idReasonForT2orT3.value;
				var reference = frmScreen.idReference.value;
				var issuer = frmScreen.idIssuer.value;
				var verificationDate = frmScreen.idVerificationDate.value;
				
				if(m_sVerificationSeqNo!="")
				{
					AddVerificationToRequest(XML,tagVERIFICATION,m_sVerificationSeqNo,sVerificationType,sIDType,reference,issuer,verificationDate,reasonT1T3);
				}
				else
				{
					AddVerificationToRequest(XML,tagVERIFICATION,"",sVerificationType,sIDType,reference,issuer,verificationDate,reasonT1T3);
				}
				
				
				switch (ScreenRules())
				{
					case 1: // Warning
					case 0: // OK
						XML.RunASP(document, "SaveVerification.asp");
						//DB BM0385 - Disable OK button 
						btnSubmit.disabled = true;
						break;
					default: // Error
						XML.SetErrorResponse();
				}

				bSuccess = XML.IsResponseOK();
			}
		}

		if(bSuccess)
		{
			//set metaAction
			scScreenFunctions.SetContextParameter(window,"idMetaAction","FromDC031");
			scScreenFunctions.SetContextParameter(window,"idCustomerNumber",m_sCurrentCustomerNumber);
			scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber",m_sCurrentCustomerVersion);
			scScreenFunctions.SetContextParameter(window,"idVerificationSeqNo","");
			frmToDC035.submit();		
		}
	}

	function cancelRequestedAction()
	{
		//Cancel and requested action.
		scScreenFunctions.SetContextParameter(window,"idMetaAction","FromDC031");
		scScreenFunctions.SetContextParameter(window,"idCustomerNumber",m_sCurrentCustomerNumber);
		scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber",m_sCurrentCustomerVersion);
		scScreenFunctions.SetContextParameter(window,"idVerificationSeqNo","");
		frmToDC035.submit();
	}

	function btnCancel.onclick()
	{
		cancelRequestedAction();
	}

	function btnAnother.onclick()
	{
		if (frmScreen.onsubmit() == true)
		{
			if(commitScreen() == true)
			{
				//Set relevant context parameters
				scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
				scScreenFunctions.SetContextParameter(window,"idCustomerNumber",m_sCurrentCustomerNumber);
				scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber",m_sCurrentCustomerVersion);
				scScreenFunctions.SetContextParameter(window,"idVerificationSeqNo","");
				frmToDC031.submit();	
			}
		}
	}

	function doesVerificationSelected()
	{
		var bReturn=false;
		if(frmScreen.idVerificationType.selectedIndex>0)
		{
			bReturn=true;
		}
		
		return bReturn;
	}
	function commitScreen()
	{
		var bSuccess=false;
		var sType = "";
		var sVerificationType="";
		var sIDType = "";	
		var reasonT1T3 = "";
		var reference = "";
		var issuer = "";
		var verificationDate="";

		//save any changed data
		if (m_sReadOnly != "1") 
		{
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

			if(m_sMetaAction=="Edit")
			{
				//Verification Update Request is 3
				XML.CreateRequestTag(window,3);			
			}
			else
			{
				//Null represents a create new
				XML.CreateRequestTag(window,null);
			}
				
			var tagVERIFICATION = XML.CreateActiveTag("VERIFICATION");
			
			if(!frmScreen.idVerificationType.selectedIndex<=0)
			{
				sVerificationType = frmScreen.idVerificationType.value;
			}
			if(!frmScreen.idIDType.selectedIndex<=0)
			{
				sIDType = frmScreen.idIDType.value;
			}
				
			reasonT1T3 = frmScreen.idReasonForT2orT3.value;
			reference = frmScreen.idReference.value;
			issuer = frmScreen.idIssuer.value;
			verificationDate = frmScreen.idVerificationDate.value;
			
				
			if(m_sVerificationSeqNo!="")
			{
				AddVerificationToRequest(XML,tagVERIFICATION,m_sVerificationSeqNo,sVerificationType,sIDType,reference,issuer,verificationDate,reasonT1T3);
			}
			else
			{
				AddVerificationToRequest(XML,tagVERIFICATION,"",sVerificationType,sIDType,reference,issuer,verificationDate,reasonT1T3);
			}
				
				
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
					if(m_sMetaAction=="Edit")
					{	
						XML.RunASP(document,"UpdateVerification.asp");
					}
					else
					{
						XML.RunASP(document, "SaveVerification.asp");
					}
							
					//DB BM0385 - Disable OK button 
					btnSubmit.disabled = true;
					break;
				default: // Error
					XML.SetErrorResponse();
				}

			bSuccess = XML.IsResponseOK();
			//alert(bSuccess);
		}

		return bSuccess;
	}


	function removeComboOptions(comboElement,lDefaultSelect)
	{
		if(comboElement!=null)
		{
			while(comboElement.options.length > 0) 
			{
				comboElement.remove(0);
			}
			
			if(lDefaultSelect==1)
			{
				var TagOPTION = document.createElement("OPTION");
						
				TagOPTION.value = "0";
				TagOPTION.text = "<SELECT>";
				comboElement.options.add(TagOPTION);
			}
		}
	}
	
	function frmScreen.idReasonForT2orT3.onkeyup()
	{	
		scScreenFunctions.RestrictLength(frmScreen, "idReasonForT2orT3", 150, true); // KRW Changed from 50 to 150 BMIDS747
	}
	
-->
</script>
</body>
</html>




