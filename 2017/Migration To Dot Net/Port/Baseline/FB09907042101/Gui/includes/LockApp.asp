<%
/*
Workfile:      LockApp.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Contains shared functions (currently between GN215 and TM020)
				which lock a particular application and all associated
				customers. Also sets up the appropriate context.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		05/01/01	Created
APS		28/02/01	SYS1990 Modifications to not route when and error is returned 
					and to clear the context out on an error
APS		03/03/01	SYS1993	Set up the context for Application Priority data
APS		04/03/01	SYS1920 Set up FreezeDataIndicator in context
SA		01/06/01	SYS2273	Context for CustomerNumber and CustomerVersionNumber incorrectly set.
LD		23/05/02	SYS4727 Use cached versions of frame functions
DB		29/05/02	SYS4767 MSMS to Core Integration
SG		31/05/02	SYS4802 Correct error caused by SYS4767
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

MV		04/07/2002	BMIDS00179 - Reverse Core Changes Ref AQR SYS4876 
DRC     13/01/2004  BMIDS670 - Changed IsApplicationatOffer to FreezeUnfreeze
DRC     03/02/2004  BMIDS692 Data freeze overrides
MC		06/07/2004	BMIDS763	- ApplicationDate value added to context parameter
GHun	29/07/2004	BMIDS821 Clear ApplicationDate and changed where it gets set in context
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

PSC		13/10/2005	MAR57	Add Customer Categories, LaunchCustomerXML and ExistingCustomerLaunch
PSC		25/10/2005	MAR300	Add Customer Synchronisation
GHun	27/02/2006	MAR753	Non processing units should not create locks
PSC		09/03/2006	MAR1353 Amend SynchroniseCustomerData so it doesn't run if the 
                            application is in an exception stage or at completion stage
GHun	05/04/2006	MAR1300	ChangeOfProperty changes
GHun	10/04/2006	MAR1604 Changed CheckChangeOfProperty to ignore RecordNotFound
JD		07/11/2006	MAR1957 MARS043 added CaseStageSequenceNo to criticaldatacontext passed to GetAndSynch.
							Also called FindCustomerForApplication after GetAndSynch to set context correctly.
HMA		04/01/2007  MAR2034 Display customer category correctly.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:
AS		18/01/2007	EP1301 Set correspondence salutation in context.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<!-- #include FILE="../Attribs/OverrideAttribs.asp" -->
<script language="JavaScript" type="text/javascript">

function LockApplication(sReadOnly, sApplicationNumber, sApplicationFactFindNumber, 
							sCustomerNumber, sChannelId, blnSyncCustomerData)
{
	var blnContinue = true;
	var blnRoute = true;
	<% /* MAR753 GHun Only lock if this is a processing unit */ %>
	var scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	var blnProcessingInd = scScreenFunctions.GetContextParameter(window, "idProcessingIndicator", null);

	if ((sReadOnly != "1") && (blnProcessingInd == 1))
	{
	<% /* MAR753 End */ %>
		
		// if the calling object already knows a customer is ready only then don't try to lock out the other customers and application
		
		<%/* SG 31/05/02 SYS4802 Reversed SYS4767 */%>
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
		<%
		//DB SYS4767 - MSMS Integration
		//var XML = new scXMLFunctions.XMLObject();
		//DB End
		//SG 31/05/02 SYS4802 END
		%>
		
		XML.CreateRequestTag(window,null);
		XML.CreateTag("APPLICATIONNUMBER", sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
		
		if (sCustomerNumber != null || sCustomerNumber != ""){
			XML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
		}		
		if (sChannelId != null || sChannelId != ""){
			XML.CreateTag("CHANNELID", sChannelId);
		}
		
		XML.RunASP(document,"LockCustomersForApplication.asp");

		if (XML.IsResponseOK()) {
			
			if (XML.GetAttribute("READONLY") == "")
				sReadOnly = "0";
			else
				sReadOnly = "1";

			if (sReadOnly == "1") {
				var sMessage = XML.GetTagText("MESSAGE");
				blnContinue = confirm(sMessage);

				if (blnContinue != true) {
					blnRoute = false;	
				}
			}
		}
		XML = null;	
	}		

	if (blnRoute == true) {

		blnRoute = LoadApplication(sApplicationNumber, sApplicationFactFindNumber);
						
		<% // not every record at the moment has a Application Prority record so will not add defensive coding to this statement at the moment %>
		if (blnRoute == true){ 
			GetApplicationPriority(sApplicationNumber);
									
			<% // last thing we do is put information into context %>
			scScreenFunctions.SetContextParameter(window,"idCustomerReadOnly",sReadOnly);
			scScreenFunctions.SetContextParameter(window,"idReadOnly",sReadOnly);
			scScreenFunctions.SetContextParameter(window,"idApplicationNumber",sApplicationNumber);
			scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber",sApplicationFactFindNumber);
			
			<% /* PSC 25/10/2005 MAR300 - Start */ %>
			if (sReadOnly != "1" && blnSyncCustomerData != null && blnSyncCustomerData)
				SynchroniseCustomerData(sApplicationNumber, sApplicationFactFindNumber);
			<% /* PSC 25/10/2005 MAR300 - End */ %>
			
			<% /* BMIDS821 GHun setting ApplicationDate is triggered by setting ApplicationNumber
			in context as GetApplicationData is called there already (see scScreenFunctions) */
			/*[MC]APPLICATIONDATE DATA TO SET ON THE CONTEXT*/
			/* try
			{
				//APPLICATION DATE 
				var xmlAPPDATE = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
			
				xmlAPPDATE.CreateRequestTag(window, null)
				xmlAPPDATE.CreateActiveTag("APPLICATION");
				xmlAPPDATE.CreateTag("APPLICATIONNUMBER", sApplicationNumber);
				xmlAPPDATE.CreateTag("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
				xmlAPPDATE.CreateTag("CALCSDATAONLY","1");
			
				xmlAPPDATE.RunASP(document,"GetApplicationData.asp");
			
				if(xmlAPPDATE.IsResponseOK())
				{
					
					//alert(xmlAPPDATE.XMLDocument.xml);
					xmlAPPDATE.SelectTag(null, "APPLICATIONFACTFIND");
					var sAppDate = xmlAPPDATE.GetTagText("APPLICATIONDATE")
					scScreenFunctions.SetContextParameter(window,"idApplicationDate",sAppDate);
				}
			}
			catch(ex)
			{
				//alert(ex);
			}
			*/
			/*[MC]SECTION END*/
			%>
			CheckChangeOfProperty(sApplicationNumber, sApplicationFactFindNumber);	<% /* MAR1300 GHun */ %>
		}
	}
	
	if (blnRoute != true) {
		ClearContext();
	}
	
	<%
	// APS SYS1993 Changed LockApplication to return variable 'iRoute' which is =
	// 0 if User decides NOT to route on READ ONLY Customer or Application 
	// 1 if User decides to Route on READ ONLY Customer or Applcation
	// 2 if Application and Customer are NOT READ ONLY
	
	// In JScript 1 = true and everything else = false
	%>
	
	if (blnRoute != true) iRoute = 0;
	else if (sReadOnly == "1") iRoute = 1;
	else if (sReadOnly != "1") iRoute = 2;
	
	return (iRoute);
}

function LoadApplication(sApplicationNumber, sApplicationFactFindNumber) {

	var blnContinue = false;
	
	blnContinue = FindCustomersForApplication(sApplicationNumber, sApplicationFactFindNumber);	
	if (blnContinue && (blnContinue = FindApplicationData(sApplicationNumber,sApplicationFactFindNumber)));
	if (blnContinue && (blnContinue = FindCurrentApplicationStage(sApplicationNumber, sApplicationFactFindNumber)));
	if (blnContinue && (blnContinue = IsApplicationUnderReview(sApplicationNumber,sApplicationFactFindNumber)));
	if (blnContinue && (blnContinue = CheckFreezeunFreeze(sApplicationNumber)));
	
	return blnContinue;
}

function FindCustomersForApplication(sApplicationNumber, sApplicationFactFindNumber)
{
	var blnSuccess = false;
	
	<%/* SG 31/05/02 SYS4802 Reversed SYS4767 */%>
	<%/* DB SYS4767 - MSMS Integration */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ComboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	<%/*var XML = new scXMLFunctions.XMLObject(); */%>
	<%/*DB End*/%>
	
	XML.CreateRequestTag(window,"SEARCH");
	XML.CreateActiveTag("APPLICATION");
	XML.CreateTag("APPLICATIONNUMBER",sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
	XML.RunASP(document,"FindCustomersForApplication.asp");

	if (XML.IsResponseOK()) {

		XML.CreateTagList("CUSTOMERROLE");
		
		var sCustomerNumber = "";
		var sCustomerVersionNumber = "";
		var sCustomerRoleType = "";
		var sCustomerOrder = "";
		var sSurname = "";
		var sForename = "";
		var iCustomerIndex = 0;
		
		ComboXML.GetComboLists(document,["CustomerCategory"]);
		
		for (var i0=0; i0<5; i0++)
		{
			sCustomerNumber = "";
			sCustomerVersionNumber = "";
			sCustomerRoleType = "";
			sCustomerOrder = "";
			sSurname = "";
			sForename = "";
			<% /* PSC 13/10/2005 MAR57 */ %>
			sCategory = "";	
			<% /* PSC 25/10/2005 MAR300 */ %>
			sOtherSystemNumber = "";	
			
			if (i0 < XML.ActiveTagList.length){
				if (XML.SelectTagListItem(i0) == true){
					sCustomerNumber = XML.GetTagText("CUSTOMERNUMBER");
					sCustomerVersionNumber = XML.GetTagText("CUSTOMERVERSIONNUMBER");
					sCustomerRoleType = XML.GetTagText("CUSTOMERROLETYPE");
					sCustomerOrder = XML.GetTagText("CUSTOMERORDER");				
					sSurname = XML.GetTagText("SURNAME");
					sForename = XML.GetTagText("FIRSTFORENAME");
					<% /* PSC 13/10/2005 MAR57 */ %>
					sCategory = XML.GetTagText("CUSTOMERCATEGORY");
					
					<% /* MAR2034 */ %>
					if (sCategory != "")
					{
						var XMLCombos = ComboXML.GetComboListXML("CustomerCategory"); 
						sCategory = XMLCombos.selectSingleNode(".//LISTENTRY[GROUPNAME = 'CustomerCategory' $and$ VALUEID = '" + sCategory + "']/VALUENAME").text; 
					}
					
					<% /* PSC 25/10/2005 MAR300 */ %>
					sOtherSystemNumber = XML.GetTagText("OTHERSYSTEMCUSTOMERNUMBER");
				}
			}
			if (i0==0){
				<%
				//SYS2273 SA 
				//scScreenFunctions.SetContextParameter(window,"idCustomerNumber", sForename + " " + sSurname);
				//scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber", sCustomerNumber);
				%>
				scScreenFunctions.SetContextParameter(window,"idCustomerNumber", sCustomerNumber);
				scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber",sCustomerVersionNumber );						
				scScreenFunctions.SetContextParameter(window,"idCustomerName",sForename + " " + sSurname);						
			}
			iCustomerIndex = i0 + 1;						
			scScreenFunctions.SetContextParameter(window,"idCustomerName" + iCustomerIndex, sForename + " " + sSurname);
			scScreenFunctions.SetContextParameter(window,"idCustomerNumber" + iCustomerIndex, sCustomerNumber);
			scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber" + iCustomerIndex, sCustomerVersionNumber);
			scScreenFunctions.SetContextParameter(window,"idCustomerRoleType" + iCustomerIndex, sCustomerRoleType);
			scScreenFunctions.SetContextParameter(window,"idCustomerOrder" + iCustomerIndex, sCustomerOrder);
			<% /* PSC 13/10/2005 MAR57 */ %>
			scScreenFunctions.SetContextParameter(window,"idCustomerCategory" + iCustomerIndex, sCategory);
			<% /* PSC 25/10/2005 MAR300 */ %>
			scScreenFunctions.SetContextParameter(window,"idOtherSystemCustomerNumber" + iCustomerIndex, sOtherSystemNumber);
		}
		blnSuccess = true;
	}
	
	return blnSuccess;
}

function FindApplicationData(sApplicationNumber, sApplicationFactFindNumber)
{
	var blnSuccess = false;
	
	<%/* SG 31/05/02 SYS4802 Reversed SYS4767 */%>
	<%/* DB SYS4767 - MSMS Integration */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/*var XML = new scXMLFunctions.XMLObject(); */%>
	<%/*DB End*/%>

	XML.CreateRequestTag(window,"SEARCH");
	XML.CreateActiveTag("APPLICATION");
	XML.CreateTag("APPLICATIONNUMBER",sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
	XML.RunASP(document,"GetApplicationData.asp");

	if (XML.IsResponseOK()) {
		
		var sApplicationType = "";
		var sApplicationTypeNo = "";
		var xmlTag = XML.SelectSingleNode("//TYPEOFAPPLICATION");
		
		if (xmlTag != null)
		{
			sApplicationTypeNo = xmlTag.text;
			sApplicationType = XML.GetAttribute("TEXT");
		}

		<%// Set up context from xml returned %>
		scScreenFunctions.SetContextParameter(window,"idMortgageApplicationDescription",sApplicationType);
		scScreenFunctions.SetContextParameter(window,"idMortgageApplicationValue",sApplicationTypeNo);
		<%// AS 18/01/2007 EP1301 Set correspondence salutation in context. %>
		xmlTag = XML.SelectSingleNode("//CORRESPONDENCESALUTATION");		
		scScreenFunctions.SetContextParameter(window,"idCorrespondenceSalutation", xmlTag != null ? xmlTag.text : "");
		
		blnSuccess = true;
	}
	
	return blnSuccess;
}

function FindCurrentApplicationStage(sApplicationNumber, sApplicationFactFindNumber)
{
	var blnSuccess = false;
	
	<%/* SG 31/05/02 SYS4802 Reversed SYS4767 */%>
	<%/* DB SYS4767 - MSMS Integration */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/*var XML = new scXMLFunctions.XMLObject(); */%>
	<%/*DB End*/%>

	// Make a Call to GetCurrentStageData 	
	XML.CreateRequestTag(window,"SEARCH");
	XML.CreateActiveTag("APPLICATIONSTAGE");
	XML.CreateTag("APPLICATIONNUMBER",sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
	XML.RunASP(document,"GetCurrentApplicationStage.asp");
		
	if (XML.IsResponseOK()) {
			
		var sApplicationStageName = "";
		var sApplicationStageNumber = "";
		var sApplicationStageSequenceNumber= "";
			
		sApplicationStageName = XML.GetTagText("STAGENAME");
		sApplicationStageNumber = XML.GetTagText("STAGENUMBER");
		sApplicationStageSequenceNumber= XML.GetTagText("STAGESEQUENCENO");
			
		scScreenFunctions.SetContextParameter(window,"idStageName",sApplicationStageName);
		scScreenFunctions.SetContextParameter(window,"idStageId",sApplicationStageNumber);
		scScreenFunctions.SetContextParameter(window,"idCurrentStageOrigSeqNo",sApplicationStageSequenceNumber);
	
		blnSuccess = true;
	}		
	XML = null;
	
	return blnSuccess;
}

function IsApplicationUnderReview(sApplicationNumber, sApplicationFactFindNumber)
{
	var blnSuccess = false;
	
	<%/* SG 31/05/02 SYS4802 Reversed SYS4767 */%>
	<%/* DB SYS4767 - MSMS Integration */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/*var XML = new scXMLFunctions.XMLObject(); */%>
	<%/*DB End*/%>

	var sAppUnderReview = "";
	
	XML.CreateRequestTag(window,"SEARCH");
	XML.CreateActiveTag("APPLICATIONREVIEWHISTORY");
	XML.CreateTag("APPLICATIONNUMBER",sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
	XML.RunASP(document,"IsApplicationUnderReview.asp");

	if (XML.IsResponseOK()) {
		
		sAppUnderReview = XML.GetTagText("UNDERREVIEWINDICATOR");
		if (sAppUnderReview == "1")
			// Set up context from xml returned
			scScreenFunctions.SetContextParameter(window,"idAppUnderReview" ,"1");
		else
			scScreenFunctions.SetContextParameter(window,"idAppUnderReview" ,"0");
		blnSuccess = true;
	}
	XML = null;
	
	return blnSuccess;
}

function GetApplicationPriority(sApplicationNumber){
	
	var blnSuccess = false;
	
	<%/* SG 31/05/02 SYS4802 Reversed SYS4767 */%>
	<%/* DB SYS4767 - MSMS Integration */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/*var XML = new scXMLFunctions.XMLObject(); */%>
	<%/*DB End*/%>

	var sApplicationPriority = "";
	
	XML.CreateRequestTag(window,"SEARCH");
	XML.CreateActiveTag("APPLICATIONPRIORITY");
	XML.CreateTag("APPLICATIONNUMBER",sApplicationNumber);
	XML.RunASP(document,"GetApplicationPriority.asp");

	if (XML.IsResponseOK()) {
		sApplicationPriority = XML.GetTagText("APPLICATIONPRIORITYVALUE");
		// Set up context from xml returned			
		scScreenFunctions.SetContextParameter(window,"idApplicationPriority" ,sApplicationPriority);
		
		blnSuccess = true;
	}
	XML = null;
	
	return blnSuccess;
}

function ClearContext()
{
	scScreenFunctions.SetContextParameter(window,"idStageName","");
	scScreenFunctions.SetContextParameter(window,"idStageId","");
	scScreenFunctions.SetContextParameter(window,"idCurrentStageOrigSeqNo","");
	scScreenFunctions.SetContextParameter(window,"idApplicationPriority" ,"");
	scScreenFunctions.SetContextParameter(window,"idAppUnderReview" ,"");
	scScreenFunctions.SetContextParameter(window,"idMortgageApplicationDescription","");
	scScreenFunctions.SetContextParameter(window,"idMortgageApplicationValue","");
	scScreenFunctions.SetContextParameter(window,"idCustomerReadOnly","");
	scScreenFunctions.SetContextParameter(window,"idReadOnly","");	
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","");
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber","");					
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber", "");
	scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber", "");
	<% /* BMIDS821 GHun Clear ApplicationDate */ %>
	scScreenFunctions.SetContextParameter(window,"idApplicationDate", "");	
	
	for (var i=1; i<=5; i++){
		scScreenFunctions.SetContextParameter(window,"idCustomerName" + i, "");
		scScreenFunctions.SetContextParameter(window,"idCustomerNumber" + i, "");
		scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber" + i, "");
		scScreenFunctions.SetContextParameter(window,"idCustomerRoleType" + i, "");
		scScreenFunctions.SetContextParameter(window,"idCustomerOrder" + i, "");
		<% /* PSC 13/10/2005 MAR57 */ %>
		scScreenFunctions.SetContextParameter(window,"idCustomerCategory" + i, "");
		<% /* PSC 25/10/2005 MAR300 */ %>
		scScreenFunctions.SetContextParameter(window,"idOtherSystemCustomerNumber" + i, "");
	}
	
	<% /* PSC 13/10/2005 MAR57  - Start */ %>
	scScreenFunctions.SetContextParameter(window,"idLaunchCustomerXML", "");
	<% /* PSC 25/10/2005 MAR300 */ %>
	scScreenFunctions.SetContextParameter(window,"idExistingApplication", "");
	scScreenFunctions.SetContextParameter(window,"idLaunchCustomerNumber", "");
	<% /* PSC 13/10/2005 MAR57  - End */ %>
	
	scScreenFunctions.SetContextParameter(window, "idPropertyChange", "");	<% /* MAR1300 GHun */ %>
}	

<%/* DRC - Renamed function - was IsApplicationAtaOffer*/%>
function CheckFreezeunFreeze(sApplicationNumber)
{
	var blnSuccess = false;
	
	<%/* SG 31/05/02 SYS4802 Reversed SYS4767 */%>
	<%/* DB SYS4767 - MSMS Integration */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/*var XML = new scXMLFunctions.XMLObject(); */%>
	<%/*DB End*/%>
	
	// get the Activity ID from the global partameter table
	var sActivityId = scScreenFunctions.GetContextParameter(window, "idActivityId", "1");
	
	if (sActivityId == "") {
		sActivityId = XML.GetGlobalParameterAmount(document, "TMOmigaActivity");	
		scScreenFunctions.SetContextParameter(window, "idActivityId", sActivityId);
		XML.ResetXMLDocument();
	}
	
	XML.CreateRequestTag(window, "FREEZEUNFREEZEAPPLICATION");
	XML.CreateActiveTag("CASESTAGE");
	
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("CASEID", sApplicationNumber);
	XML.SetAttribute("ACTIVITYID", sActivityId);
	
	XML.RunASP(document, "OmigaTMBO.asp");
	
	if (XML.IsResponseOK()) {
		var sDataFreezeIndicator = XML.GetTagAttribute("APPLICATION", "FREEZEDATAINDICATOR");
		if (sDataFreezeIndicator.length > 0)
		{
			scScreenFunctions.SetContextParameter(window,"idFreezeDataIndicator", sDataFreezeIndicator);
			<% //BMIDS693 DRC Freeze Overrides %>
			var sFreezeTaskID = "";
			if (sDataFreezeIndicator=="1") 
				sFreezeTaskID = XML.GetTagAttribute("APPLICATION", "FREEZETASKID");
			if (ClientFreezeOverRide(sFreezeTaskID,sDataFreezeIndicator))
			 {
				<% // alert("Data Freeze Over ride for " + sFreezeTaskID); %>
			 }            
		}	
		blnSuccess = true;
	}
	XML = null;
	
	return blnSuccess;
}
function SynchroniseCustomerData(sApplicationNumber, sApplicationFactFindNumber)
{
	var blnUseCustomerSynch;
	var blnSuccess = false;

	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	<% /* PSC 09/03/2006 MAR1353 - Start */ %>
	var sActivityId = "";
	var sStageId = "";
	
	blnUseCustomerSynch = XML.GetGlobalParameterBoolean(document, "UseAdminGetCustDetailAppAccess");	

	if (blnUseCustomerSynch)
	{
		sActivityId = scScreenFunctions.GetContextParameter(window,"idActivityId", ""); 
		sStageId = scScreenFunctions.GetContextParameter(window,"idStageId", "");
		
		var ParamXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var sCompletionStageID = ParamXML.GetGlobalParameterString(document,"TMCompletionsStageId");				

		if (sStageId != sCompletionStageID)
		{		
			var XMLStages = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XMLStages.CreateRequestTag(window,"FINDSTAGELIST");
			XMLStages.CreateActiveTag("STAGE");
			XMLStages.SetAttribute("ACTIVITYID",sActivityId);
			XMLStages.SetAttribute("EXCEPTIONSTAGEINDICATOR","1");
			switch (ScreenRules())
			{
				case 1: // Warning
				case 0: // OK
					XMLStages.RunASP(document,"FindStageList.asp");
					break;
				default: // Error
					XMLStages.SetErrorResponse();
			}

			var ErrorTypes = new Array("RECORDNOTFOUND");
			var ErrorReturn = XMLStages.CheckResponse(ErrorTypes);
			if(ErrorReturn[0] && ErrorReturn[1] != ErrorTypes[0])
			{
				XMLStages.SelectSingleNode ("//RESPONSE/STAGE[@STAGEID='" + sStageId + "']");
				
				if (XMLStages.ActiveTag != null)
					blnUseCustomerSynch = false;
			}
		}
		else
		{
			blnUseCustomerSynch = false;
		}
	}
	<% /* PSC 09/03/2006 MAR1353 - End */ %>
		
	if (blnUseCustomerSynch)
	{
		var sAppPriority = scScreenFunctions.GetContextParameter(window,"idApplicationPriority", "");
		var sStageId = scScreenFunctions.GetContextParameter(window,"idStageId", "");
		var sStageSeqNo = scScreenFunctions.GetContextParameter(window,"idCurrentStageOrigSeqNo", ""); //MAR1957
		
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
		XML.CreateRequestTag(window, null);
		XML.CreateActiveTag("SEARCH");
		XML.SetAttribute("CUSTOMERDATACHECK", "1");
		XML.CreateActiveTag("CUSTOMERLIST");
	
		for (var nCustIndex=1; nCustIndex<=5; nCustIndex++)
		{
			var sCustNo = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nCustIndex, "");
			var sCustVerNo = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nCustIndex, "");
			var sOtherSystemCustomerNumber = scScreenFunctions.GetContextParameter(window,"idOtherSystemCustomerNumber" + nCustIndex, "");
			
			if (sCustNo.length > 0)
			{
				XML.CreateActiveTag("CUSTOMER");
				XML.SetAttribute("NOLOCK", "1");		
			
				XML.CreateTag("CUSTOMERNUMBER", sCustNo);
				XML.CreateTag("CUSTOMERVERSIONNUMBER", sCustVerNo);
				XML.CreateTag("OTHERSYSTEMCUSTOMERNUMBER", sOtherSystemCustomerNumber);
			}
		}
		XML.SelectTag(null, "SEARCH");
		XML.CreateActiveTag("CRITICALDATACONTEXT");
		XML.SetAttribute("APPLICATIONNUMBER", sApplicationNumber);
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
		XML.SetAttribute("APPLICATIONPRIORITY", sAppPriority);
		XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
		XML.SetAttribute("ACTIVITYID", sActivityId);
		XML.SetAttribute("ACTIVITYINSTANCE", "1");
		XML.SetAttribute("STAGEID", sStageId);
		XML.SetAttribute("CASESTAGESEQUENCENO", sStageSeqNo); //MAR1957
		XML.RunASP(document, "GetAndSyncCustomerDetails.asp");
		if(XML.IsResponseOK())
		{
			//MAR1957 otherSystemCustomerNumber may have changed. 
			// call FindCustomersForApplication again to reset the context
			blnSuccess = FindCustomersForApplication(sApplicationNumber, sApplicationFactFindNumber);
		}
	}
	else
		blnSuccess = true;
	
}

<% /* MAR1300 GHun */ %>
function CheckChangeOfProperty(sApplicationNumber, sApplicationFactFindNumber)
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window, null);
	var strPropertyChange = "";

	XML.CreateActiveTag("NEWPROPERTY");
	XML.CreateTag("APPLICATIONNUMBER", sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);

	XML.RunASP(document, "GetNewPropertyGeneral.asp");

	<% /* MAR1604 GHun Ignore RecordNotFound */ %>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)
	{
	<% /* MAR1604 End */ %>
		XML.SelectTag(null, "NEWPROPERTY");
		if (XML.GetTagText("CHANGEOFPROPERTY") == "1")
			strPropertyChange = "1";
	}
	
	scScreenFunctions.SetContextParameter(window, "idPropertyChange", strPropertyChange);
}
<% /* MAR1300 End */ %>
</script>
