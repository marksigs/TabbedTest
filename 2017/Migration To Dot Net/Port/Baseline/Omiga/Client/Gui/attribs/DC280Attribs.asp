<SCRIPT LANGUAGE="JScript">
<%
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description 
SD		17/10/2005	MAR209		IF the CaseStage.StageID EQUAL TO or  GREATER THAN Stage 70 (offer)
									IF OmigaUser.UserRole (for logged on user) <{new global required "INGAfterOfferLegRepChg"} THEN
									    Disable all fields on screen 
									END IF
								END IF
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
%>
<% /* Created by automated update TW 09 Oct 2002 SYS5115 */ %>

<% /* Specify screen attributes here */ %>
function SetMasks() 
{
	<% /* BMIDS00900 MDC 28/11/2002 */ %>
	frmScreen.txtDxNumber.setAttribute("filter","[0-9a-zA-Z ]");
	<% /* BMIDS00900 MDC 28/11/2002 - End */ %>
	
	//START: MAR36 - New code added by Joyce Joseph on 11-Aug-2005
	// Capitlise text fields in Screens
	frmScreen.txtCompanyName.style.textTransform = "capitalize";
	//END: MAR36
	
}

<% /* Get data required for client validation here */ %>
function GetRulesData()
{
}

<% /* Specify client validation here */ %>
function ScreenRules()
{
        return 0;
}

<% /* Specify client code for populating screen */ %>
function ClientPopulateScreen() 
{
//MAR209 Start
	SolicitorRule();
//MAR209 End	
}

//MAR 209 Start
function SolicitorRule()
{	
	var strStageSequenceNumber, strTMOfferStage, strINGAfterOfferLegRepChg, strRole;
	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	strINGAfterOfferLegRepChg = XML.GetGlobalParameterString(document, "INGAfterOfferLegRepChg");
	strStageSequenceNumber= scScreenFunctions.GetContextParameter(window,"StageSequenceNumber",null);
	strRole = scScreenFunctions.GetContextParameter(window,"Role",null);
	strTMOfferStage = XML.GetGlobalParameterAmount(document, "TMOfferStage");
	
	//alert(" strINGAfterOfferLegRepChg " +strINGAfterOfferLegRepChg);	
	//alert(" strStageSequenceNumber " + strStageSequenceNumber);
	//alert(" strRole " +strRole);
	//alert(" strTMOfferStage " + strTMOfferStage);
	
	if(parseInt(strStageSequenceNumber) >= parseInt(strTMOfferStage))
		if(parseInt(strRole) < parseInt(strINGAfterOfferLegRepChg))
			DisableFields();			
}

function DisableFields()
{	
	var item;
	for(var nLoop = 0;nLoop < frmScreen.elements.length;nLoop++)
	{
		item = frmScreen.elements(nLoop);
		item.className = 'msgReadOnly';
		item.disabled = true;
	}	
}

//MAR209 End
</SCRIPT>
