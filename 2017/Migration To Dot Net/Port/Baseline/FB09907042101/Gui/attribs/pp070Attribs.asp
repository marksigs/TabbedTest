<SCRIPT LANGUAGE="JScript">
<%
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description 
SD		19/10/2005	MAR209		IF the CaseStage.StageID EQUAL TO or  GREATER THAN Stage 70 (offer)
									IF OmigaUser.UserRole (for logged on user) <{new global required "INGAfterOfferPayeeChg"} THEN
									    Disable Add & Delete buttons
									END IF
								END IF
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
%>
<% /* Created by automated update TW 09 Oct 2002 SYS5115 */ %>

<% /* Specify screen attributes here */ %>
function SetMasks() 
{
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
	PayeeButtonRule();
//MAR209 End	
}

//MAR 209 Start
function PayeeButtonRule()
{	
	var strStageSequenceNumber, strTMOfferStage, strINGAfterOfferPayeeChg, strRole;
	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	strINGAfterOfferPayeeChg = XML.GetGlobalParameterString(document, "INGAfterOfferPayeeChg");
	strStageSequenceNumber= scScreenFunctions.GetContextParameter(window,"StageSequenceNumber",null);
	strRole = scScreenFunctions.GetContextParameter(window,"Role",null);
	strTMOfferStage = XML.GetGlobalParameterAmount(document, "TMOfferStage");
	
	//alert(" strINGAfterOfferPayeeChg " + strINGAfterOfferPayeeChg);
	//alert(" strStageSequenceNumber " + strStageSequenceNumber);
	//alert(" strRole " +strRole);
	//alert(" strTMOfferStage " + strTMOfferStage);
	
	if(parseInt(strStageSequenceNumber) >= parseInt(strTMOfferStage))
		if(parseInt(strRole) < parseInt(strINGAfterOfferPayeeChg))
		{
			//alert("here");
			frmScreen.btnAdd.disabled = true;
			frmScreen.btnDelete.disabled = true;
		}	
}



//MAR209 End

</SCRIPT>
