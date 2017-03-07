<SCRIPT LANGUAGE="JScript">
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DC010attribs.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Application source Screen attributes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		04/11/99	Created
AD		30/01/00	Rework
DB		29/05/02	SYS4767 - MSMS to Core integration
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

MV		04/07/2002	SYS4876 - Reverse Core Changes SYS4767
MO		20/09/2002	INWP1, BMIDS00502 - Introducer changes
DRC		03/02/04	BMIDS692 Data freeze overrides
INR		13/07/04	BMIDS744 Validate fields as dates
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History :

Prog	Date		AQR			Description
pct		24/03/2005	EP15		Changes to accomodated Packager/Broker Intermediaries 
SAB		22/05/2006	EP565		Make Date Application Received mandatory
PE		06/06/2006	EP695		Add decimal point format to currency fields.
MCh		31/01/2007	EP2_226		Make Date Application Received non-mandatory
AShaw	01/02/2007	EP2_1170	Issue with comment.
SR		01/03/2007	EP2_1272
SR		14/03/2007	Ep2_1335	Make combo Channel ID mandatory.
PE		30/03/2007	EP2_1888	Make "Level of Advice" mandatory.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
*/

function SetMasks()
{
	//EP2_226 Make txtApplicationDateReceived non-mandatory
	//frmScreen.txtApplicationDateReceived.setAttribute("required", "true");
	frmScreen.cboDirectIndirect.setAttribute("required", "true");
	frmScreen.cboNatureOfLoan.setAttribute("required", "true");
	frmScreen.txtEstimatedCompletionDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtApplicationDateReceived.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtApplicationDateSigned.setAttribute("date", "DD/MM/YYYY"); //EP8 pct
	
	<% /* AW 09/05/2006 EP485 - Start */ %>
	frmScreen.txtProcurationFee.setAttribute("filter","[0-9.]");
	<% /*  SR : EP2_1272
	frmScreen.txtBrokerFee.setAttribute("filter","[0-9.]");
	frmScreen.txtBrokerFee.setAttribute("min", "0");
	frmScreen.txtBrokerFee.setAttribute("max", "99999");
	frmScreen.txtBrokerFeeRefund.setAttribute("filter","[0-9.]");
	frmScreen.txtBrokerFeeRefund.setAttribute("min", "0");
	frmScreen.txtBrokerFeeRefund.setAttribute("max", "99999");
	frmScreen.txtPackagerValFee.setAttribute("filter","[0-9.]");
	frmScreen.txtPackagerValFee.setAttribute("min", "0");
	frmScreen.txtPackagerValFee.setAttribute("max", "99999");
	frmScreen.txtPackagerValFeeRebate.setAttribute("filter","[0-9.]");
	frmScreen.txtPackagerValFeeRebate.setAttribute("min", "0");
	frmScreen.txtPackagerValFeeRebate.setAttribute("max", "99999"); */ %>
	frmScreen.txtAmtProcFeePassedOn.setAttribute("filter","[0-9.]");  
	frmScreen.txtAmtProcFeePassedOn.setAttribute("amount", ".");
	frmScreen.txtAmtProcFeePassedOn.setAttribute("min", "0");
	frmScreen.txtAmtProcFeePassedOn.setAttribute("max", "99999.99");
	frmScreen.txtAmtProcFeePassedOn.setAttribute("nym", "Amount of Proc Fee Passed on");
	frmScreen.cboChannel.setAttribute("required", "true");
	frmScreen.cboChannel.setAttribute("nym", "Channel ID");
	<% /*  SR : EP2_1272 - End */ %>
	<% /* AW 09/05/2006 EP485 - End */ %>
	
	<% /* PE 30/04/2007 EP2_1888 */ %>
	frmScreen.cboLevelOfAdvice.setAttribute("required", "true");
}

<% /* Get data required for client validation here */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
function GetRulesData()
{
}

<% /* Specify client validation here */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
function ScreenRules()
{
        return 0;
}

<% /* Specify client code for populating screen */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
function ClientPopulateScreen() 
{
<% /* BMIDS692 DRC 03/02/04 */ %>
  /* START: CODE COMMENTED BY Joyce Joseph AQR-MAR59	
  if (scScreenFunctions.GetContextParameter(window,"idFreezeOveride_Rev_and_Decis")=="1") 
  {
    frmScreen.optBMDeclarationYes.disabled=false;
    frmScreen.optBMDeclarationNo.disabled=false;
    
  }
  END: CODE COMMENTED BY Joyce Joseph AQR-MAR59 */
}
</SCRIPT>
