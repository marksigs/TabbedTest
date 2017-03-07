<SCRIPT LANGUAGE="JScript">
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specifc History:
Prog	Date	   AQR. NO		Description
JJ		30/09/2005 MAR58		cboSubPurposeOfLoan textbox and label removed.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
db Specifc History:
Prog	Date	   AQR. NO		Description
DRC		27/06/2006 EP430       MAR58 reversed	cboSubPurposeOfLoan textbox and label reinstated
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specifc History:
Prog	Date	   AQR. NO		Description
SW		21/06/2006 EP771		cboRepaymentVehicle set to required
MAH		04/01/2007 EP2_444		txtRepaymentVehicleMonthlyCost set as 0-9 with a decimal point 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
function SetMasks() 
{
	frmScreen.cboRepaymentType.setAttribute("required", "y");
	<% //SW 21/06/2006 EP771 %>
	frmScreen.cboRepaymentVehicle.setAttribute("required", "y");
	frmScreen.cboPurposeOfLoan.setAttribute("required", "y");
	frmScreen.txtLoanAmount.setAttribute("required", "y");
	frmScreen.txtLoanAmount.setAttribute("filter", "[0-9]");
	frmScreen.txtTermYears.setAttribute("required", "y");
	frmScreen.txtTermYears.setAttribute("filter", "[0-9]");
	frmScreen.txtTermMonths.setAttribute("integer", "true");
	frmScreen.txtTermMonths.setAttribute("min", "0");
	frmScreen.txtTermMonths.setAttribute("max", "11");
	//frmScreen.txtTermMonths.setAttribute("required", "y");
	frmScreen.txtTermMonths.setAttribute("filter", "[0-9]");
	
	frmScreen.txtRepaymentVehicleMonthlyCost.setAttribute("filter", "[0-9.]");<%/*EP2_444*/%>
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
	 if (scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idFreezeOveride_Shortfall")=="1") 
		{//temporarily overide the main context freezedata to disable read only status for the whole of this screen  ;
			m_sReadOnly = "0";
			scScreenFunctions.SetCollectionState(divSearch, "W");
			scScreenFunctions.SetCollectionState(divResults, "W");
		}
<% /* Reverse MAR58	
	//START: (MAR58) New code added by Joyce Joseph on 30-Sep-2005
	//cboSubPurposeOfLoan textbox and label removed.
	var grandParent = null;
	var df = null;
	df = this.document.getElementById("SubPurposeOfLoan");
	df.innerText = "";	
	//END: (MAR58)	
 */ %>		
}
</SCRIPT>