<SCRIPT LANGUAGE="JScript">
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specifc History:
Prog	Date		AQR. NO		Description
JJ		30/09/2005	MAR58		cboSubPurposeOfLoan textbox and label removed.
INR		15/01/2007	EP2_697		reinstate cboSubPurposeOfLoan
PE		20/01/2007	EP2_922		Previous version of this file never made it to the build.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
function SetMasks() 
{
	frmScreen.cboRepaymentType.setAttribute("required", "y");
	frmScreen.cboPurposeOfLoan.setAttribute("required", "y");
	frmScreen.txtLoanAmount.setAttribute("required", "y");
	frmScreen.txtLoanAmount.setAttribute("filter", "[0-9]");
	frmScreen.txtTermYears.setAttribute("required", "y");
	frmScreen.txtTermYears.setAttribute("filter", "[0-9]");
	frmScreen.txtTermMonths.setAttribute("integer", "true");
	frmScreen.txtTermMonths.setAttribute("min", "0");
	frmScreen.txtTermMonths.setAttribute("max", "11");
	frmScreen.txtTermMonths.setAttribute("filter", "[0-9]");
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
}
</SCRIPT>