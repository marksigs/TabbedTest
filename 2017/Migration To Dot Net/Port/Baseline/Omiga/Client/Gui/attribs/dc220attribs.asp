<% /*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom history

Who		When		What	Why
PB		23/05/2006	EP570	'Date built' field should not be mandatory
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

*/ %><SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.cboTenureOfProperty.setAttribute("required", "true");
	frmScreen.cboPropertyType.setAttribute("required", "true");
	frmScreen.txtNumStoreys.setAttribute("required", "true");
	<% /* PB 23/05/2006 EP570
	frmScreen.txtYearBuilt.setAttribute("required", "true"); */ %>
	frmScreen.txtYearBuilt.setAttribute("mask", "9999");
	frmScreen.txtGuaranteeExpiryDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.cboBuildingConstruction.setAttribute("required", "true");
	frmScreen.cboRoofConstruction.setAttribute("required", "true");
	frmScreen.txtNumStoreys.setAttribute("filter", "[0-9]");
	frmScreen.txtYearBuilt.setAttribute("filter", "[0-9]");
	frmScreen.txtNumGarages.setAttribute("filter", "[0-9]");
	frmScreen.txtNumOutbuildings.setAttribute("filter", "[0-9]");
	//SYS0925 Want a different message so now dealt with in DC220.asp
	//frmScreen.txtResidentialPctge.setAttribute("integer", "true");
	//frmScreen.txtResidentialPctge.setAttribute("min", "0");
	//frmScreen.txtResidentialPctge.setAttribute("max", "100");
	frmScreen.txtResidentialPctge.setAttribute("filter", "[0-9]");
	frmScreen.txtFloorNo.setAttribute("filter", "[0-9]");
	//DB SYS4767 - MSMS Integration
	frmScreen.txtDateOfEntry.setAttribute("date", "DD/MM/YYYY");
	//DB End
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
}
</SCRIPT>