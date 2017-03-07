<SCRIPT LANGUAGE="JScript">
function SetThirdPartyDetailsMasks()
{
	frmScreen.txtPostcode.setAttribute("filter", "[0-9a-zA-Z ]");
	frmScreen.txtPostcode.setAttribute("upper", "true");
	//JR - Omiplus24
	/*frmScreen.txtTelephoneNo.setAttribute("filter", "[0-9 ]");
	frmScreen.txtFaxNo.setAttribute("filter", "[0-9 ]");
	frmScreen.txtEMailAddress.setAttribute("filter", "[-A-Za-z0-9@._/ ]");*/
	//End
	if (m_ctrSortCode != null)
	{
		m_ctrSortCode.setAttribute("mask","99-99-99");
		m_ctrSortCode.setAttribute("required", "true");
		m_ctrSortCode.setAttribute("filter","[0-9- ]");
	}
	if (m_ctrOrganisationType != null)
		m_ctrOrganisationType.setAttribute("required", "true");

	//START: MAR36 - New code added by Joyce Joseph on 10-Aug-2005
	// Capitlise text fields in Screens		
	frmScreen.txtHouseName.style.textTransform = "capitalize";
	frmScreen.txtStreet.style.textTransform = "capitalize";
	frmScreen.txtDistrict.style.textTransform = "capitalize";
	frmScreen.txtTown.style.textTransform = "capitalize";
	frmScreen.txtCounty.style.textTransform = "capitalize";
	frmScreen.txtContactForename.style.textTransform = "capitalize";
	frmScreen.txtContactSurname.style.textTransform = "capitalize"; 		
	//END: MAR36	

}

<% /* Specify screen attributes here */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
function SetMasks() 
{
 
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
