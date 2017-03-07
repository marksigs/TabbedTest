<SCRIPT LANGUAGE="JScript">

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history
PE		06/03/2006	MAR1331	Omiga - Error message for an invalid email did not appear
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

<% /* Specify screen attributes here */ %>
function SetMasks() 
{
	frmScreen.txtCallDate.setAttribute("date", "DD/MM/YYYY")
	<%/* START: MAR203  Mahasen T */ %>
	frmScreen.txtCallTime.setAttribute("mask", "99:99");;
	<%/* END: MAR203 */ %>
	frmScreen.txtCallTime.setAttribute("filter", "[0-9\:]");
	frmScreen.txtPassword.setAttribute("filter", "[0-9a-zA-Z]");
	// JD MAR1287 frmScreen.txtPassword.setAttribute("required", "true");
	frmScreen.txtPasswordConfirm.setAttribute("filter", "[0-9a-zA-Z]");
	// JD MAR1287 frmScreen.txtPasswordConfirm.setAttribute("required", "true");
	
	frmScreen.txtTelNumber1.setAttribute("filter", "[0-9]");
	frmScreen.txtTelNumber2.setAttribute("filter", "[0-9]");
	frmScreen.txtTelNumber3.setAttribute("filter", "[0-9]");
	frmScreen.txtTelNumber4.setAttribute("filter", "[0-9]");
	frmScreen.txtTime1.setAttribute("filter", "[-.,:0-9a-zA-Z ]");
	frmScreen.txtTime2.setAttribute("filter", "[-.,:0-9a-zA-Z ]");
	frmScreen.txtTime3.setAttribute("filter", "[-.,:0-9a-zA-Z ]");
	frmScreen.txtTime4.setAttribute("filter", "[-.,:0-9a-zA-Z ]");
	frmScreen.txtExtensionNumber1.setAttribute("filter", "[0-9]");
	frmScreen.txtExtensionNumber2.setAttribute("filter", "[0-9]");
	frmScreen.txtExtensionNumber3.setAttribute("filter", "[0-9]");
	frmScreen.txtExtensionNumber4.setAttribute("filter", "[0-9]");
	frmScreen.txtCountryCode1.setAttribute("filter", "[0-9]");
	frmScreen.txtCountryCode2.setAttribute("filter", "[0-9]");
	frmScreen.txtCountryCode3.setAttribute("filter", "[0-9]");
	frmScreen.txtCountryCode4.setAttribute("filter", "[0-9]");
	frmScreen.txtAreaCode1.setAttribute("filter", "[0-9]");
	frmScreen.txtAreaCode2.setAttribute("filter", "[0-9]");
	frmScreen.txtAreaCode3.setAttribute("filter", "[0-9]");
	frmScreen.txtAreaCode4.setAttribute("filter", "[0-9]");
	<%/* START: MAR323*/ %>
	frmScreen.txtCountryCode1.setAttribute("phone", "[0-9]");
	frmScreen.txtCountryCode2.setAttribute("phone", "[0-9]");
	frmScreen.txtCountryCode3.setAttribute("phone", "[0-9]");
	frmScreen.txtCountryCode4.setAttribute("phone", "[0-9]");
	frmScreen.txtAreaCode1.setAttribute("phone", "[0-9]");
	frmScreen.txtAreaCode2.setAttribute("phone", "[0-9]");
	frmScreen.txtAreaCode3.setAttribute("phone", "[0-9]");
	frmScreen.txtAreaCode4.setAttribute("phone", "[0-9]");
	frmScreen.txtTelNumber1.setAttribute("phone", "[0-9]");
	frmScreen.txtTelNumber2.setAttribute("phone", "[0-9]");
	frmScreen.txtTelNumber3.setAttribute("phone", "[0-9]");
	frmScreen.txtTelNumber4.setAttribute("phone", "[0-9]");
	frmScreen.txtExtensionNumber1.setAttribute("phone", "[0-9]");
	frmScreen.txtExtensionNumber2.setAttribute("phone", "[0-9]");
	frmScreen.txtExtensionNumber3.setAttribute("phone", "[0-9]");
	frmScreen.txtExtensionNumber4.setAttribute("phone", "[0-9]");
	<%/* END: MAR323*/ %>

	<% // MAR1331 - Omiga - Error message for an invalid email did not appear %>
	<% // Peter Edney %>	
	frmScreen.txtContactEMailAddress.setAttribute("filter", "[-A-Za-z0-9@._/]");
	frmScreen.txtContactEMailAddress.setAttribute("msg", "Please enter a valid email address.");
	frmScreen.txtContactEMailAddress.setAttribute("regexp", "^([a-zA-Z0-9_\\-\\.]+)@((\\[[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.)|(([a-zA-Z0-9\\-]+\\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\\]?)$");
	
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
frmScreen.txtAgreeNextSteps.value="Agree next steps configurable text";
}

</SCRIPT>