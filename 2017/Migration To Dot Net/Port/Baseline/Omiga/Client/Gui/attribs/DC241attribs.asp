<SCRIPT LANGUAGE="JScript">
<%
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description 
SD		19/10/2005	MAR209		Allow only alphabets in name fields and numeric characters in phone fields
PE		06/03/2006	MAR1331		Omiga - Error message for an invalid email did not appear
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
%>
function SetMasks()
{
	frmScreen.txtCountryCode1.setAttribute("filter", "[0-9 ]");
	frmScreen.txtCountryCode2.setAttribute("filter", "[0-9 ]");
	frmScreen.txtAreaCode1.setAttribute("filter", "[0-9 ]");
	frmScreen.txtAreaCode2.setAttribute("filter", "[0-9 ]");
	frmScreen.txtTelephoneNumber1.setAttribute("filter", "[0-9 ]");
	frmScreen.txtTelephoneNumber2.setAttribute("filter", "[0-9 ]");
	frmScreen.txtWorkExtNo1.setAttribute("filter", "[0-9 ]");
	frmScreen.txtWorkExtNo2.setAttribute("filter", "[0-9 ]");
	// ASu - BMIDS00106 Start
	frmScreen.txtCountryCode1.setAttribute("phone", "[0-9 ]");
	frmScreen.txtCountryCode2.setAttribute("phone", "[0-9 ]");
	frmScreen.txtAreaCode1.setAttribute("phone", "[0-9 ]");
	frmScreen.txtAreaCode2.setAttribute("phone", "[0-9 ]");
	frmScreen.txtTelephoneNumber1.setAttribute("phone", "[0-9 ]");
	frmScreen.txtTelephoneNumber2.setAttribute("phone", "[0-9 ]");
	frmScreen.txtWorkExtNo1.setAttribute("phone", "[0-9 ]");
	frmScreen.txtWorkExtNo2.setAttribute("phone", "[0-9 ]");
	// ASu - End
	
	// SD - MAR209 Start
	frmScreen.txtCountryCode1.setAttribute("filter", "[0-9]");
	frmScreen.txtCountryCode2.setAttribute("filter", "[0-9]");
	frmScreen.txtAreaCode1.setAttribute("filter", "[0-9]");
	frmScreen.txtAreaCode2.setAttribute("filter", "[0-9]");
	frmScreen.txtTelephoneNumber1.setAttribute("filter", "[0-9]");
	frmScreen.txtTelephoneNumber2.setAttribute("filter", "[0-9]");
	frmScreen.txtWorkExtNo1.setAttribute("filter", "[0-9]");
	frmScreen.txtWorkExtNo2.setAttribute("filter", "[0-9]");
	// SD -MAR209 End
	
	<% // MAR1331 - Omiga - Error message for an invalid email did not appear %>
	<% // Peter Edney %>	
	frmScreen.txtEmail.setAttribute("filter", "[-A-Za-z0-9@._/]");
	frmScreen.txtEmail.setAttribute("msg", "Please enter a valid email address.");
	frmScreen.txtEmail.setAttribute("regexp", "^([a-zA-Z0-9_\\-\\.]+)@((\\[[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.)|(([a-zA-Z0-9\\-]+\\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\\]?)$");
	
}
</SCRIPT>
