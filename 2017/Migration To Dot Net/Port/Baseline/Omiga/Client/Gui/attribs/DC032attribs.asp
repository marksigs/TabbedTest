<SCRIPT LANGUAGE="JScript">
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DC032attribs.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Alias/Association Details Screen attributes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		24/11/99	Created
AD		30/01/00	Rework
IW		19/04/00	Added Work Extension
JR		04/10/01	Omiplus24, Added AreaCode and CountryCode
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
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
	frmScreen.txtTelNumber1.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtTime1.setAttribute("filter", "[-.,:0-9a-zA-Z ]");
	frmScreen.txtTelNumber2.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtTime2.setAttribute("filter", "[-.,:0-9a-zA-Z ]");
	frmScreen.txtTelNumber3.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtTime3.setAttribute("filter", "[-.,:0-9a-zA-Z ]");
	frmScreen.txtTelNumber4.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtTime4.setAttribute("filter", "[-.,:0-9a-zA-Z ]");
	frmScreen.txtExtensionNumber1.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtExtensionNumber2.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtExtensionNumber3.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtExtensionNumber4.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtCountryCode1.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtCountryCode2.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtCountryCode3.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtCountryCode4.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtAreaCode1.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtAreaCode2.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtAreaCode3.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtAreaCode4.setAttribute("filter", "[-0-9\(\) ]");
	
	//SD MAR209 - Start
	frmScreen.txtTelNumber1.setAttribute("filter", "[0-9]");
	frmScreen.txtTelNumber2.setAttribute("filter", "[0-9]");
	frmScreen.txtTelNumber3.setAttribute("filter", "[0-9]");
	frmScreen.txtTelNumber4.setAttribute("filter", "[0-9]");
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
	//SD MAR209 - End
	
	<% // MAR1331 - Omiga - Error message for an invalid email did not appear %>
	<% // Peter Edney %>	
	frmScreen.txtEMailAddress.setAttribute("filter", "[-A-Za-z0-9@._/]");
	frmScreen.txtEMailAddress.setAttribute("msg", "Please enter a valid email address.");
	frmScreen.txtEMailAddress.setAttribute("regexp", "^([a-zA-Z0-9_\\-\\.]+)@((\\[[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.)|(([a-zA-Z0-9\\-]+\\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\\]?)$");
	
}
</SCRIPT>