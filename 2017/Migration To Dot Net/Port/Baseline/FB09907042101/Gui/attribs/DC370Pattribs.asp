<SCRIPT LANGUAGE="JScript">

<% /* Specify screen attributes here */ %>
<% /* MAR639: Maha T - 30/11/2005 -  Amended */ %>
function SetMasks() 
{
	frmScreen.txtMaximumBorrowing.style.color = "orange" ;
	frmScreen.txtMaximumBorrowing.style.fontWeight ="bolder";
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
//frmScreen.txtIndemnityPremium.setAttribute("filter", "[0-9]");
//MAR507 Add text
	var sText = "";
	sText += "                                                                            ";
	sText += "Decline:\r\n\r\n";
	sText += "                                                                       ";
	sText += "If Bureau Decline:\r\n";
	sText += "We are unable to proceed with your application for an ING Direct mortgage ";
	sText += "as it does not meet our lending criteria.\r\n";
	sText += "You may call Experian at 0870 241 6212 to obtain a copy of your credit file. ";
	sText += "Alternatively, you can appeal this decision if you feel there is information "
	sText += "you have not been able to make us aware of.\r\n\r\n";
	sText += "                                                                      ";
	sText += "If Policy Decline:\r\n";
	sText += "We are unable to proceed with your application for an ING Direct mortgage ";
	sText += "as it does not meet our lending criteria.\r\n\r\n";
	sText += "Inform customer of specific policy decline reason";

	frmScreen.txtAffordability.value = sText;
	frmScreen.txtAffordability.style.color = "Red";

frmScreen.txtMaximumBorrowing.style.color = "Orange" ;
frmScreen.txtMaximumBorrowing.style.fontWeight ="bolder";

}

</SCRIPT>