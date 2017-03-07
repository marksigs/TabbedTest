<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtCardNumber.setAttribute("filter", "[0-9]");
	frmScreen.txtIssueNumber.setAttribute("filter", "[0-9]");
	frmScreen.txtExpiryDate.setAttribute("mask", "99/99");
	// Note that following fields are shown/hidden in the client:
	frmScreen.txtNameOfCardHolders.setAttribute("required", "true");
	frmScreen.cboCardType.setAttribute("required", "true");
	frmScreen.txtCardNumber.setAttribute("required", "true");

<% /* SR 16/05/00 - SYS0708 : IssueNumber is not a mandatory field
							  CardHolder's name should be displayed in Upper case			
	frmScreen.txtIssueNumber.setAttribute("required", "true");
	*/
%>
	frmScreen.txtNameOfCardHolders.setAttribute("upper", "true");
<% 
/***  FIX ME : SR 16/05/00 - SYS0708-
	  this line is commented to just allow saving date for time being. To be uncommented later.
	  ASu - BMIDS00420 - Uncomment mask for complete validation */ %>
	frmScreen.txtExpiryDate.setAttribute("required", "true");
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