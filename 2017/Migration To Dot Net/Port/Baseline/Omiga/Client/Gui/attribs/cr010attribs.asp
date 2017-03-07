<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
<%	/*	DOWNLOADXTABLE
		This variable specifies the page size of data to be downloaded during a search
		and when scrolling the table.
		The size required is expressed as a multiple of the table size on the screen.
		For example, the table on cr010 has 10 rows.
			Setting DOWNLOADXTABLE = 1 will download 10 records at a time.
			Setting DOWNLOADXTABLE = 10 will download 100 records at a time.
		If this variable is not set in this file it defaults to 100 records
	*/
%>	DOWNLOADXTABLE = 10;

	frmScreen.txtSurname.setAttribute("wildcard", "true");
	frmScreen.txtForenames.setAttribute("wildcard", "true");
	//START: MAR72 - New code added by Joyce Joseph on 20-Sep-2005
	//Included % character for wildcard search
	// MAR369 PJO 09/11/2005 - Allow - and '
	frmScreen.txtSurname.setAttribute("filter", "[a-zA-Z%'-]");
	frmScreen.txtForenames.setAttribute("filter", "[a-zA-Z%'-]");	
	//END: MAR72
<%	//this line has been commented out because the spec made no reference
	//to a specified character set apart from the standard
	//frmScreen.txtForenames.style.setAttribute ("charset", "0123456789 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ");
%>
	frmScreen.txtDateOfBirth.setAttribute("date", "DD/MM/YYYY");
//	frmScreen.txtDateOfBirth.setAttribute("filter", "[0-9/]");
	frmScreen.txtPostcode.setAttribute("filter", "[0-9a-zA-Z ]");
	frmScreen.txtPostcode.setAttribute("wildcard", "true");
	frmScreen.txtPostcode.setAttribute("upper","true");
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
	