<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      ap100Customise.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Cosmetic customisation file for ap100

History:
Prog	Date		Description
----	----		-----------
AT		25/04/02	SYS4359	- First version, Client specific cosmetic customisation
GD		25/06/02	BMIDS0077 - remove refs to ids that have been commented out
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<SCRIPT LANGUAGE="JScript">
function Customise() 
{
	//document.all("idGrossBasic1").innerHTML = "Gross Basic<BR>(Incl. gtd PRP)"; 
	//document.all("idGrossBasic2").innerHTML = "Gross Basic<BR>(Incl. gtd PRP)"; 
	//document.all("idOvertime1").innerHTML = "Regular Overtime,<BR>PRP,Bonus,Comn."; 
	//document.all("idOvertime2").innerHTML = "Regular Overtime,<BR>PRP,Bonus,Comn."; 
	document.all("idNINumber1").innerHTML = "N.I. Number";
	document.all("idNINumber2").innerHTML = "N.I. Number";
}
</SCRIPT>
