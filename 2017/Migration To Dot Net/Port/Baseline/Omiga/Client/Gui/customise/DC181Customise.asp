<%
/*
Workfile:      dc181Customise.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Cosmetic customisation file for dc181
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DJP		30/01/02	First version, SYS2564(parent) SYS3806(child) Client specific cosmetic customisation
*/
%>

<SCRIPT LANGUAGE="JScript">

function Customise() 
{
	document.all("idPayrollNumber").innerHTML = "Payroll Number";
	document.all("idP60Seen").innerHTML = "P60 Seen?";
}

</SCRIPT>
