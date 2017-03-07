<%
/*
Workfile:      dc010Customise.asp
Copyright:     Copyright © 2001 Marlborough Stirling

Description:   Cosmetic customisation file for dc014
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
INR		15/06/04	Client specific cosmetic customisation
*/
%>
<SCRIPT LANGUAGE="JScript">
function Customise() 
{
	document.all("TPDDeclarationLabel").innerHTML = "Has the Third Party Declaration been seen by the customer";
	document.all("TPDDeclarationText").innerHTML = "Declaration Text to be inserted here. <br> Declaration Text<br> Some Declaration Text<br> More Declaration Text<br>";
}

</SCRIPT>
