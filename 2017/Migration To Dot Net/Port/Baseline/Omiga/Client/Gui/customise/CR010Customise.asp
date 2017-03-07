<%
/*
Workfile:      cr010Customise.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Cosmetic customisation file for CR010
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DJP		25/09/01	SYS2564/SYS2735 (child) First version, Client specific cosmetic customisation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS History:

Prog	Date		Description
PSC		05/10/2005	MAR57 Change Forename (s) to First Name
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<SCRIPT LANGUAGE="JScript">
var m_sPostCode = "Postcode";
var m_sForename = "First Name";
var m_sSurname = "Surname";
function Customise() 
{
	document.all("idSurname").innerHTML = m_sSurname;
	document.all("idForename").innerHTML = m_sForename;
	document.all("idPostCode").innerHTML = m_sPostCode;
}
</SCRIPT>