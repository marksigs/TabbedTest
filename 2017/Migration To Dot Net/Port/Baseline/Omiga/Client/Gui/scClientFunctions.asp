<%@  Language=JScript %>
<%/*

Created by automated update TW 09 Oct 2002 SYS5115

Workfile:      scClientFunctions.htm
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Generic client functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog		Date		Description
TW		25/07/2002	Created to contain client's own common functions 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

List the external function calls and their parameters here:
			
*/
%>

<%/* If XML functions are required, un-comment this OBJECT


<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>


*/
%>

<HTML id=ClientFunctions>
<script language="JavaScript">

public_description = new CreateClientScreenFunctions;

function CreateClientScreenFunctions()
{
	this.ClientScreenFunctionsObject = ClientScreenFunctionsObject;
}

function ClientScreenFunctionsObject()
{
<% /* Set up references for each ClientFunction.
 	Please note that the following is only an example and may be removed.
	If it is removed, the function ClientCheckAge(cAge) below must also be removed.
*/
%>
	this.ClientCheckAge = ClientCheckAge;
}

<% /* ClientFunctions here 
 	Please note that the following is only an example and may be removed
	If it is removed, the corresponding line in ClientScreenFunctionsObject
	above must also be removed.
*/
%>


function ClientCheckAge(cAge)
{
// Check Age is within defined limits. Return 1 if OK, else return 0
    	if(cAge==null)
	{
		return 0;
	}
	else
	{
	//START: New Code added 06-Sep-2005 to get Global Parameter Values for Minimum & Maximum Age.
		var m_sMinimumAge = "";
		var m_sMaximumAge = "";

		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_sMinimumAge = XML.GetGlobalParameterAmount(document, "MinimumAge");
		m_sMaximumAge = XML.GetGlobalParameterAmount(document, "MaximumAge");
		XML = null;
	
		if(cAge < m_sMinimumAge || cAge > m_sMaximumAge)
		{
			return 0;
		}
		else
		{
			return 1;
		}
	//END: New Code added on 06-Sep-2005 to get Global Parameter Values for Minimum & Maximum Age.	
	}
}

</script>
