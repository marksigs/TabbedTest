<%
/* 
Workfile:      fw020Customise.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Cosmetic customisation file for fw020
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MC		25/09/01	SYS2564/SYS2767 (child) First version, Client specific cosmetic customisation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
MV		25/05/2002	BMIDS00013 - BM044 - Removed Dependants (DC040) from AIP menu and from Screen Routing 
SR		19/06/2004  BMIDS772    Remove CCJ History
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		Description
PE		14/12/2005	MAR868 - If the global parameter "NewPropertySummary" is 0, the screen flow is changed.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

*/
%>
<SCRIPT LANGUAGE="JScript">
function Customise() 
{
	// AIP
	<% /* document.all("idNavLink6").innerHTML = "<A href=#ToAIP onclick=Navigate('AIP','DC040.ASP') tabIndex=-1>Dependants</A>"; */ %>
	<% /* SR 19/06/2004 : BMIDS772 - remove CCj History from menu 
	document.all("idNavLink7").innerHTML = "<A href=#ToAIP onclick=Navigate('AIP','DC150.ASP') tabIndex=-1>CCJ History</A>";
	*/ %>
	// Mortgage Application
	
	var XML = new scXMLFunctions.XMLObject();	
	var bNewPropSummary = XML.GetGlobalParameterBoolean(document,"NewPropertySummary");	
	var bThirdSummary = XML.GetGlobalParameterBoolean(document,"ThirdPartySummary");	
	if (bNewPropSummary){
		document.all("idNavLink220").style.display="none";
		document.all("idNavLink225").style.display="none";
		document.all("idNavLink201").innerHTML = "<a href=#ToMA onclick=Navigate('MA','DC201.ASP') tabIndex=-1>Property Details</a>";		
		document.all("idNavLink230").style.display="none";
		document.all("idNavLink295").style.display="none";
		document.all("idNavLink300").style.display="none";				
	}else{
		document.all("idNavLink201").style.display="none";
		document.all("idNavLink200").innerHTML = "<A href=#ToMA onclick=Navigate('MA','DC200.ASP') tabIndex=-1>New Loan & Property</A>";				
		document.all("idNavLink220").innerHTML = "<A href=#ToMA onclick=Navigate('MA','DC220.ASP') tabIndex=-1>Property Description</A>";
		document.all("idNavLink225").innerHTML = "<A href=#ToMA onclick=Navigate('MA','DC225.ASP') tabIndex=-1>Property Details</A>";
		document.all("idNavLink225").style.width = "130px";		
		document.all("idNavLink230").innerHTML = "<A href=#ToMA onclick=Navigate('MA','DC230.ASP') tabIndex=-1>Other Residents</A>";
		document.all("idNavLink230").style.width = "140px";
		document.all("idNavLink295").innerHTML = "<A href=#ToMA onclick=Navigate('MA','DC295.ASP') tabIndex=-1>Other Insurance Co</A>";
		document.all("idNavLink300").innerHTML = "<A href=#ToMA onclick=Navigate('MA','DC300.ASP') tabIndex=-1>B &amp; C Declaration</A>";
	}		
	if(bThirdSummary){
		document.all("idNavLink240").innerHTML = "<A href=#ToMA onclick=Navigate('MA','DC240.ASP') tabIndex=-1>Third Party Summary</A>";
		document.all("idNavLink280").style.display="none";
		document.all("idNavLink270").style.display="none";		
	}else{
		document.all("idNavLink240").style.display="none";
		document.all("idNavLink280").innerHTML = "<a href=#ToMA onclick=Navigate('MA','DC280.ASP') tabIndex=-1>Legal Representative</a>";
		document.all("idNavLink270").innerHTML = "<a href=#ToMA onclick=Navigate('MA','DC270.ASP') tabIndex=-1>Bank Details</a>";
	}
}

</SCRIPT>
