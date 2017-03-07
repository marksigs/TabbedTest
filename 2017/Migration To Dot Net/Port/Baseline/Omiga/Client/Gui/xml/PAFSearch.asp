<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 09 Oct 2002 12:10:12 */ %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);
		
	/* 
	Example of Server side scripting to apply a different criteria to the PAF search
		
	var bCriteriaOK = true;
	
	if (xmlIn.selectSingleNode("//POSTCODE") == null || xmlIn.selectSingleNode("//POSTCODE").text.length == 0)
	{
		if((xmlIn.selectSingleNode("//STREET") == null || xmlIn.selectSingleNode("//STREET").text.length == 0) ||
			(xmlIn.selectSingleNode("//TOWN") == null || xmlIn.selectSingleNode("//TOWN").text.length == 0))
		{
			xmlOut = "<RESPONSE TYPE=\"APPERR\">";
			xmlOut += "<ERROR>"
			xmlOut += "<NUMBER>999</NUMBER>"
			xmlOut += "<SOURCE>PAFSEarch.asp</SOURCE>"
			xmlOut += "<DESCRIPTION>Incomplete PAF search criteria, enter [Post Code] or [Street and Town]</DESCRIPTION>"
			xmlOut += "</ERROR>"
			xmlOut += "</RESPONSE>";
				
			bCriteriaOK = false;
		}
	}		
	*/
		
	var thisPAFBO = new ActiveXObject("omPAF.PAFBO");
	Response.Write(thisPAFBO.FindPAFAddress(xmlIn.xml));
%>
