<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 09 Oct 2002 12:10:12 */ %>
<%	/* BMIDS00764 GHun 29/10/2002 - removed SYS4372 changes as SYS5242 conflicts with them */
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisRA = new ActiveXObject("omRA.RiskAssessmentBO");
	Response.Write(thisRA.GetRiskAssessment(xmlIn.xml));
%>
