<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisRA = new ActiveXObject("omRA.RiskAssessmentBO");
	Response.Write(thisRA.UpdateRAUnderwriterDecision(xmlIn.xml));
%>
