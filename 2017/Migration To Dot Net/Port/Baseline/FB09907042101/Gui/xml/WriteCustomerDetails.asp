<%@ Language=JScript %>
<% 
	Response.Buffer = true; 
%>

<html>
<head>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>

<body>
	<form id=idXmlForm method="post" action="WriteCustomerDetails.asp">
		<input id=idXmlReq name="XmlReq" type="hidden"></input>
	</form>

<%
	if(Request.Form.Count != 0)
	{
		var sRequest = Request.Form("XmlReq");
		var sResponse;
		var thisCustreg = new ActiveXObject("omCR.CustomerBO");

		var xmlRequest = new ActiveXObject("microsoft.xmldom");
		xmlRequest.async = false;
		xmlRequest.loadXML(sRequest);
		
		var sRequestType = xmlRequest.firstChild.firstChild.nodeName;

		if(sRequestType == "CREATE")		
		{
			sResponse = thisCustreg.CreateCustomerDetails(sRequest);
		}
		else
		{
			if(sRequestType == "UPDATE")
			{
				sResponse = thisCustreg.UpdateCustomerDetails(sRequest);
			}
		}
		
		Response.Write("<comment id=xmlResponse>\r\n");
		Response.BinaryWrite(sResponse);
		Response.Write("</comment>\r\n");
		Response.Write("<SCRIPT LANGUAGE=javascript>\r\n");
		Response.Write("var thisForm = parent.document.forms(\"frmXMLState\");\r\n");
		Response.Write("thisForm.txtXMLState.value = \"Response\";\r\n");
		Response.Write("thisForm.btnResponse.click();\r\n");
		Response.Write("</SCRIPT>\r\n");
		
		thisCustreg = null;
		xmlRequest = null;
	}
%>

</body>
</html>
