<%@  Language=JScript %>
<%
/*
Workfile:      LogonFrameset.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Frameset for Omiga 4 logon screens
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Reason	Description
PSC		13/10/05	MAR992	Created
PE		24/04/06	MAR992	Removed password from XML.
PE		28/04/06	MAR992	Return XML even if login failed.
PE		04/05/2006	MAR992	When launched all buttons are disabled. (Context variables not set)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<HTML>
	<TITLE>System Logon</TITLE> 
	<!-- #include FILE="scServerXMLFunctions.asp" -->
	<%

	var sRequest = Request.Form(1);	
	var xmlDoc = new XMLObject();
	xmlDoc.LoadXML(sRequest);	
	
	xmlDoc.SelectTag(null, "REQUEST");
	spassword = xmlDoc.GetAttribute("PASSWORD");
	suserid = xmlDoc.GetAttribute("USERID");
	sunitid = GetUnitId(suserid);
	//xmlDoc.SetAttribute("UNITID", sunitid);
	xmlDoc.RemoveAttribute("PASSWORD");
	xmlDoc.RemoveAttribute("UNITID");
	xmlDoc.RemoveAttribute("MACHINEID");
	xmlDoc.RemoveAttribute("CHANNELID");

	var XMLRequest = new XMLObject();
	var XMLResponse = new XMLObject();
	
	XMLRequest.CreateActiveTag("REQUEST");
	XMLRequest.SetAttribute("USERID", suserid);
		
	XMLRequest.CreateActiveTag("OMIGAUSER");
	XMLRequest.CreateTag("USERID", suserid);
	XMLRequest.CreateTag("PASSWORDVALUE", spassword);
	XMLRequest.CreateTag("UNITID", sunitid);
	XMLRequest.CreateTag("AUDITRECORDTYPE", "1");

	var thisObject = new ActiveXObject("omOrg.OrganisationBO");
	XMLResponse.LoadXML(thisObject.ValidateUserLogon(XMLRequest.XMLDocument.xml));

	XMLResponse.SelectTag(null, "RESPONSE");
									
	if (XMLResponse.GetAttribute("TYPE") == "SUCCESS")
	{	
		xmlDoc.SetAttribute("CHANGEPASSWORDINDICATOR", XMLResponse.GetTagBoolean("CHANGEPASSWORDINDICATOR"));
		xmlDoc.SetAttribute("ACCESSAUDITATTEMPT", xmlDoc.GetTagBoolean("ACCESSAUDITATTEMPT"));
		xmlDoc.SetAttribute("AUTHENTICATIONID", GetAccessID(suserid));						
		xmlDoc.SetAttribute("USERNAME", XMLResponse.GetTagText("USERNAME"));
		xmlDoc.SetAttribute("ACCESSTYPE", XMLResponse.GetTagText("ACCESSTYPE"));
		xmlDoc.SetAttribute("COMPETENCYTYPE", XMLResponse.GetTagText("COMPETENCYTYPE"));
		xmlDoc.SetAttribute("CREDITCHECKACCESS", XMLResponse.GetTagText("CREDITCHECKACCESS"));
		xmlDoc.SetAttribute("ROLE", XMLResponse.GetTagText("ROLE"));
		
		sRequest = xmlDoc.XMLDocument.xml;
	}
	else
	{
		sRequest = xmlDoc.XMLDocument.xml;
	}			
	
	function GetUnitId(sUserId)
	{
		var XML = new XMLObject();
		var sUnitId = "";

		XML.CreateActiveTag("REQUEST");
		XML.SetAttribute("USERID", sUserId);
		XML.CreateActiveTag("SEARCH");
		XML.CreateTag("USERID", sUserId);
		/* PSC 09/01/2006 MAR975 */
		XML.CreateTag("ALLOWOMIGALOGON", "0");
		
		var thisObject = new ActiveXObject("omOrg.OrganisationBO");
		XML.LoadXML(thisObject.FindCurrentUnitList(XML.XMLDocument.xml));
		
		XML.SelectTag(null, "RESPONSE");
			
		if (XML.GetAttribute("TYPE") == "SUCCESS")
		{
			/* PSC 09/01/2006 MAR975 */
			XML.SelectSingleNode("DEPARTMENTLIST/DEPARTMENT/UNITLIST/UNIT[ALLOWOMIGALOGON != '1']");
			
			if (XML.ActiveTag != null)
				sUnitId = XML.GetTagText("UNITID");
		}			
					
		return sUnitId;
	}
	
	function GetAccessID(sUserID)			
	{
		var XMLRequest = new XMLObject();
		var XMLResponse = new XMLObject();
		
		XMLRequest.CreateActiveTag("REQUEST");
		XMLRequest.SetAttribute("CRUD_OP", "READ");
		XMLRequest.SetAttribute("SCHEMANAME", "ACCESSAUDITCHECK");
	
		XMLRequest.CreateActiveTag("ACCESSAUDITCHECK");	
		XMLRequest.SetAttribute("p_UserID", sUserID);
		
		var thisObject = new ActiveXObject("omCRUD.omCRUDBO");
		XMLResponse.LoadXML(thisObject.OmRequest(XMLRequest.XMLDocument.xml));		
		XMLResponse.SelectTag(null, "RESPONSE");
		
		if (XMLResponse.GetAttribute("TYPE") == "SUCCESS")
		{
			XMLResponse.SelectTag(null, "ACCESSAUDIT");
			return 	XMLResponse.GetAttribute("ACCESSAUDITGUID");
		}
		else
		{
			return "";
		}
	}
	
%>

	<xml id="xmlLaunchData">
		<%=sRequest%>
	</xml>
	<frameset cols="0,*" rows="*" frameborder="YES" bordercolor="#000066" border="5" framespacing="0">
		<frame src="OmigaLogon.asp" id="fraDetails" name="fraDetails" bordercolor="#000066" frameborder="YES"
			scrolling="yes" marginwidth="0" noresize />
		<frame src="sc015.asp" id="fraMain" name="fraMain" bordercolor="#000066" frameborder="YES"
			scrolling="yes" />
	</frameset>
</HTML>
