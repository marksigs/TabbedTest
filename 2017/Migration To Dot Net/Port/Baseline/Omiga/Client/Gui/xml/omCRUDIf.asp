<%@ Language="JScript" CodePage=65001 %>
<%
	try
	{
		var xmlIn = new ActiveXObject("Msxml2.FreeThreadedDOMDocument.4.0");
		xmlIn.async = false;
		xmlIn.load(Request);
		
	//	ik_debug
	//	xmlIn.save("c:\\omiga4Trace\\omCrudIf_Request.xml");	

		var o = new ActiveXObject("omCRUD.omCRUDBO");
		Response.Write(o.omrequest(xmlIn.xml));
	}
	catch(e)
	{
		Response.Write(
			'<RESPONSE TYPE="SYSERR">' +
				'<ERROR>' +
					'<NUMBER>' + e.number + '</NUMBER>' + 
					'<SOURCE>omCRUDIf.asp</SOURCE>' + 
					'<DESCRIPTION>' + e.description + '</DESCRIPTION>' +
					'<VERSION/>' +
				'</ERROR>' + 
			'</RESPONSE>');
	}
	finally
	{
	}
%>
