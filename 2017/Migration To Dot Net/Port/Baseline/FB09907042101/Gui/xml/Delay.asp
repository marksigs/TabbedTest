<%@ Language="JScript" CodePage=65001 %>
<%

function Sleep(nMilliseconds)
{
	var dtStart = new Date();
	var nEnd = dtStart.valueOf() + nMilliseconds;
	var bLoop = true;
	while (bLoop)
	{
		var dtNow = new Date();
		if (dtNow.valueOf() >= nEnd)
		{
			bLoop = false;
		}
	}
}

var xmlIn = new ActiveXObject("microsoft.xmldom");
xmlIn.async = false;
xmlIn.load(Request);

var strSleep = xmlIn.documentElement.getAttribute("SLEEP");
var nDelay = parseInt(strSleep);

Sleep(nDelay * 1000);

Response.Write("<RESPONSE SLEEP=\"" + nDelay + "\"/>");

%>
