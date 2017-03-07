<%@ LANGUAGE="JSCRIPT" %>
<html>
<comment>
Workfile:      cu010.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Completios Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
NB		30/01/2002	Created for SYS3949

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
</comment>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>

<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>

<% /* FORMS */ %>
<form id="frmToDC240" method="post" action="cm010.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark   validate ="onchange">
<div style="HEIGHT: 582px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 490px" class="msgGroup">
This is the returned completions data&nbsp; <TEXTAREA id=TEXTAREA1 name=TEXTAREA1 style="HEIGHT: 549px; WIDTH: 480px"></TEXTAREA>
</div> 

</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 600px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/cu010Attribs.asp" -->
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sXML = "";
var scScreenFunctions;

var ArchitectXML = null;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber  = "";
var m_blnReadOnly = false;
var m_sUserId = "";


function btnCancel.onclick()
{
	
	
	window.close();
}

function btnSubmit.onclick()
{
	
		
	window.close();
	
}



<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	

	scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	//FW030SetTitles("Completions","cu010",scScreenFunctions);

	RetrieveContextData();
	//SetAttributeOnTag(sTagName, sAttributeName, vAttributeValue)
	var XML = new scXMLFunctions.XMLObject();
		
	XML.CreateActiveTag("REQUEST");
	XML.SetAttribute("USERID", m_sUserId);
	XML.SetAttribute("COMBOLOOKUP", "NO");
	XML.CreateActiveTag("APPLICATION");
	XML.SetAttribute("_SCHEMA_","COMPLETIONS");
	XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	
	
	XML.RunASP(document,"GetCompletionsData.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
			
	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		if(ErrorReturn[0] == true)
		{
			
			frmScreen.TEXTAREA1.value = XML.XMLDocument.xml;
		}
	}			
	ErrorTypes = null;
	ErrorReturn = null;
	
	
		
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}


function RetrieveContextData()
{

	scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];

	var sArgArray = sArguments[4];

	m_sApplicationNumber = sArgArray[0];
	m_sUserId = sArgArray[1];
	

	
	
}


-->
</script>
</body>
</html>
<% /* OMIGA BUILD VERSION 028.02.04.17.00 */ %>




