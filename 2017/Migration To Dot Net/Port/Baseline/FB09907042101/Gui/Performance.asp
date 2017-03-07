<%@ Language=JScript %>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body>
<!-- PERFORMANCE ENHANCEMENTS FOR THE GUI -->

<!-- Part 1: Making existing screens suitable for IE5 -->

<!-- Replace the line: -->
	<object data="scFormManager.htm" height="1" id="scFormMgr" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object> 
<!-- with: -->
	<script src="validation.js" language="JScript"></script>


<!-- In the scScreenFunctions object, change scScreenFunctions.htm to scScreenFunctions.asp -->
<!-- In the scXMLFunctions object, change scXMLFunctions.htm to scXMLFunctions.asp -->
<!-- In the scMathFunctions object, change scMathFunctions.htm to scMathFunctions.asp -->
<!-- In the scTableListScroll object, change scTableListScroll.htm to scTableListScroll.asp -->


<!-- Replace: -->
	<!-- Button field to keep tabbing within this screen -->
	<span style="VISIBILITY: HIDDEN">
		<input id="btnToLastField" type="button">
	</span>
<!-- with: -->
	<% /* Span to keep tabbing within this screen */ %>
	<span id="spnToLastField" tabindex="0"></span>


<!-- Change: -->
	<form id="frmScreen">
<!-- To: -->
	<form id="frmScreen" mark validate="onchange" year4>
<!-- N.B. Only include year4 if you have date fields -->


<!-- Replace: -->
	<!-- Button field to keep tabbing within this screen -->
	<span style="VISIBILITY: HIDDEN">
		<input id="btnToFirstField" type="button">
	</span>
<!-- with: -->
	<% /* Span to keep tabbing within this screen */ %>
	<span id="spnToFirstField" tabindex="0"></span>


<!-- Any table definitions must have &nbsp in each empty <TD> element e.g. -->
	<tr id="row01">		<td width="50%" class="TableTopLeft">&nbsp</td>		<td class="TableTopRight">&nbsp</td></tr>


<!-- Within the code: -->
<script language="JScript">
// Replace: 
			//Initialise the Form Manager - this is needed even if there is no soft coding for the screen in order to apply the standard character set
			var refFormOne = document.forms("frmScreen");
			scFormMgr.initialise(refFormOne);
// with:
			Validation_Init();
// N.B. If there is a Masks() function call ensure Validation_Init() is called after it


// Replace:
		function btnToFirstField.onfocus()
// with:
		function spnToFirstField.onfocus()


// Replace:
		function btnToLastField.onfocus()
// with:
		function spnToLastField.onfocus()


// Replace any occurance of:
		scFormMgr.isSubmitOK()
// with:
		frmScreen.onsubmit()
// (This also returns true or false)


// Replace any occurance of
		scFormMgr.flagChange(bFlag);
// with
		FlagChange(bFlag);
		

// Replace any occurance of
		scFormMgr.isChanged()
// with
		IsChanged()
// (This also returns true or false)


// The following functions now have these prototypes (all are members of scXMLFunctions):
		RunASP(thisDocument,sASPFile)

		RunASPWithTextInput(thisDocument,sASPFile,sRequest)

		bReturn = GetComboLists(thisDocument,sGroupList)

		bReturn = PopulateCombo(thisDocument,refFieldId,sListName,bAddSELECT)

		bReturn = PopulateComboFromXML(thisDocument,refFieldId,XMLSource,bAddSELECT)

		bReturn = IsInComboValidationList(thisDocument,sListName, sValueID, ValidationList)

		bReturn = GetGlobalParameterBoolean(thisDocument,sParameterName)

		sAmount = GetGlobalParameterAmount(thisDocument,sParameterName)
// All these functions now have a parameter which needs to be set to the document object
// GetComboLists also no longer takes UserID UnitType and UnitID parameters


// There are also new functions which now generate the <REQUEST> XML tag with a standard set of attributes
// They are:

		AttributeArray = CreateRequestAttributeArray(thisWindow)
		TagNew = CreateRequestTagFromArray(AttributeArray,sAction)
		TagNew = CreateRequestTag(thisWindow,sAction)
		
// Any code which generates a <REQUEST> tag should be replaced by these functions e.g. code like:
		ListXML.CreateActiveTag("REQUEST");
		ListXML.SetAttribute("USERID", m_sUserId);
		ListXML.SetAttribute("USERTYPE", m_sUserType);
		ListXML.SetAttribute("UNIT", m_sUnitId);

// If this code is on a main screen, simply replace with CreateRequestTag

// If this code is on a popup, CreateRequestAttributeArray is called on the main screen and the resulting
// array is passed into the popup.  The code on the popup is replaced by CreateRequestTagFromArray.
// The sAction parameter is the data action e.g. CREATE, UPDATE, DELETE etc.  It may also be set to null if
// an action does not need to be specified.  Remove the existing CREATE, UPDATE, DELETE etc tags.
// See the comments in scXMLFunctions for more details


// CHANGES TO THE XML ASP FILES
// The code in the xml files should be changed from:
<%
	var thisObject = new ActiveXObject("object.method");
	Response.ContentType = "text/xml";
	Response.Write(thisObject.Method(Request.QueryString("Request")));
%>
// to:
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisObject = new ActiveXObject("object.method");
	Response.Write(thisObject.Method(xmlIn.xml));
%>
// N.B. Any data access functionality which uses an IFRAME to call an XML asp file should be changed to
// use the RunASP style of call.


// CHANGES TO THE ATTRIBS FILES
// See AY for details


// Part 2 - Creating New Screens
// The modal_template and html_cookbook have been changed to match the new best practices.
// General rules for writing new code are:
// Only include scriptlets that you require
// Avoid spaces/tabs where possible, but not so much as to lose readability
// Make the JScript code as tight as possible, but again not so much as to lose readability
// Enclose all comments in server-side script.  Try to group comments as much as possible

// See the Customer Registration screens for examples
</script>
</body>
</html>
