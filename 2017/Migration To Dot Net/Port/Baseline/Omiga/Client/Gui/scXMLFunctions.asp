<%@  Language=JScript %>
<%/*
Workfile:      scXMLFunctions.htm
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Helper functions for XML parser.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
RF		03/11/99	Added filename parameter to WriteXMLToFile.
RF		09/11/99	Improved error handling in SetAttribute.
RF		03/12/99	Display error source in CheckResponse.
AY		28/01/00	Optimised
					Functions using document objects altered for IE5
					RunASP etc altered to use HTTP method
					Central functionality for REQUEST tag generation
AY		28/01/00	GetComboList and GetComboValidationList calls changed
AY		31/01/00	CreateRequestTag functions altered to have an action attribute
AY		04/02/00	Error in rewritten GetTagFloat and GetTagInt functions
AY		16/02/00	SYS0240 - HEAD and TITLE elements removed
APS		07/03/00	SYS0275 - Warning messages
IW		18/05/00	Added DebugXMLCheck() to RunAsp
APS		08/06/00	SYS0833 - Added LOANCOMPONENTSNOTFOUND error message
PSC		08/08/00    Display build number in CheckResponse
JLD		09/11/00	Added debug method ReadXMLFromFile
JLD		15/11/00	Altered REQUEST tag to include OPERATION attribute
GD		15/12/00	Added USERAUTHORITYLEVEL(idRole) attribute to REQUEST tag
SR		20/12/00	Added new function GetComboDescriptionForValidation
JLD		02/01/01	Added NOTASKAUTHORITY and NOSTAGEAUTHORITY error numbers.
MV		05/01/01	Enhanced functionality of PopulateComboFromXML to Support Element based or Attribute based
APS		09/01/01	SYS1808: Defensive coding for this tag as not all Popups have been changed
JLD		10/01/01	SYS1805: Added MANDTASKSOUTSTANDING error.
JLD		10/01/01	SYS1805: Altered message numbers.
JLD		29/01/01	SYS1832: changed GetAttribute to return "" instead of null if the attribute is not found
JLD		22/02/01	Added GetGlobalParameterString
APS		26/02/01	SYS1982: Added two new methods for SelectSingleNode and SelectNodes
BG		10/05/01	SYS2096 Added new error NOTIMPLEMENTED for printing.
SR		28/08/01	SYS2412 - new method GetNodeAttribute
DS		07/02/02	SYS4023	- updated the default request header to include adminSystemState
DS		20/02/02	SYS4120 Fix for systems with no adminSystemState set.
-------------------------------------------------------------------------
BMIDS Specific
GD		14/10/2002	BMIDS00572 - Added 'WRONGPASSWORD' as a trapable error in TranslateErrorType
MV		14/10/2002	BMIDS00630	Added EncodeXML() and SetErrorResponse()
MV		30/10/2002	BMIDS00780	Added EncodeXML() and SetErrorResponse() in XMLObject()
MV		30/10/2002	BMIDS00754	Amended CheckResponse()
MO		25/11/2002	BMIDS01076  Added 'AUTOTASKERROR' as a trapable error in TranslateErrorType
GHun	09/04/2003	BM0462		Problem with nested ASP tags in RunASP
MC		30/04/2004	BM0468		Cancel and Decline stage freeze screens functionality
KRW     15/04/2004  BMIDS774    GetComboIdsForValidation
HMA     19/07/2004  BMIDS798    Fix combo clearing in PopulateComboFromXML
JD		11/08/2004	BMIDS841	Added CREDITCHECKOKIMPORTFAILED error code.
-------------------------------------------------------------------------
MARS SPECIFIC
JD		04/05/2006	MAR1616		Added GetValidationForValueName
IK		05/06/2006	MAR1849		combo & global param cache
GHun	09/06/2006	MAR1849		changed caching to not freeze on popups
GHun	12/06/2006	MAR1849		Changed caching to also work for popups
GHun	15/06/2006	MAR1849		Minor tidy up and removed unnecessary loops
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM SPECIFIC
INR		14/11/2006	EP2_12		Added PopulateComboFromXMLByValidation
GHun	24/01/2007	EP2_921		Apply EP1288 Added COMMANDFAILED to error types.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

List the external function calls and their parameters here:

XMLObject()
	Generates an object containing an XML document and functions to manipulate it.
			
The contents of XMLObject() are:
		
		
ELEMENTS:

XMLDocument		- an XMLDOM document object. All the standard IE5 XML functionality is available through this.
ActiveTag		- Used to store a tag (node object) which is then used by the functions within XMLObject()
				  as the 'active' one. i.e. all actions are based on this tag.
ActiveBeforeTag	- Used to store a tag (node object) which is then used in the CreateTag_UseBeforeText()
				  processing.
ActiveTagList	- Used to store a tag list (node list object) which is then used by the functions within
				  XMLObject() as the 'active' one.
		
		
FUNCTIONS:

TagNew = CreateActiveTag(sTagName)
	sTagName	- The name of the tag to be created.
	TagNew		- The newly created node object.

	Creates a new tag (node object).
	Attaches it as a child to the ActiveTag, or the base XML document if ActiveTag = null.
	Sets the new tag as the ActiveTag.
	Clears the ActiveBeforeTag.
	Meant primarily for tags which define blocks e.g.
	<TAGNEW>
		<SOME>data</SOME>
		<OTHER>data</OTHER>
		<TAGS>data</TAGS>
	</TAGNEW>


TagNew = CreateTag(sTagName, vText)
	sTagName	- The name of the tag to be created.
	vText		- The text to be assigned to the tag
	TagNew		- The newly created node object.

	Creates a new tag (node object).
	Attaches it as a child to the ActiveTag. An active tag is required.
	Assigns the specified text to the tag. If no text is required, set sText to null.
	Meant primarily for data-bearing tags contained within tag blocks e.g.
	<TAGBLOCK>
		<TAGNEW>sText</TAGNEW>
	</TAGBLOCK>


TagNew = CreateTag_UseBeforeText(sTagName)
	sTagName	- The name of the tag to be created.
	TagNew		- The newly created node object.

	Creates a new tag (node object).
	Attaches it as a child to the ActiveTag. An active tag is required.
	If ActiveBeforeTag is set, it searches this for a child which matches the sTagName. If a match
	is found, then the text from that tag is copied to the newly created one. Otherwise no text is written.
	Meant primarily for tags which are not generated from the screen data.  When performing an update
	action, the tag must contain the same data in the <UPDATE TYPE="AFTER"> section as it does in
	the <UPDATE TYPE="BEFORE"> section.


SetAttribute(sAttributeName, vAttributeValue)
	sAttributeName	- The name of the tag attribute to be set.
	vAttributeValue	- The value that the attribute is to be set to.
			
	Sets an attribute on the ActiveTag. An active tag is required.
	The attribute value is automatically enclosed in "".
	Attributes are usually set on block tags such as <REQUEST> and <UPDATE> e.g.
	<REQUEST sAttributeName="vAttributeVale">
		<SOME>data</SOME>
		<OTHER>data</OTHER>
		<TAGS>data</TAGS>
	</REQUEST>

SetAttributeOnTag(sTagName, sAttributeName, vAttributeValue)
	sTagName		- The name of the tag to which the attribute applies
	sAttributeName	- The name of the tag attribute to be set.
	vAttributeValue	- The value that the attribute is to be set to.
			
	Sets an attribute on the specified Tag. An active tag is required.
	Searches from the Active tag to find the first occurrence of the specified Tag.
	The attribute value is automatically enclosed in "".
	Attributes are usually set on block tags such as <REQUEST> and <UPDATE> e.g.
	<REQUEST sAttributeName="vAttributeVale">
		<SOME>data</SOME>
		<OTHER>data</OTHER>
		<TAGS>data</TAGS>
	</REQUEST>

*************************************************************************************************************
This group of functions are designed to provide a central method of creating the REQUEST tag along
with its standard set of attributes.  They should be used as follows:

For creating a REQUEST tag on a main screen, simply call CreateRequestTag.

For creating a REQUEST tag on a popup, call CreateRequestAttributeArray on the main screen
and pass the created array into the popup.  In the popup use CreateRequestTagFromArray, giving
it the array passed in.

AttributeArray = CreateRequestAttributeArray(thisWindow)
	thisWindow		- the window object of the calling page
	AttributeArray	- the array containing the values of the attributes required

	Creates an array containing the values of the following context parameters, in this order:
	UserID, UnitId, MachineId, DistributionChannelId
	WILL NOT WORK ON A POPUP

TagNew = CreateRequestTagFromArray(AttributeArray,sOperation)
	AttributeArray	- the array created by CreateRequestAttributeArray
	sOperation		- the Operation attribute.  Set to name of method to call. Required.
	TagNew	- The newly created node object.

	Creates a Request active tag and its required attributes using the array created.
	The sOperation parameter is used to set the OPERATION attribute to the appropriate value
	
TagNew = CreateRequestTag(thisWindow,sOperation)
	thisWindow	- the window object of the calling page
	sOperation	- the operation attribute.  Set to the name of the method to call. Required.
	TagNew		- The newly created node object.

	Creates a REQUEST active tag containing its required attributes by calling CreateRequestAttributeArray
	followed by CreateRequestTagFromArray
	WILL NOT WORK ON A POPUP
*************************************************************************************************************

sText = GetTagText(sTagName)
	sTagName	- The name of the tag to be read.
	sText		- The text of the tag.
			
	Searches the ActiveTag for a child which matches the sTagName.  An active tag is required.
	If a match is found the text from that tag is returned, otherwise an empty string is returned.
	Meant primarily for reading data-bearing tags.  As the destination of the text is often a screen
	field an empty string must be returned if the tag doesn't exist. If null is returned the word null
	appears in the field.

dValue = GetTagFloat(sTagName)
	sTagName	- The name of the tag to be read.
	dValue		- The value of the tag as a floating point number.
			
	Calls GetTagText and then from the return tag value converst the value into a 
	floating point number and returns the number.
		
iValue = GetTagInt(sTagName)
	sTagName	- The name of the tag to be read.
	iValue		- The value of the tag as a integer number.
			
	Calls GetTagText and then from the return tag value converst the value into a 
	integer and returns the number.
		
SetTagText(sTagName, sText)
	sTagName	- The name of the existing tag.
	sText		- The new text of the tag.
			
	Searches the ActiveTag for a child which matches the sTagName.  An active tag is required.						
	If an the serach finds the specified tagname then the associated text for that tag is
	replaced by sText.
			

bValue = GetTagBoolean(sTagName)
	sTagName	- The name of the tag to be read.
	sValue		- The value of the tag as a boolean
			
	Searches the ActiveTag for a child which matches the sTagName.  An active tag is required.
	If a match is found the boolean value (true or false) is calculated based on the text returned 
	being "1" equalling true or other values equal false. 			
	Meant primarily for reading data-bearing tags.


sText = GetTagAttribute(sTagName,sAttributeName)
	sTagName		- The name of the tag to be read.
	sAttributeName	- The name of the attribute to be read.
	sText			- The value of the attribute.
			
	Searches the ActiveTag for a child which matches the sTagName.  An active tag is required.
	If a match is found the attribute on this tag is read.
	Returns the value of the attribute.
	Normally attributes exist on block tags but in the case of, for example, fields which store
	combo based data then the text which the stored data represents is returned as a TEXT attribute.
			

sText = GetAttribute(sAttributeName)
	sAttributeName	- The name of the attribute to be read.
	sText			- The value of the attribute.
			
	Reads an attribute on the ActiveTag. An active tag is required.
	Returns the value of the attribute.
	Attributes are usually read from block tags (see SetAttribute).


RemoveActiveTag()
	Removes the ActiveTag from the XML Document.
	Sets the ActiveTag to the deleted tag's parent, unless the parent is the top level document node,
	in which case set the ActiveTag to null.
		
		
TagListNew = CreateTagList(sTagName)
	sTagName	- The name of the tags to search for.
	TagListNew	- The newly created node list object
			
	Creates a node list object containing all the tags which match the specified name.
	If ActiveTag is set, only the XML from the active tag down is searched, otherwise the entire XML is searched.
	Sets ActiveTagList to the node list object, even if no matches are found.
	For example, in a block such as:
	<PARENT>
		<CHILD>data</CHILD>
		<CHILD>more data</CHILD>
	</PARENT>
			
	With ActiveTag equal to the parent, a CreateTagList("CHILD") would generate a node list object
	containing 2 items.


bReturn = SelectTagListItem(nItem)
	nItem	- The zero based index of the item to select from ActiveTagList.
	bReturn	- True if the item exists, false if not.
			
	Sets the ActiveTag to the ActiveTagList item matching the nItem index.  An active tag list is required.
	If there is no item at the nItem index then ActiveTag is set to null.  Using the example for
	CreateTagList, SelectTagListItem(1) would set ActiveTag to the <CHILD> node object containing 'more data'.
			

TagReturn = SelectTag(TagParent,sTagName)
	TagParent	- The tag node to search for the specified tag. If null, the whole document is searched.
	sTagName	- The name of the tag to search for.
	TagReturn	- The selected node object.

	This function is designed to search for a single tag matching the name specified and set the
	ActiveTag to that tag.  If more than one is found, the first one is selected.
	Designed to make searching for several block tags in a row easier by effetively combining the
	processing which would be done by a CreateTagList("ITEM") and SelectTagListItem(0) process.  It
	also circumvents the need to reset the ActiveTag to the tag you wish to search.


TagReturn = GetTagListItem(nItem)
	nItem		- The zero based index of the item to select from ActiveTagList.
	TagReturn	- The selected node object.

	Similar to SelectTagListItem, except it returns the ActiveTagList item matching the nItem indexand
	does not set ActiveTag.
	If there is no item at the nItem index then null is returned.
	Used primarily to store node objects for later use, such as in update-before processing.
			
			
LoadXML(sXML)
	sXML	- The XML string to be loaded.
			
	Loads the XML string into the XML document, replacing any previous contents.
	Clears the ActiveTag, ActiveBeforeTag and ActiveTagList.


AddXMLBlock(XMLSource)
	XMLSource	- The XML document or document fragment to add.
			
	Copies and adds all children (and their siblings) to the ActiveTag, or the base XML document
	if ActiveTag = null.
	Used, for example, to add the 'before' image to the <UPDATE TYPE="BEFORE"> tag.


bReturn = SelectUPDATE_Before()
	bReturn	- True if successful, otherwise false.
			
	Searches the XML document for the <UPDATE TYPE="BEFORE"> tag.
	If successful, sets the ActiveTag to this tag and returns true, otherwise sets ActiveTag to null
	and returns false.


XMLFragment = CreateFragment()
	XMLFragment	- The document fragment created.
			
	Creates a document fragment object.
	Copies and adds all children (and their siblings) of the ActiveTag to the document fragment.
	An ActiveTag is required.
	Note that the ActiveTag itself is not included.
	Used, for example, to make a copy of the XML on entry ready for the update processing.

RunASP(thisDocument,sASPFile)
	thisDocument	- the document object of the calling page
	sASPFile		- The name of the ASP file to run.
			
	Performs ResetXMLDocument.
	The ASP file must be one which generates an XML string. It is expected to be in the xml virtual directory.
	Loads the result of the file run into the XML document, replacing any previous contents.


RunASPWithTextInput(thisDocument,sASPFile,sRequest)
	thisDocument	- the document object of the calling page
	sASPFile		- The name of the ASP file to run.
	sRequest		- The request for the ASP file in text form.
			
	This function operates as above, except that instead of generating the request from the contents of
	the XML document it uses the string passed in.
	Meant primarily for such calls as GlobalParameterBO.GetCurrentParameter, where merely the name of the
	parameter is required.

		
WriteXMLToFile(sFileName)	***** FOR DEBUG PURPOSES ONLY *****
	Writes the XML string contained in the XML document to the file specified.
	Allows the developer to view any XML using the Microsoft XML Notepad.
	sFileName	- Name of the file to write to. If this is an empty string,
				  the xml will be written to "c:\script.xml".

ReadXMLFromFile(sFileName)  ***** FOR USE WHEN OBJECTS NOT YET AVAILLABLE *****
	Returns a string containing the XML held within sFileName
	Useful for imitating the return XML from an object method call which is not yet availlable.
	Use the return string to load into an XML object and continue as normal.

bReturn = GetComboLists(thisDocument,sGroupList)
	thisDocument	- the document object of the calling page
	sGroupList		- An array of combo groups
	bReturn			- True if no errors occur, otherwise false
			
	Performs ResetXMLDocument.
	Retrieves the data for all groups specified in the list and loads it into the XML document.


XMLFragment = GetComboListXML(sListName)
	sListName	- The group name for the combo list required from the XML document.
	XMLFragment	- The document fragment created.

	Searches the XML document for the combo group specified.
	If the call is successful a document fragment containing the result of the query is returned,
	otherwise null is returned.
	This function should be used in conjunction with GetComboLists.


bReturn = PopulateCombo(thisDocument,refFieldId,sListName,bAddSELECT)
	thisDocument	- the document object of the calling page
	refFieldId		- The combo field to be populated (e.g. frmScreen.cbo????)
	sListName		- The group name for the combo list required from the XML document.
	bAddSelect		- If true, the <SELECT> option is added. 
	bReturn			- True if the combo population was successful, false if not.
						
	Performs GetComboListXML, passing it sListName.
	Performs PopulateComboFromXML, passing it the XML fragment returned from GetComboListXML.

	
bReturn = PopulateComboFromXML(thisDocument,refFieldId,XMLSource,bAddSELECT)
	thisDocument	- the document object of the calling page
	refFieldId		- The combo field to be populated (e.g. frmScreen.cbo????)
	XMLSource		- The XML fragment to be used for populating the combo.
	bAddSelect		- If true, the <SELECT> option is added. 
	bReturn			- True if the XML fragment is not null, false if not.
			
	Loads the contents of the XML fragment into a newly created XML document.
	Any current contents of the combo field are cleared out.
	If required,a <SELECT> option is added to the combo, followed by all options retrieved 
	from the database.  Any VALIDATIONTYPE data is added as attributes to the <OPTION> tag.
 
			
bReturn = IsInComboValidationList(thisDocument,sListName, sValueID, ValidationList)
	thisDocument	- the document object of the calling page
	sListName		- The group name for the combo list required from the database.
	sValueID		- The combo group entry value-identification number
	ValidationList	- An array of validation strings
						
	Performs ResetXMLDocument.
	Retrieves the combo validation entries for the specified sListName and sValueID.
	Loads the result of the RunASP call to GetComboValidationList.asp into the XML document.
	Compares the validation entries returned in the XML document against those passed in.
	Returns true if a match is found.
	Returns false otherwise.
											
sReturn = GetComboIdForValidation(sListName, sValidation, XMLCombo)
	sListName		- The group name for the combo list required from the database.
	sValidation		- Validation Id for which combo value is to be found
	XMLCombo		- Optional, Use this to find the value if not null. Else, generate the ComboXML
	
	Find the Combo ValueId corresponding to this validation and return. If there are multiple IDs for the 
	the same validation, return the first one.
	
ResetXMLDocument()
	Deletes and recreates the XML document, thereby losing all the current contents.
	Clears the ActiveTag, ActiveBeforeTag and ActiveTagList.
	This function should not normally be explicitly called.  It is used by all the functions which reload
	the XML document.


bReturn = IsResponseOK()
	bReturn	- True if <RESPONSE TYPE="SUCCESS">, else false.
			
	Calls CheckResponse with an argument of null.


ErrorReturn = CheckResponse(sErrorType)
	sErrorType	- A one dimensional array of error types i.e. "RECORDNOTFOUND"
	ErrorReturn	- A one dimensional array, which is returned to the calling routine
				ErrorType[0] - True if <RESPONSE TYPE="SUCCESS">, else false.
				ErrorType[1] - The ErrorType i.e. "RECORDNOTFOUND" which the method failed on, if
				               that error is one contained in sErrorType. Else it is null.
				ErrorType[2] - The XML error description if ErrorType[1] is set. Else it is null.
			
	Checks for a <RESPONSE> tag and if one is found checks the Error number returned against
	the error number(s) associated with the error type(s) passed into the routine.
	If there is an error false is returned.  If the error is contained within sErrorType then
	the error type and message are returned.
									
 		
iErrorNumber = TranslateErrorType(sError)
	sError		 - The error type to translate i.e. RECORDNOTFOUND
	iErrorNumber - The VB object error number as returned in the XML response block 
					of a server side ASP call
			
	Translates the supplied sError description into its associated iErrorNumber and returns 
	this number. This function is necessary because we must always check against the 
	Error Number and not the Error description.		

bReturn = GetGlobalParameterBoolean(thisDocument,sParameterName)
	thisDocument	- the document object of the calling page
	sParameterName	- the name of the boolean type field to read from the global parameter table.
	bReturn			- the value of the field, or false if not found.
			
	Performs ResetXMLDocument.
	Gets the value of the specified boolean type field from the global parameter table.
	Loads the result of the global parameter access into the XML document.
		
sAmount = GetGlobalParameterAmount(thisDocument,sParameterName)
	thisDocument	- the document object of the calling page
	sParameterName	- the name of the boolean type field to read from the global parameter table.
	sAmount			- the value of the field, or "" if not found.
			
	Performs ResetXMLDocument.
	Gets the value of the specified boolean type field from the global parameter table.
	Loads the result of the global parameter access into the XML document.
	
sString = GetGlobalParameterString(thisDocument,sParameterName)
	thisDocument	- the document object of the calling page
	sParameterName	- the name of the string type field to read from the global parameter table.
	sValue			- the value of the field, or "" if not found.
			
	Performs ResetXMLDocument.
	Gets the value of the specified string type field from the global parameter table.
	Loads the result of the global parameter access into the XML document.

blnReturn = ConvertBoolean(sTagText)
	sTagText	- the tag value to convert
	blnReturn	- the boolean value of the converted tag value

TagListNew = SelectNodes(sPattern)
	sPattern	- XSL pattern defining the type of nodelist to search for.
	TagListNew	- The newly created node list object
			
	Creates a node list object containing all the tags which match the specified XSL pattern. If no
	tags match the XSL pattern then we set the ActiveTagList = null. If an Active tag has been 
	specified then search will start from this tag unless...
	
	WARNING:- 
	
	Even if we have specified an ActiveTag but we use the following XSL pattern is used "//author" 
	then this means the search will commence from the root of the XML document!!
	
TagReturn = SelectSingleNode(sPattern)
	sPattern	- XSL pattern defining the type of node to search for.
	TagReturn	- The selected node object.

	This function is designed to search for a single tag matching the XSL pattern specified and set 
	the ActiveTag to that tag. Thios method will set the ActiveTag to Null if no tag is found that 
	matches the XSL pattern. If an Active tag has been specified then search will start from 
	this tag unless...
	
	WARNING:- 
	
	Even if we have specified an ActiveTag but we use the following XSL pattern is used "//author" 
	then this means the search will commence from the root of the XML document!!
			
*/
%>
<html id="XMLFunctions">
<script language="JavaScript" type="text/javascript">
public_description = new CreateXMLFunctions;

function CreateXMLFunctions()
{
	this.XMLObject	= XMLObject;
}

function XMLObject()
{
	this.XMLDocument = new ActiveXObject("microsoft.XMLDOM");
	this.ActiveTag = null;
	this.ActiveBeforeTag = null;
	this.ActiveTagList = null;

<%	// Tag creation functions %>
	this.CreateActiveTag = CreateActiveTag;
	this.CreateTag = CreateTag;
	this.CreateTag_UseBeforeText = CreateTag_UseBeforeText;
	this.SetAttribute = SetAttribute;
	this.SetAttributeOnTag = SetAttributeOnTag;
	this.CreateRequestAttributeArray = CreateRequestAttributeArray;
	this.CreateRequestTagFromArray = CreateRequestTagFromArray;
	this.CreateRequestTag = CreateRequestTag;

<%	// Tag read functions %>
	this.GetTagText = GetTagText;
	this.GetTagFloat = GetTagFloat;
	this.GetTagInt = GetTagInt;
	this.GetTagBoolean = GetTagBoolean;
	this.ConvertBoolean = ConvertBoolean;
	this.GetTagAttribute = GetTagAttribute;
	this.GetAttribute = GetAttribute;
	this.GetNodeAttribute = GetNodeAttribute;

<%	// Tag Write functions %>
	this.SetTagText = SetTagText;

<%	// Tag deletion functions %>
	this.RemoveActiveTag = RemoveActiveTag;

<%	// Tag List functions %>
	this.CreateTagList = CreateTagList;
	this.SelectTagListItem = SelectTagListItem;
	this.SelectTag = SelectTag;
	this.GetTagListItem = GetTagListItem;

<%	// XML block functions %>
	this.LoadXML = LoadXML;
	this.AddXMLBlock = AddXMLBlock;
	this.SelectUPDATE_Before = SelectUPDATE_Before;
	this.CreateFragment = CreateFragment;
	this.RunASP = RunASP;
	this.RunASPWithTextInput = RunASPWithTextInput;

<%	// Miscellaneous %>
	this.WriteXMLToFile = WriteXMLToFile;
	this.ReadXMLFromFile = ReadXMLFromFile;
	this.GetComboLists = GetComboLists;
	this.GetComboListXML = GetComboListXML;
	this.PopulateCombo = PopulateCombo;
	this.PopulateComboFromXML = PopulateComboFromXML;
	this.PopulateComboFromXMLByValidation = PopulateComboFromXMLByValidation;//EP2_12
	this.IsInComboValidationList = IsInComboValidationList;
	this.IsInComboValidationXML = IsInComboValidationXML;
	this.ResetXMLDocument = ResetXMLDocument;
	this.IsResponseOK = IsResponseOK;
	this.GetGlobalParameterBoolean = GetGlobalParameterBoolean;
	this.GetGlobalParameterAmount = GetGlobalParameterAmount;
	this.GetGlobalParameterString = GetGlobalParameterString;
	this.GetComboDescriptionForValidation = GetComboDescriptionForValidation;
	this.GetComboIdForValidation = GetComboIdForValidation;
	this.TranslateErrorType = TranslateErrorType;
	this.CheckResponse = CheckResponse;
	this.SelectNodes = SelectNodes;
	this.SelectSingleNode = SelectSingleNode;
	this.EncodeXML = EncodeXML; 
	this.SetErrorResponse = SetErrorResponse;  
	this.GetComboIdsForValidation = GetComboIdsForValidation; //KRW BMIDS774
	this.GetValidationForValueName = GetValidationForValueName; //MAR1616
			
}

function CreateActiveTag(sTagName)
{
	var TagNew = this.XMLDocument.createElement(sTagName);

	if(this.ActiveTag != null)
		this.ActiveTag.appendChild(TagNew);
	else
		this.XMLDocument.appendChild(TagNew);

	this.ActiveTag = TagNew;
	this.ActiveBeforeTag = null;

	return TagNew;
}

function CreateTag(sTagName, vText)
{
<%	/*	There must be an ActiveTag
		There doesn't have to be any text, vText may be null
	*/
%>	var TagNew = this.XMLDocument.createElement(sTagName);

	if(this.ActiveTag != null)
		this.ActiveTag.appendChild(TagNew);
	else
		alert("CreateTag(" + sTagName + ") - No Active Tag");

	if(vText != null)
		TagNew.text = vText;

	return TagNew;
}

function CreateTag_UseBeforeText(sTagName)
{
<%	/*	There must be an ActiveTag
		If there is an ActiveBeforeTag:
			The active before tag should be a tag block, which is then searched to find a tag name which
			matches the one created. If one is found the text in the new tag is set to match the text
			in the old one. Only one match is assumed
	*/
%>	var TagNew = this.XMLDocument.createElement(sTagName);

	if(this.ActiveTag != null)
	{
		this.ActiveTag.appendChild(TagNew);

		if(this.ActiveBeforeTag != null)
		{
			var TagList = this.ActiveBeforeTag.getElementsByTagName(sTagName);
			if(TagList.length > 0)
				TagNew.text = TagList.item(0).text;
		}
	}
	else
		alert("CreateTag_UseBeforeText(" + sTagName + ") - No Active Tag");

	return TagNew;
}

function SetAttribute(sAttributeName, vAttributeValue)
{
<%	/*	There must be an active tag
		RF 09/11/99 Attribute value must be specified
	*/
%>	if(this.ActiveTag != null)
	{
		if (vAttributeValue !=null)
			this.ActiveTag.setAttribute(sAttributeName, vAttributeValue);
		else
			alert("SetAttribute - Null value for attribute " + sAttributeName);
	}
	else
		alert("SetAttribute(" + sAttributeName + ") - No Active Tag");
	
}

function SetAttributeOnTag(sTagName, sAttributeName, vAttributeValue)
{
<%	// There must be an active tag
%>	if(this.ActiveTag != null)
	{
		var TagList = this.ActiveTag.getElementsByTagName(sTagName);
		if (TagList.length > 0)
			TagList.item(0).setAttribute(sAttributeName, vAttributeValue);
	}
	else
		alert("SetAttributeOnTag(" + sTagName + ") - No Active Tag");
}

function CreateRequestAttributeArray(thisWindow)
{
	var AttributeArray = new Array(6);
	var frmContext;
	
	<% /* MAR1849 GHun Removed unnecessary loop */ %>
	try
	{
		frmContext = thisWindow.parent.frames("omigamenu").document.forms("frmContext");
	}
	catch (e)
	{
		frmContext = null;
	}
	<% /* MAR1849 End */ %>

	if (frmContext)
	{
		AttributeArray[0] = frmContext("idUserId").value;
		AttributeArray[1] = frmContext("idUnitId").value;
		AttributeArray[2] = frmContext("idMachineId").value;
		AttributeArray[3] = frmContext("idDistributionChannelId").value;
		//SYS4023 DS adminSystemState added
		AttributeArray[4] = frmContext("idAdminSystemState").value;
		AttributeArray[5] = frmContext("idRole").value;
	}
	else
	{
		AttributeArray[0] = thisWindow.prompt("Enter value for User ID","");
		AttributeArray[1] = thisWindow.prompt("Enter value for Unit ID","");
		AttributeArray[2] = thisWindow.prompt("Enter value for Machine ID","");
		AttributeArray[3] = thisWindow.prompt("Enter value for Channel ID","");
		//SYS4023 DS adminSystemState added
		AttributeArray[4] = thisWindow.prompt("Enter value for Role AdminSystemState","");
		AttributeArray[5] = thisWindow.prompt("Enter value for Role ID","");
	}

	return AttributeArray;
}

function CreateRequestTagFromArray(AttributeArray, sOperation)
{
	var TagNew = this.CreateActiveTag("REQUEST");
	this.SetAttribute("USERID", AttributeArray[0]);
	this.SetAttribute("UNITID", AttributeArray[1]);
	this.SetAttribute("MACHINEID", AttributeArray[2]);
	this.SetAttribute("CHANNELID", AttributeArray[3]);
	//SYS4023 DS adminSystemState added
	if (AttributeArray[4] == "" || AttributeArray[4] == null) {
		this.SetAttribute("ADMINSYSTEMSTATE", "");
	}
	else {
		this.SetAttribute("ADMINSYSTEMSTATE", AttributeArray[4]);
	}
	
	// APS SYS1808: Defensive coding for this tag as not all Popups have been changed
	if (AttributeArray.length == 6)
		this.SetAttribute("USERAUTHORITYLEVEL",AttributeArray[5]);
	
	if(sOperation != null) {
		this.SetAttribute("OPERATION",sOperation);
		this.SetAttribute("ACTION",sOperation);
	}
	return TagNew;
}

function CreateRequestTag(thisWindow, sOperation)
{
	return this.CreateRequestTagFromArray(CreateRequestAttributeArray(thisWindow), sOperation);
}

function GetTagText(sTagName)
{
<%	/*	There must be an ActiveTag
		Search the children of the ActiveTag for a match to sTagName
		If no match is found return an empty string, returning null directly into a screen field actually
		causes the word null to appear
	*/
%>	if(this.ActiveTag != null)
	{
		var TagList = this.ActiveTag.getElementsByTagName(sTagName);
		if(TagList.length > 0)
			return TagList.item(0).text;
	}
	else
		alert("GetTagText(" + sTagName + ") - No Active Tag");

	return "";
}

function GetTagFloat(sTagName)
{
	var sValue = this.GetTagText(sTagName);
	var dValue = 0.000;

	if(sValue != "")
		dValue = parseFloat(sValue);
	if(isNaN(dValue))
		alert("GetTagFloat - Tag value is not a number");

	return dValue;
}

function GetTagInt(sTagName)
{
	var sValue = this.GetTagText(sTagName);
	var iValue = 0;
	
	if(sValue != "")
		iValue = parseInt(sValue);
	if (isNaN(iValue))
		alert("GetTagInt - Tag value is not a number");

	return iValue;
}

function SetTagText(sTagName, sText)
{
<%	/* There must be an ActiveTag
		Search the active tag for a child which matches the tag name. If one is found set the text.
		Only one match is assumed
	*/
%>	if(this.ActiveTag != null)
	{
		var TagList = this.ActiveTag.getElementsByTagName(sTagName);
		if(TagList.length > 0)
			TagList.item(0).text = sText;
	}
	else
		alert("SetTagText(" + sTagName + ") - No Active Tag");
}

function GetTagBoolean(sTagName)
{
<%	/*	There must be an ActiveTag
	*/
%>	if(this.ActiveTag != null)
	{
		var TagList = this.ActiveTag.getElementsByTagName(sTagName);
		if(TagList.length > 0)
			return this.ConvertBoolean(TagList.item(0).text);
	}
	else
		alert("GetTagBoolean(" + sTagName + ") - No Active Tag");

	return false;
}

function ConvertBoolean(sTagText)
{
	if((sTagText == "0") || (sTagText == ""))
		return false;
	return true;
}

function GetTagAttribute(sTagName, sAttributeName)
{
<%	/*	There must be an ActiveTag
		Search the ActiveTag for a child which matches the tag name
	*/
%>	if(this.ActiveTag != null)
	{
		var TagList = this.ActiveTag.getElementsByTagName(sTagName);

		if(TagList.length > 0)
		{
			<% /* MAR1849 GHun Removed unnecessary loop */ %>
			var attrib = TagList.item(0).attributes.getNamedItem(sAttributeName);
			if (attrib)
				return(attrib.text);
			<% /* MAR1849 End */ %>
		}
	}
	else
		alert("GetTagAttribute(" + sTagName + ") - No ActiveTag");

	return "";
}

function GetAttribute(sAttributeName)
{
<%	// There must be an active tag %>
	if(this.ActiveTag != null)
	{
		<% /* MAR1849 GHun Removed unnecessary loop */ %>
		var attrib = this.ActiveTag.attributes.getNamedItem(sAttributeName);
		if (attrib)
			return(attrib.text);
		<% /* MAR1849 End */ %>
	}
	else
		alert("GetAttribute(" + sAttributeName + ") - No Active Tag");

	return "";
}

function GetNodeAttribute(xmlNode, sAttributeName)
{
	<% /* MAR1849 GHun Removed unnecessary loop */ %>
	var attrib = xmlNode.attributes.getNamedItem(sAttributeName);
	if (attrib)
		return(attrib.text);
	<% /* MAR1849 End */ %>
}

function RemoveActiveTag()
{
<%	/*	There must be an ActiveTag
		Get the parent of the active tag then remove the active tag from the parent tag
		If the parent tag itself has a parent then set the active tag to it
		otherwise it must be the document node so set active tag to null
	*/
%>	if(this.ActiveTag != null)
	{
		var TagParent = this.ActiveTag.parentNode;
		TagParent.removeChild(this.ActiveTag);

		if(TagParent.parentNode != null)
			this.ActiveTag = TagParent;
		else
			this.ActiveTag = null;
	}
	else
		alert("RemoveActiveTag - No Active Tag");
}

function CreateTagList(sTagName)
{
	var TagListNew = null;

	if(this.ActiveTag != null)
		TagListNew = this.ActiveTag.getElementsByTagName(sTagName);
	else
		TagListNew = this.XMLDocument.getElementsByTagName(sTagName);

	this.ActiveTagList = TagListNew;

	return TagListNew;
}

function SelectTagListItem(nItem)
{
<%	/*	There must be an ActiveTagList
		If the specified index is valid, set the active tag to that item and return true,
		otherwise set the active tag to null and return false.
	*/
%>	if (nItem >= 0)
	{
		if (this.ActiveTagList != null)
		{
			if (this.ActiveTagList.length >= nItem + 1)
			{
				this.ActiveTag = this.ActiveTagList.item(nItem);
				return true;
			}
			else
				this.ActiveTag = null;
		}
		else
			alert("SelectedTagListItem - No Active Tag List");
	}

	return false;
}

function SelectTag(TagParent, sTagName)
{
<%	/*	Generate a tag list for the specified tag, generated from the parent tag
		or the base document if the parent tag is null
		This function assumes there will be only one tag of that name returned in the list,
		so if at least one is found it sets the active tag to the first one in the list and returns it
	*/
%>	var TagReturn = null;
	var TagListNew;

	if(TagParent != null)
		TagListNew = TagParent.getElementsByTagName(sTagName);
	else
		TagListNew = this.XMLDocument.getElementsByTagName(sTagName);

	if(TagListNew.length > 0)
	{
		TagReturn = TagListNew.item(0);
		this.ActiveTag = TagReturn;
	}
	else
		this.ActiveTag = null;

	return TagReturn;
}

function GetTagListItem(nItem)
{
<%	/*	There must be an ActiveTagList
		If the specified index is valid return that item
		otherwise null will be returned.
	*/
%>	if(this.ActiveTagList != null)
	{
		if(this.ActiveTagList.length >= nItem + 1)
			return this.ActiveTagList.item(nItem);
	}
	else
		alert("SelectedTagListItem - No Active Tag List");

	return null;
}

function LoadXML(sXML)
{
	this.ResetXMLDocument();
	if(sXML != null)
		this.XMLDocument.loadXML(sXML);
	else
		this.XMLDocument.loadXML("");
}

function AddXMLBlock(XMLSource)
{
<%	/*	Get all the children from the base source document
		The child must be cloned and the clone added to the active tag or base document.
		If the child is added directly it is actually removed from the source document.
	*/
%>	if(XMLSource != null)
	{
		var TagListChildren = XMLSource.childNodes;

		for(var nLoop = 0;nLoop < TagListChildren.length;nLoop++)
		{
			var TagNew = TagListChildren.item(nLoop).cloneNode(true);
			if(this.ActiveTag != null)
				this.ActiveTag.appendChild(TagNew);
			else
				this.XMLDocument.appendChild(TagNew);
		}
	}
}

function SelectUPDATE_Before()
{
	var bReturn = false;
	var TagList = this.XMLDocument.getElementsByTagName("UPDATE");

	for(var nLoop = 0;nLoop < TagList.length && !bReturn;nLoop++)
	{
		if(TagList.item(nLoop).attributes.getNamedItem("TYPE").text == "BEFORE")
		{
			this.ActiveTag = TagList.item(nLoop);
			bReturn = true;
		}
	}

	if(!bReturn)
		this.ActiveTag = null;

	return bReturn;
}

function CreateFragment()
{
<%	/*	There must be an ActiveTag
		Add all children of the active tag to the document fragment.
		Note that the active tag itself is not added.
		The child must be cloned and the clone added to the document fragment.
		If the child is added directly it is actually removed from the source document.
	*/
%>	var XMLFragment = null;
		
	if(this.ActiveTag != null)
	{
		XMLFragment = this.XMLDocument.createDocumentFragment();
		var TagList = this.ActiveTag.childNodes;

		for(var nLoop = 0;nLoop < TagList.length;nLoop++)
		{
			var TagNew = TagList.item(nLoop).cloneNode(true);
			XMLFragment.appendChild(TagNew);
		}
	}
	else
		alert("CreateFragment - No Active Tag");

	return XMLFragment;
}

function RunASP(thisDocument, sASPFile)
{
	var thisHTTP = new ActiveXObject("Microsoft.XMLHTTP");
	this.XMLDocument.async = false;
	thisHTTP.open("POST", getURL(thisDocument, sASPFile), false);
<% /* DEBUGLOGEXTRASTART DO NOT EDIT
	var guidTrans = DebugLogCreateGuid();
	thisHTTP.setRequestHeader("TRANSACTIONGUID", guidTrans);
	if (this.XMLDocument.documentElement != null)
	{
		this.XMLDocument.documentElement.setAttribute("TRANSACTIONGUID", guidTrans);
	}
	DebugLogWriteLine(guidTrans, 0, +1, "xml", sASPFile, this.XMLDocument.xml);
DEBUGLOGEXTRAEND DO NOT EDIT */ %>
	thisHTTP.send(this.XMLDocument);
<% /* DEBUGLOGEXTRASTART DO NOT EDIT
	var nElapsedServerTime = parseInt(thisHTTP.getResponseHeader("ElapsedServerTime"));
	if (isNaN(nElapsedServerTime)) nElapsedServerTime = 0;
	DebugLogWriteLine(guidTrans, nElapsedServerTime,  -1, "xml", sASPFile, thisHTTP.responseText);
DEBUGLOGEXTRAEND DO NOT EDIT */ %>
	this.LoadXML(thisHTTP.responseText);
}
<% /* DEBUGLOGEXTRASTART DO NOT EDIT
function DebugLogCreateGuid()
{
	var strGuid = "";
	try
	{
		if (window.dialogArguments != null)
		{
			// Popup window.
			strGuid = window.dialogArguments[5].top.frames[1].document.all.scDebugFunctions.CreateGuid();
		}
		else
		{
			strGuid = top.frames[1].document.all.scDebugFunctions.CreateGuid();
		}
	}
	catch(e)
	{
		// Uncomment to see errors.
		// alert("Exception in DebugLog.asp:DebugLogCreateGuid(): " + e.description + " (" + e.number + ")");
	}
	return strGuid;
}
DEBUGLOGEXTRAEND DO NOT EDIT */ %>

function RunASPWithTextInput(thisDocument,sASPFile,sRequest)
{
	this.ResetXMLDocument();
	this.XMLDocument.async = false;  <%/* Ensures this code waits for a response to the load */%>
	this.XMLDocument.load(getURL(thisDocument,sASPFile) + "?Request=" + sRequest);
}

function getURL(thisDocument,sTarget)
{
<%	// Generates a URL for the target screen
%>	var s0 = new String(thisDocument.location.href);
	return (s0.substring(0, s0.lastIndexOf("/")) + "/xml/" + sTarget);
}

function WriteXMLToFile(sFileName)
{
	if (sFileName == "")
		sFileName = "c:\\script.xml";
	var fs = new ActiveXObject("Scripting.FileSystemObject");
	var a = fs.OpenTextFile(sFileName,2,true);
	a.Write(this.XMLDocument.xml);
	a.Close();
}

function ReadXMLFromFile(sFileName)
{
	var sXMLString = "";
	if (sFileName != "")
	{
		var fs = new ActiveXObject("Scripting.FileSystemObject");
		var a = fs.OpenTextFile(sFileName,1);
		if (a != null)
		{
			sXMLString = a.ReadAll();
			a.Close();
		}
		else
			alert("file " + sFileName + " not found");
	}
	return (sXMLString);
}

<%	// IK_MAR1849 cached replacement %>	
function GetComboLists(thisDocument,sGroupList)
{
<%	// Make sure at least one group is specified %>	
	if (sGroupList.length == 0)
	{
		alert("GetComboLists - No ListNames specified");
		return(false);
	}

	var xComboCache = window.top.xCombos;	<% /* MAR1849 GHun */ %>
	<% /* MAR1849 GHun If the cache is not found, it is probably due to the current window being a popup.
	 This relies on popups defining a global variable called m_BaseNonPopupWindow to reference the base window. */ %>
	if (!xComboCache)
	{
		try
		{ 
			xComboCache = window.top.m_BaseNonPopupWindow.top.xCombos;
		}
		catch (e)
		{
			xComboCache = null;
		}
	}
	<% /* MAR1849 End */ %>

	this.ResetXMLDocument();
	this.CreateActiveTag("REQUEST");
	this.CreateActiveTag("SEARCH");
	this.CreateActiveTag("LIST");
	
	var bRunIt = false;
	for(var nLoop = 0;nLoop < sGroupList.length;nLoop++) 
	{
		var bGetIt = false;
		if(xComboCache)
		{
			if(!xComboCache.documentElement.selectSingleNode("LIST/LISTNAME[@NAME='" + sGroupList[nLoop] + "']"))
				bGetIt = true;
		}
		else
			bGetIt = true;
		
		if(bGetIt)
		{
			this.CreateTag("LISTNAME", sGroupList[nLoop]);
			bRunIt = true;
		}
	}

	if(bRunIt)
	{		
		this.RunASP(thisDocument, "GetComboList.asp");
		if(!this.IsResponseOK())
			return false;

		if(xComboCache) 
		{
			var e = new Enumerator(this.XMLDocument.selectSingleNode("RESPONSE/LIST").childNodes);
			for(; !e.atEnd(); e.moveNext())
				xComboCache.selectSingleNode("RESPONSE/LIST").appendChild(e.item().cloneNode(true));
		}
	}
	else
		this.XMLDocument.loadXML("<RESPONSE><LIST/></RESPONSE>");
	
	if(xComboCache)
	{
		for(var nLoop = 0;nLoop < sGroupList.length;nLoop++) 
		{
			if(!this.XMLDocument.documentElement.selectSingleNode("LIST/LISTNAME[@NAME='" + sGroupList[nLoop] + "']"))
			{
				this.XMLDocument.documentElement.selectSingleNode("LIST").appendChild
				(xComboCache.documentElement.selectSingleNode("LIST/LISTNAME[@NAME='" + sGroupList[nLoop] + "']").cloneNode(true));
			}
		}
	}

	this.ActiveTag = this.XMLDocument.documentElement;
	return true;
}
<%	// IK_MAR1849 ends %>	

function GetComboListXML(sListName)
{
<%	/*	Search the XML document for the group specified
		Generate a document fragment from the result
	*/
%>	var XMLFragment = null;
	var TagListLISTNAME = this.CreateTagList("LISTNAME");

	for(var nLoop = 0;this.SelectTagListItem(nLoop);nLoop++)
	{
		if(this.GetAttribute("NAME") == sListName)
		{
			XMLFragment = this.XMLDocument.createDocumentFragment();
			var TagNew = this.ActiveTag.cloneNode(true);
			XMLFragment.appendChild(TagNew);
		}
	}

	return XMLFragment;
}

function GetComboDescriptionForValidation(sListName, sValidation)
{
<%	/*  Search the XML document for the group specified and from this group
		return the description with given validation type. A particular value can
		have multiple validation ids.
	*/
%>
	var sComboValueName;
	var XMLFragment = this.GetComboListXML(sListName);
	
	if(XMLFragment != null)
	{
		var XML = new XMLObject();
		XML.AddXMLBlock(XMLFragment);
		
		var TagListLISTENTRY = XML.CreateTagList("LISTENTRY");
		
		for(var nLoop = 0;nLoop < TagListLISTENTRY.length;nLoop++)
		{
			XML.ActiveTagList = TagListLISTENTRY;
			XML.SelectTagListItem(nLoop);
			
			sComboValueName = XML.GetTagText("VALUENAME");
			
			XML.CreateTagList("VALIDATIONTYPE");
			for(var nListLoop = 0;nListLoop < XML.ActiveTagList.length;nListLoop++)
			{
				if(XML.ActiveTagList.item(nListLoop).text == sValidation) return sComboValueName;
			}			
		}
	} 
}
function GetValidationForValueName(sListName, sValueName, XMLCombo, thisDocument)
{
<%	/*  Search the XML document for the group specified and, from this group return the first validation with the given Valueid.
		If multiple Ids have the same validation, return the Id that comes first in the XML. 
		
		If XMLCombo(Input) is null, create the XML with the required combo values, before finding the ID, else
		use the Input XML. The input param thisDocument is required to be passed in for building the comboXML.
	*/
%>

	var sComboValueId;	
	var bBuildComboXML = false; 
	var XMLFragment;
	var XML = new XMLObject();
	
	if(arguments.length < 3)
		bBuildComboXML = true;
	else 
	{
		if(XMLCombo == null)	
			bBuildComboXML = true;
	}
	
	if(bBuildComboXML)
	{	
		if(arguments.length < 4)
			return "";
		if(thisDocument == null)
			return "";
		
	 	if(XML.GetComboLists(thisDocument, Array(sListName)))
	 	{
	 		XMLFragment = XML.GetComboListXML(sListName);
	 		XML.ResetXMLDocument();
	 	} 
	 }
	else
		XMLFragment = XMLCombo;

	XML.AddXMLBlock(XMLFragment);
	
	var TagListLISTENTRY = XML.CreateTagList("LISTENTRY");
	for(var nLoop = 0;nLoop < TagListLISTENTRY.length;nLoop++)
	{	
		XML.ActiveTagList = TagListLISTENTRY;
		XML.SelectTagListItem(nLoop);
		
		sComboValueId = XML.GetTagText("VALUENAME");
		XML.CreateTagList("VALIDATIONTYPE");
		for(var nListLoop = 0; nListLoop < XML.ActiveTagList.length; nListLoop++)
		{
			if(sComboValueId == sValueName)
				return XML.ActiveTagList.item(nListLoop).text;
		}
	}
}
function GetComboIdForValidation(sListName, sValidation, XMLCombo, thisDocument)
{
<%	/*  Search the XML document for the group specified and, from this group return the ValueId with the given validation type.
		If multiple Ids have the same validation, return the Id that comes first in the XML. 
		
		If XMLCombo(Input) is null, create the XML with the required combo values, before finding the ID, else
		use the Input XML. The input param thisDocument is required to be passed in for building the comboXML.
	*/
%>

	var sComboValueId;	
	var bBuildComboXML = false; 
	var XMLFragment;
	var XML = new XMLObject();
	
	if(arguments.length < 3)
		bBuildComboXML = true;
	else 
	{
		if(XMLCombo == null)
			bBuildComboXML = true;
	}
	
	if(bBuildComboXML)
	{	
		if(arguments.length < 4)
			return "";
		if(thisDocument == null)
			return "";
		
	 	if(XML.GetComboLists(thisDocument, Array(sListName)))
	 	{
	 		XMLFragment = XML.GetComboListXML(sListName);
	 		XML.ResetXMLDocument();
	 	} 
	 }
	else
		XMLFragment = XMLCombo;

	XML.AddXMLBlock(XMLFragment);
	
	var TagListLISTENTRY = XML.CreateTagList("LISTENTRY");
	for(var nLoop = 0;nLoop < TagListLISTENTRY.length;nLoop++)
	{	
		XML.ActiveTagList = TagListLISTENTRY;
		XML.SelectTagListItem(nLoop);
		
		sComboValueId = XML.GetTagText("VALUEID");
		XML.CreateTagList("VALIDATIONTYPE");
		for(var nListLoop = 0; nListLoop < XML.ActiveTagList.length; nListLoop++)
		{
			if(XML.ActiveTagList.item(nListLoop).text == sValidation)
				return sComboValueId;
		}
	}
}

function PopulateCombo(thisDocument,refFieldId,sListName,bAddSELECT,bAppend)
{
<%	// Check that the field is a combo
%>	if(refFieldId.tagName == "SELECT")
	{
		var vUndefined;
		var XMLFragment = this.GetComboListXML(sListName);
		if(bAppend != "undefined")
		{
			if(XMLFragment != null)
				return this.PopulateComboFromXML(thisDocument,refFieldId,XMLFragment,bAddSELECT,vUndefined,vUndefined,vUndefined,vUndefined,bAppend);
		}
		else
		{
			if(XMLFragment != null)
				return this.PopulateComboFromXML(thisDocument,refFieldId,XMLFragment,bAddSELECT);
		}
	}
	else
		alert("PopulateComboFromDB - field must be a combo");

	return false;
}

function PopulateComboFromXML(thisDocument,refFieldId,XMLSource,bAddSELECT,bAttibuteBase,ValueActiveTag,ValueIDTagName,ValueTagName,bAppend)
{
	/* This is a Function to Populate Combos  based on 
									1.Combo Values from DB 
									2.Element based Response Object
									3.Attribute based Response Object
									
	refFeildId =  Combobox Field Name 
	bAddSelect = Boolean Value( TRUE or FALSE ) whether to Add <SELECT id=select1 name=select1> as a first Element
	bAttribute = Boolean Value( TRUE or FALSE ) Sets to Attribute based ( DEFAULT ELEMENT BASED)
	ValueActiveTag = ActiveTag List name from Response object to Read 
	ValueIDTagName = The Field used for Combo Index 
	ValueTagName = The Field used to Fill Combo Display text
	
	Example :- 
	
	Element Based :										Attribute Based  :
	
	<STAGELIST>											<STAGELIST>
		<STAGE>												<STAGE  STAGEID = "10" STAGENAME = "XYZ" </STAGE>	
			<STAGEID>10</STAGEID>							<STAGE  STAGEID = "20" STAGENAME = "ABC" </STAGE>	
			<STAGENAME>XYX</STAGENAME>					</STAGELIST>
		<STAGE>
		<STAGE>												
			<STAGEID>20</STAGEID>
			<STAGENAME>ABC</STAGENAME>
		<STAGE>
	</STAGELIST>	
	
	Example :- 
	
		ValueActiveTag = "STAGE"
		ValueIDTagName = "STAGEID"
		VAlueTagName = "STAGENAME"
	
	*/
		
	var bReturn = false;
	var XMLCombo = null;
	
	if(refFieldId.tagName == "SELECT")
	{
		if(XMLSource != null)
		{
			bReturn = true;
			
			XMLCombo = new  XMLObject();
			XMLCombo.AddXMLBlock(XMLSource);
			if(bAppend != true)     <% // BMIDS798 %>
			{
				while(refFieldId.options.length > 0) refFieldId.options.remove(0);
			}
			 if(bAddSELECT)
			{
				var TagSELECT = thisDocument.createElement("OPTION");
				TagSELECT.value = "";
				TagSELECT.text = "<SELECT>";
				refFieldId.options.add(TagSELECT);
			} 
			
			if (typeof(ValueActiveTag) == "undefined") 
				var TagListLISTENTRY = XMLCombo.CreateTagList("LISTENTRY");
			else
				var TagListLISTENTRY = XMLCombo.CreateTagList(ValueActiveTag);
			
			for(var nLoop = 0;nLoop < TagListLISTENTRY.length;nLoop++)
			{
				XMLCombo.ActiveTagList = TagListLISTENTRY;
				XMLCombo.SelectTagListItem(nLoop);
				var TagOPTION = thisDocument.createElement("OPTION");
				
				if (typeof(ValueActiveTag) == "undefined" ) 
				{
					TagOPTION.value = XMLCombo.GetTagText("VALUEID");
					TagOPTION.text = XMLCombo.GetTagText("VALUENAME");
					refFieldId.options.add(TagOPTION);
					XMLCombo.CreateTagList("VALIDATIONTYPE");
					for(var nListLoop = 0;nListLoop < XMLCombo.ActiveTagList.length;nListLoop++)
						TagOPTION.setAttribute("ValidationType" + nListLoop, XMLCombo.ActiveTagList.item(nListLoop).text);
				}
				else
				if (bAttibuteBase == true) 
				{
					TagOPTION.value = XMLCombo.GetAttribute(ValueIDTagName);
					TagOPTION.text = XMLCombo.GetAttribute(ValueTagName);
					refFieldId.options.add(TagOPTION);
				}
				else
				{
					TagOPTION.value = XMLCombo.GetTagText(ValueIDTagName);
					TagOPTION.text = XMLCombo.GetTagText(ValueTagName);
					refFieldId.options.add(TagOPTION);
				}	
				
			}

			refFieldId.selectedIndex = 0;
			XMLCombo = null;
		}
		else
			alert("PopulateComboFromXML - document fragment is null");
	}
	else
		alert("PopulateComboFromXML - field must be a combo");

	return bReturn;
}

<%/*	EP2_12 New Function */%>
function PopulateComboFromXMLByValidation(thisDocument,refFieldId,XMLSource,bAddSELECT,bAttibuteBase,ValueActiveTag,ValueIDTagName,ValueTagName,bAppend,validateType)
{
	/* This is a Function to Populate Combos  based on 
									1.Combo Values from DB 
									2.Element based Response Object
									3.Attribute based Response Object
									4. ValidationType
	refFeildId =  Combobox Field Name 
	bAddSelect = Boolean Value( TRUE or FALSE ) whether to Add <SELECT id=select1 name=select1> as a first Element
	bAttribute = Boolean Value( TRUE or FALSE ) Sets to Attribute based ( DEFAULT ELEMENT BASED)
	ValueActiveTag = ActiveTag List name from Response object to Read 
	ValueIDTagName = The Field used for Combo Index 
	ValueTagName = The Field used to Fill Combo Display text
	validateType = The ValidationType to match against when loading values into	the combo
	
	Example :- 
	
	Element Based :										Attribute Based  :
	
	<STAGELIST>											<STAGELIST>
		<STAGE>												<STAGE  STAGEID = "10" STAGENAME = "XYZ" </STAGE>	
			<STAGEID>10</STAGEID>							<STAGE  STAGEID = "20" STAGENAME = "ABC" </STAGE>	
			<STAGENAME>XYX</STAGENAME>					</STAGELIST>
		<STAGE>
		<STAGE>												
			<STAGEID>20</STAGEID>
			<STAGENAME>ABC</STAGENAME>
		<STAGE>
	</STAGELIST>	
	
	Example :- 
	
		ValueActiveTag = "STAGE"
		ValueIDTagName = "STAGEID"
		VAlueTagName = "STAGENAME"
	
	*/
		
	var bReturn = false;
	var XMLCombo = null;
	
	if(refFieldId.tagName == "SELECT")
	{
		if(XMLSource != null)
		{
			bReturn = true;
			
			XMLCombo = new  XMLObject();
			XMLCombo.AddXMLBlock(XMLSource);
			if(bAppend != true)     <% // BMIDS798 %>
			{
				while(refFieldId.options.length > 0) refFieldId.options.remove(0);
			}
			if(bAddSELECT)
			{
				var TagSELECT = thisDocument.createElement("OPTION");
				TagSELECT.value = "";
				TagSELECT.text = "<SELECT>";
				refFieldId.options.add(TagSELECT);
			} 
			
			if ((typeof(ValueActiveTag) == "undefined") || (ValueActiveTag == null))
				var TagListLISTENTRY = XMLCombo.CreateTagList("LISTENTRY");
			else
			{
				var xP = ValueActiveTag;
				xP += "[VALIDATIONTYPELIST/@VALIDATIONTYPE='" + validateType + "']";
				var TagListLISTENTRY = XMLCombo.CreateTagList(xP);
				
			}
			for(var nLoop = 0;nLoop < TagListLISTENTRY.length;nLoop++)
			{
				XMLCombo.ActiveTagList = TagListLISTENTRY;
				XMLCombo.SelectTagListItem(nLoop);
				var TagOPTION = thisDocument.createElement("OPTION");
				
				if ((typeof(ValueActiveTag) == "undefined" ) || (ValueActiveTag == null ))
				{
					TagOPTION.value = XMLCombo.GetTagText("VALUEID");
					TagOPTION.text = XMLCombo.GetTagText("VALUENAME");
					XMLCombo.CreateTagList("VALIDATIONTYPE");
					for(var nListLoop = 0;nListLoop < XMLCombo.ActiveTagList.length;nListLoop++)
					{
						if(XMLCombo.ActiveTagList.item(nListLoop).text == validateType)
						{
							refFieldId.options.add(TagOPTION);
							for(var nListLoop = 0;nListLoop < XMLCombo.ActiveTagList.length;nListLoop++)
								TagOPTION.setAttribute("ValidationType" + nListLoop, XMLCombo.ActiveTagList.item(nListLoop).text);
						}
					}
				}
				else
				if (bAttibuteBase == true) 
				{
					TagOPTION.value = XMLCombo.GetAttribute(ValueIDTagName);
					TagOPTION.text = XMLCombo.GetAttribute(ValueTagName);
					refFieldId.options.add(TagOPTION);
				}
				else
				{
					TagOPTION.value = XMLCombo.GetTagText(ValueIDTagName);
					TagOPTION.text = XMLCombo.GetTagText(ValueTagName);
					refFieldId.options.add(TagOPTION);
				}	
				
			}

			refFieldId.selectedIndex = 0;
			XMLCombo = null;
		}
		else
			alert("PopulateComboFromXML - document fragment is null");
	}
	else
		alert("PopulateComboFromXML - field must be a combo");

	return bReturn;
}

function IsInComboValidationList(thisDocument,sListName, sValueID, ValidationList)
{
	var bReturn = false;

	this.ResetXMLDocument();
	this.CreateActiveTag("REQUEST");

	this.CreateActiveTag("SEARCH");
	this.CreateActiveTag("LIST");
	this.CreateTag("GROUPNAME",sListName);
	this.CreateTag("VALUEID", sValueID);

	this.RunASP(thisDocument,"GetComboValidationList.asp");
	if (this.IsResponseOK())
	{
		var TagListLISTENTRY = this.CreateTagList("LISTENTRY");

		for(var nLoop = 0;nLoop < TagListLISTENTRY.length;nLoop++)
		{
			this.ActiveTagList = TagListLISTENTRY;
			this.SelectTagListItem(nLoop);
			bReturn = this.IsInComboValidationXML(ValidationList);
		}
	}

	return bReturn;
}

function IsInComboValidationXML(ValidationList)
{
<%	// Get all <VALIDATIONTYPE> entries and add them as elements
	// in the array					
%>	this.CreateTagList("VALIDATIONTYPE");

	for(var nListLoop = 0;nListLoop < this.ActiveTagList.length;nListLoop++)
	{
		for (var nValLoop = 0;nValLoop < ValidationList.length;nValLoop++)
		{
			if (this.ActiveTagList.item(nListLoop).text == ValidationList[nValLoop])
				return true;
		}
	}

	return false;
}

function ResetXMLDocument()
{
	this.XMLDocument = null;
	this.XMLDocument = new ActiveXObject("microsoft.XMLDOM");
	this.ActiveTag = null;
	this.ActiveBeforeTag = null;
	this.ActiveTagList = null;
}

function IsResponseOK()
{
<%	// Checks the <RESPONSE> tag
	// APS 05/10/99 - Function implementation changed and XML code into 
	// the CheckResponse function
%>	var bReturn = false;

	var ErrorReturn = this.CheckResponse(null);
	bReturn = ErrorReturn[0];
	ErrorReturn = null;

	return bReturn;
}

function CheckResponse(sErrorTypes)
{
<%	/*	ErrorReturn is a one dimensional array, which is returned to the calling routine
		ErrorReturn[0] - A boolean indicating if the object call was successful
		ErrorReturn[1] - Returns the error type which failed, if that failure is listed in sErrorTypes.
		                 Otherwise it is null
		ErrorReturn[2] - The error description of the error type specified by ErrorReturn[1], or null

		Initialise the ErrorReturn array to a standard  failure
		Default message is "No Response Received"
		If an error occurs but no message found, message "Unspecified error" returned

		RF 03/12/99 Display error source

		Check to see if we are looking for specific errors as contained in sErrorTypes.  If the returned
		error matches one of these, set ErrorReturn[1] & [2]
		
		If there is an error and it isn't one we're looking for, display the message
		
		BG 12/09/00 SYS1277 added a 4th element to the Array (ErrorReturn[3]) in order to hold a 
		boolean for whether or not the response contains "WARNING".  I did this specifically for 
		use in DC270, but the functionality can be reused.  Rather than calling IsResponseOK to 
		determine whether or not the database call was successful, using CheckResponse you are able
		to determine if the call was successful AND if the response contains "WARNING" by checking the
		values of ErrorReturn[0] and ErrorReturn[3]in the calling function.		
	*/
%>	var ErrorReturn = new Array(4);

	ErrorReturn[0] = false;
	ErrorReturn[1] = null;
	ErrorReturn[2] = null;
	ErrorReturn[3] = false;

	this.ActiveTag = null;
	this.ActiveBeforeTag = null;

	var sErrorMessage = "No Response Received";
	var TagRESPONSE = this.SelectTag(null,"RESPONSE");

	if (TagRESPONSE != null)
	{
		if (this.GetAttribute("TYPE") == "SUCCESS")
			ErrorReturn[0] = true;
		else if (this.GetAttribute("TYPE") == "WARNING") 
		{
			<% 
			/* SYS0275 - Warning messages. At the moment WARNING and INFORMATION 
			messages are the only message types that are handled */
			%>
			ErrorReturn[0] = true;
			ErrorReturn[3] = true;
			var xmlMessageTagList = this.CreateTagList("MESSAGE");
			for (var iLoop=0;iLoop<xmlMessageTagList.length;iLoop++)
			{
				if (this.SelectTagListItem(iLoop))
				{
					sErrorMessage = this.GetTagText("MESSAGETEXT");
					var sMessageType = this.GetTagText("MESSAGETYPE");
					if (sMessageType.toUpperCase() == "WARNING") 
						alert("WARNING: " + sErrorMessage);
					else if (sMessageType.toUpperCase() == "INFORMATION")
						alert("INFORMATION: " + sErrorMessage);
				}
			}
			this.ActiveTag = TagRESPONSE;
		}	
		else
		{
			sErrorMessage = "Unspecified Error";

			if(this.SelectTag(TagRESPONSE,"ERROR") != null)
			{
				var nErrorNumber = this.GetTagText("NUMBER");
				sErrorMessage = this.GetTagText("DESCRIPTION");
				var sErrorSource = this.GetTagText("SOURCE");
				var sVersion = this.GetTagText("VERSION");

				if(sErrorTypes != null)
				{
					for(var nLoop = 0; nLoop < sErrorTypes.length && ErrorReturn[1] == null; nLoop++)
					{
						if(nErrorNumber == this.TranslateErrorType(sErrorTypes[nLoop]))
						{
							ErrorReturn[1] = sErrorTypes[nLoop];
							ErrorReturn[2] = sErrorMessage;
						}
					}
				}
			}
			<% // Added for User Validation - SYS5115 %>
 			else
			{
				this.ActiveTag = this.SelectTag(TagRESPONSE,"CLIENTVALIDATIONERROR");
				if(this.ActiveTag != null)
				{
					ErrorReturn[1] = "CLIENTVALIDATIONERROR";
				}
			}
		}
	}

	if(!ErrorReturn[0] && ErrorReturn[1] == null)
		alert(sErrorMessage + "\nSource: " + sErrorSource + "\nVersion: " + sVersion);

	return ErrorReturn;
}

function TranslateErrorType(sError)
{
	var iErrorNumber = -2147221504 + 512;

	switch (sError)
	{	
		case "NOTIMPLEMENTED":
			iErrorNumber += 900;
		break;
		
		case "RECORDNOTFOUND":
			iErrorNumber += 500;
		break;

		<% /*EP2_921 AS 16/01/2007 EP1288 START */ %>
		case "COMMANDFAILED":
			iErrorNumber += 505;
		break;
		case "TSQLSEVERITY16":
			iErrorNumber += 3092;
		break;
		<% /*EP2_921 AS 16/01/2007 EP1288 END */ %>

		case "ONEOFFCOSTS":
			iErrorNumber += 167;
		break;

		case "EXCEEDEDPPCOVERAMOUNT":
			iErrorNumber += 230;
		break;

		case "NOTAUTHORISED":
			iErrorNumber += 111;
		break;

		case "LOANCOMPONENTSNOTFOUND":
			iErrorNumber += 290;
		break;
		
<%		// TODO: Add further error types here...
%>		case "NOMORESTAGES":
			iErrorNumber += 4800;
		break;
		case "NOTASKAUTHORITY":
			iErrorNumber += 4807;
		break;
		case "NOSTAGEAUTHORITY":
			iErrorNumber += 4806;
		break;
		case "MANDTASKSOUTSTANDING":
			iErrorNumber += 4802;
		break;
		
		<%//GD BMIDS00572 START %>
		case "WRONGPASSWORD":
			iErrorNumber += 112;
		break;
		<%//GD BMIDS00572 END %>
		
		<% /*MO BMIDS01076 START */ %>
		case "AUTOTASKERROR":
			iErrorNumber += 7011;
		break;
		<% /*MO BMIDS01076 END */ %>
		
		<% /*JD BMIDS  */ %>
		case "CREDITCHECKOKIMPORTFAILED":
			iErrorNumber += 7005;
		break;
		
		default:
			iErrorNumber = 0;
		break;
	}

	return iErrorNumber;
}

<%	// IK_MAR1849 cached replacement %>	
function GetGlobalParameterBoolean(thisDocument,sParameterName)
{
	var xGlobal = getGlobalFromCache(sParameterName);
	if(xGlobal)
	{
		if(xGlobal.selectSingleNode("BOOLEAN"))
			return(xGlobal.selectSingleNode("BOOLEAN").text == 1);
		else
			return(false);
	}
		
	this.RunASPWithTextInput(thisDocument,"GetCurrentParameter.asp",sParameterName);

	if(this.IsResponseOK())
	{
		var xn = this.XMLDocument.selectSingleNode("RESPONSE/GLOBALPARAMETER");
		if(xn)
			addGlobalToCache(xn);
		if(xn.selectSingleNode("BOOLEAN"))
			return(xn.selectSingleNode("BOOLEAN").text == 1);
	}

	return false;
}

function GetGlobalParameterAmount(thisDocument, sParameterName)
{
	var xGlobal = getGlobalFromCache(sParameterName);
	if(xGlobal)
	{
		if(xGlobal.selectSingleNode("AMOUNT"))
			return xGlobal.selectSingleNode("AMOUNT").text;
		else
			return(0);
	}
	
	this.RunASPWithTextInput(thisDocument,"GetCurrentParameter.asp",sParameterName);
	
	if(this.IsResponseOK())
	{
		var xn = this.XMLDocument.selectSingleNode("RESPONSE/GLOBALPARAMETER");
		if(xn)
			addGlobalToCache(xn);
		if(xn.selectSingleNode("AMOUNT"))
			return xn.selectSingleNode("AMOUNT").text;
	}

	return "";
}
function GetGlobalParameterString(thisDocument, sParameterName)
{
	var xGlobal = getGlobalFromCache(sParameterName);
	if(xGlobal)
	{
		if(xGlobal.selectSingleNode("STRING"))
			return xGlobal.selectSingleNode("STRING").text;
		else
			return "";
	}
	
	this.RunASPWithTextInput(thisDocument,"GetCurrentParameter.asp",sParameterName);
	
	if(this.IsResponseOK())
	{
		var xn = this.XMLDocument.selectSingleNode("RESPONSE/GLOBALPARAMETER");
		if(xn)
			addGlobalToCache(xn);
		if(xn.selectSingleNode("STRING"))
			return xn.selectSingleNode("STRING").text;
	}

	return "";
}

function getGlobalFromCache(sParameterName)
{
	var xGlobalCache = getGlobalCache();
	if(xGlobalCache)
		return xGlobalCache.selectSingleNode("GLOBALS/GLOBALPARAMETER[NAME='" + sParameterName.toUpperCase() + "']");
}

function addGlobalToCache(xn)
{
	var xGlobalCache = getGlobalCache();
	if(xGlobalCache) 
	{
		var xnn = xGlobalCache.documentElement.appendChild(xn.cloneNode(true));
		xnn.selectSingleNode("NAME").text = xnn.selectSingleNode("NAME").text.toUpperCase();
	}
}

function getGlobalCache()
{
	var xGlobalCache = window.top.xGlobals;	<% /* MAR1849 GHun */ %>
	<% /* MAR1849 GHun If the cache is not found, it is probably due to the current window being a popup.
	 This relies on popups defining a global variable called m_BaseNonPopupWindow to reference the base window. */ %>
	if (!xGlobalCache)
	{
		try
		{
			xGlobalCache = window.top.m_BaseNonPopupWindow.top.xGlobals;
		}
		catch (e)
		{
			xGlobalCache = null;
		}
	}
	<% /* MAR1849 End */ %>
	
	return(xGlobalCache);
}
<%	// IK_MAR1849 ends %>	

function SelectNodes(sPattern)
{
	var TagListNew = null;

	if(this.ActiveTag != null)
		TagListNew = this.ActiveTag.selectNodes(sPattern);
	else
		TagListNew = this.XMLDocument.selectNodes(sPattern);

	this.ActiveTagList = TagListNew;

	return TagListNew;
}

function SelectSingleNode(sPattern)
{
	if(this.ActiveTag != null)
		this.ActiveTag = this.ActiveTag.selectSingleNode(sPattern);
	else
		this.ActiveTag = this.XMLDocument.selectSingleNode(sPattern);

	return this.ActiveTag;
}

function EncodeXML(xml)
{
	var returnXML="";
	if(xml.trim()!='')
	{
		//returnXML = xml;
		returnXML = xml.replace(/\£/g, "&#163;");
	}
	else
	{
		returnXML = xml;
	}
	
	return returnXML;
}

function GetComboIdsForValidation(sListName, sValidation, XMLCombo, thisDocument) // BMIDS774 KRW
{
<%	/*  Search the XML document for the group specified and, from this group return the ValueId with the given validation type.
		If multiple Ids have the same validation, return all Ids in the XML. 
		
		If XMLCombo(Input) is null, create the XML with the required combo values, before finding the ID, else
		use the Input XML. The input param thisDocument is required to be passed in for building the comboXML.
	*/
%>
	var sComboValueId;
	var arrComboValueIds = new Array(3);
	var iTotalArray = 0;

	var bBuildComboXML = false; 
	var XMLFragment;
	var XML = new XMLObject();
	
	if(arguments.length < 3)
		bBuildComboXML = true;
	else 
	{
		if(XMLCombo == null)
			bBuildComboXML = true;
	}
	
	if(bBuildComboXML)
	{	
		if(arguments.length < 4)
			return "";
		if(thisDocument == null)
			return "";
		
	 	if(XML.GetComboLists(thisDocument, Array(sListName)))
	 	{
	 		XMLFragment = XML.GetComboListXML(sListName);
	 		XML.ResetXMLDocument();
	 	} 
	}
	else
		XMLFragment = XMLCombo;

	XML.AddXMLBlock(XMLFragment);
	
	var TagListLISTENTRY = XML.CreateTagList("LISTENTRY");
	for(var nLoop = 0;nLoop < TagListLISTENTRY.length;nLoop++)
	{	
		XML.ActiveTagList = TagListLISTENTRY;
		XML.SelectTagListItem(nLoop);
		
		sComboValueId = XML.GetTagText("VALUEID");
		XML.CreateTagList("VALIDATIONTYPE");
		
		for(var nListLoop = 0; nListLoop < XML.ActiveTagList.length; nListLoop++)
		{
			if(XML.ActiveTagList.item(nListLoop).text == sValidation)
			{
				arrComboValueIds[iTotalArray] = sComboValueId;
				iTotalArray = iTotalArray + 1;
			}
		}
	}

	return arrComboValueIds;
}

function SetErrorResponse()
{
	this.ResetXMLDocument();
	this.CreateActiveTag("RESPONSE");
	this.SetAttribute("TYPE","CLIENTVALIDATIONERROR");
	this.CreateActiveTag("CLIENTVALIDATIONERROR");
}

</script>
