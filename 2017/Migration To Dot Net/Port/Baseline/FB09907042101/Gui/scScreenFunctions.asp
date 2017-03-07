<%@  Language=JScript %>
<html id="ScreenFunctions">
<%/*
Workfile:      scScreenFunctions.htm
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Generic screen functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		10/11/1999	Addition of EnableDrillDown and DisableDrillDown
AY		03/12/1999	Alteration to setting textarea fields read only
AY		28/01/2000	Optimised
AY		10/02/2000	IsOptionValidationType was not looping through all the validation types
AY		11/02/2000	SetFocusToFirst/LastField functionality altered to include msgButtons
AY		16/02/2000	SYS0240 - HEAD and TITLE elements removed
AD		24/02/2000	Added GetDateObjectFromString
AD		24/02/2000	Added DateToString
AY		25/02/2000	SetContextParameter enables/disables the navigation menu
AY		29/02/2000	Application number now displayed within modal_omigamenu
AD		20/03/2000	Made DoFocusProcessing public
AY		28/03/2000	SizeTextToField altered to work on labels
					scScreenFunctions can now be declared as an object
AY		04/04/2000	APPLICATION MENU title moved from FW020 to modal_OmigaMenu
AY		06/04/00	SYS0569 - New context field for stage information
AY		11/04/00	Spelling mistake in above fix
AY		11/04/00	SYS0328 - Dynamic currency display
AY		13/04/00	SYS0328 - Currency function required for popups
AY		12/04/00	Textarea read only problem fixed in IE5
MH      02/05/00    SYS0618 Add Postcode validation subroutine
MC		06/06/00	SYS0837 Restrict textarea's to 255 characters
JLD		14/11/00	Added time information to GetDateObjectFromString.
GD		03/01/01	Sys 0203 :Added 2 new fns LogonLegallyOpened and Omiga4LegallyOpened.
JLD		03/01/01	Added functions GetSystemDateTime, CompareDateStringToSystemDateTime
GD		10/01/01	Sys 0203 :Amendments to fns LogonLegallyOpened and Omiga4LegallyOpened
JLD		15/01/01	SYS1808 removed reference to context param idApplicationStage.
APS		03/03/01	SYS1920 Added new method called SetMainScreenToReadOnly()
JLD		08/03/01	SYS1879 added new method dateTimeToString()
APS		12/03/01	SYS1920 Added new method called IsMainScreenReadOnly()
APS		15/03/01	SYS2076
PSC		15/03/01	SYS2077 Amend DateTimeToString to left pad hours minutes and seconds with 0
IK		15/03/01	SYS1924 Completeness Check, call FW020 functionality on change of Stage Id.
IK		05/04/01	SYS2244 Add SyncCustomerIndex function
SR		24/05/01	SYS2325 - New context variable 'idOtherSystemAccountNumber'; set value for this 
					every time a value is set for 'idApplicationNumber'.
GD		20/07/01	SYS2509 SetOtherSystemACNoInContext changed, to only call GetApplicationDetails if app num exists
DS		16/11/01	OmiPLus26 Added function FormatAsCurrency.
JLD		4/12/01		SYS2806 - completeness check routing.
JLD		17/12/01	SYS3204 - correction to SetComboOnValidationType
DRC     03/05/02    SYS4530 Made SetOtherSystemACNoInContext method public
LD		23/05/02	SYS4727 Use cached versions of frame functions
SG		06/06/02	SYS4828 Corrected typo for SYS4727
MHeys	18/04/2006	MAR1587	Screen does not centre popup when called from browser cancel
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :
 
Prog	Date		AQR			Description  
MV		22/08/2002	BMIDS0355	IE 5.5 upgrade Errors - Modified GetOtherSystemAccountNumber()
								I suggest whenever builidng a request tag to call an ASP page 
								Please use the request tag format from the above function by specifying the frames name 
MV		14/10/2002	BMIDS00630	Added DisplayClientWarning() and DisplayClientError()
MV		30/10/2002	BMIDS00780	Amended CreateScreenFunctions() and ScreenFunctionsObject()
MO		13/11/2002	BMIDS00807	Created GetAppServerDate function
GHun	29/01/2003	BM0242		ValidatePostcode should support partial postcodes and ignore duplicate spaces
GD		10/06/2003	BM0356      Amended SetOtherSystemACNoInContext to cope with IE5/IE5.5 differences.
KRW     23/02/2004  BMIDS713    Changed DisplayPopup()No longer requires screen event indicator to display screens
MC		30/04/2004	BM0468		Cancel and Decline stage freeze screens functionality
MC		06/05/2004	BM0468		getCancelledStageValue() and getDeclinedStageValue() functions added.
GHun	29/07/2004	BMIDS821	Changed GetOtherSystemAccountNumber to GetApplicationData, and added SetOtherAppDetailsInContext
HMA     23/11/2004  BMIDS948    Add ApplicationType
JD		24/11/2004	BM0095		Changed SizeTextToField to put a dummy value into the field before measuring the start height.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History :
 
Prog	Date		AQR			Description  
SD		12/10/2005	MAR86		Changed DisplayClientWarning & DisplayClientError fuctions so that the MSGBox does not have a statusbar and scrollbar
PSC		13/10/2005	MAR57		Added customer categories to SetContextParameter
PE		21/02/2006	MAR1289		tm030 - long company name throws page layout
PJO     08/03/2006  MAR1359     Add Progress Message functionality
PJO     09/03/2006  MAR1359     Prositioning of Progress Message window changed
GHun	20/03/2006	MAR1453		Reduce progress message window height
GHun	05/04/2006	MAR1300		Change of property changes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

List the external function calls and their parameters here:
			
SetScreenToReadOnly(frmReference)
	frmReference	- the form object which contains all the fields you wish to make read only.  In the standard template this is frmScreen.

	Loops through the specified form and changes all fields to read only.  Does not alter buttons.
		

SetFieldState(frmReference, sFieldId, sInputState)
	frmReference	- the form object which contains the field you wish to make read only.  In the standard template this is frmScreen.
	sFieldId		- string specifying the id of the field you wish to set
	sInputState		- string specifying the state the field is to be set to: "R"eadonly, "W"riteable or "D"isabled

	Changes the specified field to the state specified.
	Does not operate on buttons.
	Do not use against a radio group. Use SetRadioGroupState instead.

		
SetRadioGroupState(frmReference, sGroupName, sInputState)
	frmReference	- the form object which contains the radio button group you wish to make read only.  In the standard template this is frmScreen.
	sGroupName		- string specifying the name of the group you wish to make read only
	sInputState		- string specifying the state the field is to be set to: "R"eadonly, "W"riteable or "D"isabled

	Changes all elements of the specified radio button group to the state specified.
	Use only against radio groups.
	Calls SetFieldState


A NOTE REGARDING THE FOLLOWING FOUR FUNCTIONS
Normally, the <SPAN> and <DIV> tags are not given ids during screen builds, so remember to assign ids.
Also, even when showing or hiding an individual field its better to use these functions on the field's surrounding
<SPAN> as they ensure the field is hidden/shown correctly.  They also avoid dealing with the associated labels.


SetCollectionState(refSpnOrDivId, sInputState)
	refSpnOrDivId	- the id of the span/div (or form) element to change
	sInputState		- string specifying the state the field is to be set to: "R"eadonly, "W"riteable or "D"isabled

	Changes all fields contained in a <SPAN> or <DIV> element to the state specified.
	(It may also be used on <FORM> elements)
	Calls SetFieldState


HideCollection(refSpnOrDivId)
	refSpnOrDivId	- the id of the span/div (or form) element to hide

	Hides a <SPAN> or <DIV> element.  (It can also work on a <FORM>)
	It also sets all fields contained within the element to "hidden" so that the mandatory
	processing doesn't flag any mandatory fields and removes the contents of the fields.
	Calls ClearCollection as part of its processing


ShowCollection(refSpnOrDivId)
	refSpnOrDivId	- the id of the span/div (or form) element to show
	Shows a <SPAN> or <DIV> element.  (It can also work on a <FORM>)
	It also sets all fields contained within the element to "visible" so that the mandatory
	processing will work.			


bIsChanged = ClearCollection(refSpnOrDivId)
	refSpnOrDivId	- the id of the span/div (or form) element to show
	bIsChanged		- returns true if a field's contents has been changed

	Clears the contents of all fields in a <SPAN> or <DIV> element.  (It can also work on a <FORM>)


vReturn = DisplayPopup(thisWindow, thisDocument, sPopup, sArguments, nPopupWidth, nPopupHeight)
	thisWindow		- the window object (usually window)
	thisDocument	- the document object (usually document)
	sPopup			- the filename of the popup screen
	sArguments		- the vArguments parameter for ShowModalDialog
	nPopupWidth		- the required width, in pixels, of the popup
	nPopupHeight	- the required height, in pixels, of the popup
	vReturn			- the value returned by ShowModalDialog

	Displays a centralised popup.


sValue = GetRadioGroupValue(frmReference, sGroupName)
	frmReference	- the form object which contains the radio group you wish to search.  In the standard template this is frmScreen.
	sGroupName		- the name of the group you wish to search			
	sValue			- the value of the radio button set, or null if no button set

	Searches the specified radio group for the button which is set and returns the value for that button.


SetRadioGroupValue(frmReference, sGroupName, sValue)
	frmReference	- the form object which contains the radio group you wish to set.  In the standard template this is frmScreen.
	sGroupName		- the name of the group you wish to set
	sValue			- the value to match against the radio buttons

	Searches the specified radio group for the button whose value matches sValue.
	Radio buttons which do not match are unchecked and the radio button which matches is checked.


SetCheckBoxValue(frmReference, sFieldId, sValue)						
	frmReference	- the form object which contains the checkbox you wish to set.  In the standard template this is frmScreen.
	sFieldId		- string specifying the field id of the check box you wish to set
	sValue			- the value to determine whether or not the checkbox is checked

	Sets the checked box specified by the field id and the value passed in.


sValue = GetCheckBoxValue(frmReference, sFieldId)						
	frmReference	- the form object which contains the checkbox you wish to set.  In the standard template this is frmScreen.
	sFieldId		- string specifying the field id of the check box you wish to set
	sValue			- the value returned indicating if the checkbox is checked

	Gets the checked box value specified by the field id


CopyComboList(refFieldFrom, refFieldTo)
	refFieldFrom	- the combo field to copy from e.g. frmScreen.cboFrom
	refFieldTo		- the combo field to copy to e.g. frmScreen.cboTo

	Copies the options contained in one combo to another combo


bReturn = IsValidationType(refField, sValue)
	refField	- the combo field to check
	sValue		- the value to look for
	bReturn		- true if successful, otherwise false

	Searches the selected combo option for a ValidationType attribute whose value matches sValue
	(see also scXMLFunctions - GetComboList)


bReturn = IsOptionValidationType(refField, nIndex, sValue)
	refField	- the combo field to check
	nIndex		- the option to check
	sValue		- the value to look for
	bReturn		- true if successful, otherwise false

	Searches the combo option(nIndex) for a ValidationType attribute whose value matches sValue
	(see also scXMLFunctions - GetComboList)


bReturn = IsOmigaMenuFrame(thisWindow)
	thisWindow	- the window object (usually window)
	bReturn		- true if found, otherwise false

	Checks whether the parent of thisWindow has a frame called "omigamenu".  This should be present
	if an omiga screen is being run within the framework.


sValue = GetNameForStageNumber(thisWindow,sStageNumber)

	*** SYS1808 METHOD REMOVED AS APPLICATIONSTAGE COMBO NO LONGER EXISTS ***
	
	thisWindow		- the window object (usually window)
	sStageNumber	- The stage number for which the name is required
	sValue			- the returned name

	If the omigamenu frame exists then the name of the specified stage number is retrieved from
	the ApplicationStage context field.
	If the omigamenu frame is not present then the user is prompted to input a value.


sValue = GetContextParameter(thisWindow,sContextParameterId,sDefaultValue)
	thisWindow			- the window object (usually window)
	sContextParameterId	- the id of the context field you wish to read
	sDefaultValue		- the default value for the prompt if omigamenu is not present
	sValue				- the value returned from the context field, or from the prompt

	If the omigamenu frame exists then the value of the specified context parameter is obtained.
	If the omigamenu frame is not present then the user is prompted to input a value.  The prompt
	is defaulted to the value specified.
	sDefaultValue may be set to null.


SetContextParameter(thisWindow,sContextParameterId,sValue)
	thisWindow			- the window object (usually window)
	sContextParameterId	- the id of the context field you wish to set
	sValue				- the value you wish to set the context field to

	If the omigamenu frame exists then the value of the specified context parameter is set.
	sValue may be set to null.


SetFocusToFirstField(frmReference)
	frmReference	- the form object which contains the fields you wish to check.  In the standard template this is frmScreen.

	Sets the focus to the first field or button on a screen which is enabled.


SetFocusToLastField(frmReference)
	frmReference	- the form object which contains the fields you wish to check.  In the standard template this is frmScreen.

	Sets the focus to the last field or button on a screen which is enabled.


SizeTextToField(refField,sValue)
	refField	- the field to set
	sValue		- the value you wish to set the field to 

	Currently works on TD and LABEL fields.
	Places the specified text into the field. If it is too big for the field it shortens it to fit
	and adds ... to the end.
	If the full text fits into the field then the title attribute is not set, otherwise it is set to
	the full text.

bIsValid = CompareDateFieldToSystemDate(refDateField,sComparison)
	refDateField	- the id of the field to check. Must be a date field.
	sComparison		- the comparison to make. It must be "<", "<=", "=", ">=" or ">".
	bIsValid		- returns true if the check meets the comparison criteria or false if it fails
					  or is blank.  Therefore it is best to check for the failure criteria rather
					  than the success. e.g. if the date has to be in the past do a ">" check and if
					  this is true then the date is invalid.

	Gets the contents of a date field and checks it against the system date, using the comparison specified.

bIsValid = CompareDateStringToSystemDate(sDate,sComparison)
	sDate		 - A string containing a date.
	sComparision - the comparison to make. It must be "<", "<=", "=", ">=" or ">".

	As for CompareDateFieldToSystemDate except a string containing a date is passed in rather than a reference to
	an onscreen field.
	
bIsValid = CompareDateStringToSystemDateTime(sDate,sComparison)
	sDate		 - A string containing a date.
	sComparision - the comparison to make. It must be "<", "<=", "=", ">=" or ">".

	As for CompareDateFieldToSystemDate except a string containing a date is passed in rather than a reference to
	an onscreen field and the system date is compared with time information also.

sString = DateToString(dtDate)
	dtDate - A date object
	
	Returns the passed date as a string in the format DD/MM/YYYY. If any one can think of a better way of doing this
	then please feel free to change this!

sString = DateTimeToString(dtDate)
	dtDate - A date object
	
	Returns the passed date as a string in the format 'DD/MM/YYYY HH:MM:SS'.
	
bIsValid = CompareDateFields(refFirstDateField,sComparison,refSecondDateField)
	refFirstDateField	- the id of the first field to check. Must be a date field.
	sComparison			- the comparison to make. It must be "<", "<=", "=", ">=" or ">".
						  e.g. ">" will check whether the first date is greater than the second
	refSecondDateField	- the id of the second field to check. Must be a date field.
	bIsValid			- returns true if the check meets the comparison criteria or false if it fails
						  or one or both fields are blank.  Therefore it is best to check for the failure criteria
						  rather than the success. e.g. if the first date has to be earlier than the second
						  date do a ">" check and if this is true then one of the dates is invalid.

	Gets the contents of two date field and compares them, using the comparison specified.


dtFieldDate = GetDateObject(refDateField)
	refDateField	- the id of the field to convert. Must be a date field.
	dtFieldDate		- the date object.  If the field was blank or an invalid format then this is null.

	Converts the contents of a date field into a date object.

dtFieldDate = GetDateObjectFromString(sDate)
	sDate - the string to convert to a date object

	Converts a string containing a date to a date object.

sValidationType = GetComboValidationType(frmReference, sFieldId, iIndex)
	frmReference	- the form object which contains the fields you wish to check.  In the standard template this is frmScreen.
	sFieldId		- The id of the combo field to set
	iIndex			- optional parameter specifiying the combo index to use, if this parameter is not
					  present the method will use the selectedIndex property.

	Gets the currently selected validation record for the combo specified
	This function will only get the first validation record.

bReturnToCompletenessCheck = CompletenessCheckRouting(thisWindow)
	thisWindow	 - current window object
	
	Checks the context parameter idProcessContext and if it is "CompletenessCheck" then returns TRUE.
	Calling screen should then route to COmpleteness check screen (GN300).
	This is standard routing procedure for all screens.
	
SetComboOnValidationType(frmReference, sFieldId, sValue)
	frmReference	- the form object which contains the fields you wish to check.  In the standard template this is frmScreen.
	sFieldId		- The id of the combo field to set
	sValue			- The value to match the validation record against

	Sets a combo selection based on the validation record (sValue)


EnableDrillDown(refField)
	refField	- the field to enable

	Enables a drilldown button, changing the classname to show the difference visually.
	Will only work on a button whose classname has been set to msgDisabledDDButton.


DisableDrillDown(refField)
	refField	- the field to disable

	Disables a drilldown button, changing the classname to show the difference visually.
	Will only work on a button whose classname has been set to msgDDButton.

DoFocusProcessing(frmReference, nElement)
	frmReference	- a handle to the form containing the element to set focus to
	nElement		- the index of the element within the form to set focus to

	Sets focus to the specified element, if it is valid to do so.


sCurrency = SetCurrency(thisWindow,frmReference)
	thisWindow		- the window object (usually window)
	frmReference	- the form object which contains the fields you wish to populate.  In the standard template this is frmScreen.
	sCurrency		- the character for the currency required
	
	This function is called from main screens only.
	Gets the currency contect parameter and passes this into the SetThisCurrency function.
	See description below for details


sCurrency = SetThisCurrency(frmReference,sCurrencyParameter)
	frmReference		- the form object which contains the fields you wish to populate.  In the standard template this is frmScreen.
	sCurrencyParameter	- the string specifying the currency required
	sCurrency			- the character for the currency required
	
	This function is called from popups and is called by SetCurrency.
	Checks the setting of the sCurrencyParameter variable, then runs through the frmReference collection,
	setting any labels of the msgCurrency class to the appropriate currency character.  Additionally it
	returns this character.
	If sCurrencyParameter is not explicitly set then £ is defaulted
	
bIsValid = ValidatePostcode(refField)
    refField - the Field form object that supports the value property]
    
    This function validates the field value against all known UK postcode templates.
    Empty fields are ignored. Trailing white space is ignored.
    If the value of refField matches one of the masks then True is returned. If it
    does not match any mask then false is returned, an alert is raised and the field 
    is given focus.
    
bTrimmed = RestrictLength(frmReference, sFieldId, nMaxLength, bMessage)
	frmReference	- the form object which contains the fields you wish to populate.  In the standard template this is frmScreen.
	sFieldId		- The id of the textarea field to set
	nMaxLength		- The maximum length of text that the textarea can contain
	bMessage		- true if warning message to be displayed
	
	This function restricts the amount of text entered in a textarea field (textarea's do not support
	the maxlength attribute as input fields do). Returns true if the text was truncated or false otherwise.


LogonLegallyOpened
	Used to eradicate the possibility of jumping to pages by typing the URL into
	the address bar, or using the browser's History functionality, for pages before the second
	instance of Internet Explorer is opened


Omiga4LegallyOpened(blnFirstScreen)
blnFirstScreen - a boolean flag set to false when called from the first screen
in the application (currently mn010.asp - 10/01/01), and true otherwise.
The reason for this flag is that when mn010 loads there is nothing currently in context
and therefore you do not want to check it.
This means that a user could reach this page illegally if they built another page to
access it and had all frames and forms in place and named correctly. But as soon as they tried to go to any other page,
then it would fail, because for all other pages the context is checked.

	Used to eradicate the possibility of jumping to pages by typing the URL into
	the address bar, or using the browser's History functionality, for pages within the second
	instance of Internet Explorer

blnReadOnly = SetMainScreenToReadOnly(thisWindow, frmReference, sScreenId)
	frmReference	- the form object which contains the fields.  In the standard template this is frmScreen.
	thisWindow		- the window object (usually window)
	sScreenId		- Screen Identification
	blnReadOnly		- Either true including setting the screen to read only or false incdicating the screen is writable
					
	This method determines whether or not the screen input should be in read-only mode by checking
	for ReadOnly, ProcessingUnit and DataFreezing context parameters being set. 
	And then Sets the screen to READ ONLY

blnReadOnly = IsMainScreenReadOnly(thisWindow, sScreenId)	
	thisWindow		- the window object (usually window)
	sScreenId		- Screen Identification
	blnReadOnly		- Either true including setting the screen to read only or false incdicating the screen is writable
					
	This method determines whether or not the screen input should be in read-only mode by checking
	for ReadOnly, ProcessingUnit and DataFreezing context parameters being set.
	
nIndex = SyncCustomerIndex(thisWindow)
	thisWindow	- the window object (usually window)
	sets Context Parameter idCustomerIndex, based on current idCustomerNumber 
	& returns idCustomerIndex value
	returns -1 if idCustomerIndex is not set 


sString = FormatAsCurreny(value)
	value must be preformatted using the math function RoundValue(value,2) before being passed in.
	returns the value with commas addedd appropriate for a currency value
	
PJO 08/03/2006 - Added Progress Message Functionality
progressOn(sText, iWidth)
	Opens a Progress Message Window with the revolving Ing logo and a message
	sText is the message to be displayed.
	iWidth is the width of the box which may thus be changed depending on the message

progressOff()
	Closes the Progress Message window
*/
%>
<object data=scXMLFunctions.asp height=1 id=scXMLFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 viewastext></OBJECT>

<script language="JavaScript" type="text/javascript">

public_description = new CreateScreenFunctions;

function CreateScreenFunctions()
{
	this.ScreenFunctionsObject = ScreenFunctionsObject;

	this.SetScreenToReadOnly = SetScreenToReadOnly;
	this.SetFieldState = SetFieldState;
	this.SetRadioGroupState = SetRadioGroupState;
	this.SetCollectionState = SetCollectionState;
	this.HideCollection = HideCollection;
	this.ShowCollection = ShowCollection;
	this.ClearCollection = ClearCollection;
	this.DisplayPopup = DisplayPopup;
	this.GetRadioGroupValue = GetRadioGroupValue;
	this.SetRadioGroupValue = SetRadioGroupValue;
	this.SetCheckBoxValue = SetCheckBoxValue;
	this.GetCheckBoxValue = GetCheckBoxValue;
	this.CopyComboList = CopyComboList;
	this.IsValidationType = IsValidationType;
	this.IsOptionValidationType = IsOptionValidationType;
	this.IsOmigaMenuFrame = IsOmigaMenuFrame;
	this.IsCancelDeclineStage = IsCancelDeclineStage;
	this.GetContextParameter = GetContextParameter;
	this.SetContextParameter = SetContextParameter;
	this.SyncCustomerIndex = SyncCustomerIndex;
	this.GetContextCustomerName = GetContextCustomerName;
	this.SetFocusToFirstField = SetFocusToFirstField;
	this.SetFocusToLastField = SetFocusToLastField;
	this.SizeTextToField = SizeTextToField;
	this.CompareDateFieldToSystemDate = CompareDateFieldToSystemDate;
	this.CompareDateStringToSystemDateTime = CompareDateStringToSystemDateTime;
	this.CompareDateFields = CompareDateFields;
	this.DateToString = DateToString;
	this.DateTimeToString = DateTimeToString;
	this.GetDateObject = GetDateObject;
	this.GetDateObjectFromString = GetDateObjectFromString;
	this.GetSystemDate = GetSystemDate;
	this.GetSystemDateTime = GetSystemDateTime;
	this.GetComboValidationType = GetComboValidationType;
	this.SetComboOnValidationType = SetComboOnValidationType;
	this.EnableDrillDown = EnableDrillDown;
	this.DisableDrillDown = DisableDrillDown;
	this.DoFocusProcessing = DoFocusProcessing;
	this.SetCurrency = SetCurrency;
	this.SetThisCurrency = SetThisCurrency;
	this.ValidatePostcode = ValidatePostcode;
	this.RestrictLength = RestrictLength;
	this.getCancelledStageValue = getCancelledStageValue;
	this.getDeclinedStageValue =  getDeclinedStageValue;

	this.LogonLegallyOpened = LogonLegallyOpened;
	this.Omiga4LegallyOpened = Omiga4LegallyOpened;
	
<% /* SYS 4530 */ %>
	this.SetOtherSystemACNoInContext = SetOtherSystemACNoInContext;	
	
<%	/* The following functions should no longer be called */ %>
	this.SetFieldToReadOnly = SetFieldToReadOnly;
	this.SetFieldToWritable = SetFieldToWritable;
	this.SetFieldToDisabled = SetFieldToDisabled;
	this.SetRadioGroupToReadOnly = SetRadioGroupToReadOnly;
	this.SetRadioGroupToWritable = SetRadioGroupToWritable;
	this.SetRadioGroupToDisabled = SetRadioGroupToDisabled;
	this.EnableCollection = EnableCollection;
	this.DisableCollection = DisableCollection;
	this.SetCollectionToReadOnly = SetCollectionToReadOnly;
	this.SetCollectionToWritable = SetCollectionToWritable;
	this.FormatAsCurrency = FormatAsCurrency;
	this.CompletenessCheckRouting = CompletenessCheckRouting;
	this.DisplayClientError = DisplayClientError;
	this.DisplayClientWarning = DisplayClientWarning;
	
	this.IsCancelDeclineStage = IsCancelDeclineStage;
	this.IsDataFreezeScreen=IsDataFreezeScreen;
<% /* PJO 08/03/2006 MAR1359 */ %>
    this.progressOn=progressOn;	
    this.progressOff=progressOff;	
}

function ScreenFunctionsObject()
{
	this.SetScreenToReadOnly = SetScreenToReadOnly;
	this.SetFieldState = SetFieldState;
	this.SetRadioGroupState = SetRadioGroupState;
	this.SetCollectionState = SetCollectionState;
	this.HideCollection = HideCollection;
	this.ShowCollection = ShowCollection;
	this.ClearCollection = ClearCollection;
	this.DisplayPopup = DisplayPopup;
	this.GetRadioGroupValue = GetRadioGroupValue;
	this.SetRadioGroupValue = SetRadioGroupValue;
	this.SetCheckBoxValue = SetCheckBoxValue;
	this.GetCheckBoxValue = GetCheckBoxValue;
	this.CopyComboList = CopyComboList;
	this.IsValidationType = IsValidationType;
	this.IsOptionValidationType = IsOptionValidationType;
	this.IsOmigaMenuFrame = IsOmigaMenuFrame;
	
	this.GetContextParameter = GetContextParameter;
	this.SetContextParameter = SetContextParameter;
	this.SyncCustomerIndex = SyncCustomerIndex;
	this.GetContextCustomerName = GetContextCustomerName;
	this.SetFocusToFirstField = SetFocusToFirstField;
	this.SetFocusToLastField = SetFocusToLastField;
	this.SizeTextToField = SizeTextToField;
	this.CompareDateFieldToSystemDate = CompareDateFieldToSystemDate;
	this.CompareDateStringToSystemDate = CompareDateStringToSystemDate;
	this.CompareDateStringToSystemDateTime = CompareDateStringToSystemDateTime;
	this.CompareDateFields = CompareDateFields;
	this.DateToString = DateToString;
	this.DateTimeToString = DateTimeToString;
	this.GetDateObject = GetDateObject;
	this.GetDateObjectFromString = GetDateObjectFromString;
	this.GetSystemDate = GetSystemDate;
	this.GetSystemDateTime = GetSystemDateTime;
	this.GetComboValidationType = GetComboValidationType;
	this.SetComboOnValidationType = SetComboOnValidationType;
	this.EnableDrillDown = EnableDrillDown;
	this.DisableDrillDown = DisableDrillDown;
	this.DoFocusProcessing = DoFocusProcessing;
	this.SetCurrency = SetCurrency;
	this.SetThisCurrency = SetThisCurrency;
	this.ValidatePostcode = ValidatePostcode;
	this.RestrictLength = RestrictLength;

	this.LogonLegallyOpened = LogonLegallyOpened;
	this.Omiga4LegallyOpened = Omiga4LegallyOpened;

<% // APS SYS1920 Added new method
%>
	this.SetMainScreenToReadOnly = SetMainScreenToReadOnly;
	this.IsMainScreenReadOnly = IsMainScreenReadOnly;
<% // DRC SYS4530 Made SetOtherSystemACNoInContext method public
%>	
	this.SetOtherSystemACNoInContext = SetOtherSystemACNoInContext;

	this.IsCancelDeclineStage = IsCancelDeclineStage;
	this.getCancelledStageValue = getCancelledStageValue;
	this.getDeclinedStageValue =  getDeclinedStageValue;
	this.IsDataFreezeScreen = IsDataFreezeScreen;
<% /* PJO 08/03/2006 MAR1359 */ %>
    this.progressOn=progressOn;	
    this.progressOff=progressOff;	
	

<%	/* The following functions should no longer be called */
%>	this.SetFieldToReadOnly = SetFieldToReadOnly;
	this.SetFieldToWritable = SetFieldToWritable;
	this.SetFieldToDisabled = SetFieldToDisabled;
	this.SetRadioGroupToReadOnly = SetRadioGroupToReadOnly;
	this.SetRadioGroupToWritable = SetRadioGroupToWritable;
	this.SetRadioGroupToDisabled = SetRadioGroupToDisabled;
	this.EnableCollection = EnableCollection;
	this.DisableCollection = DisableCollection;
	this.SetCollectionToReadOnly = SetCollectionToReadOnly;
	this.SetCollectionToWritable = SetCollectionToWritable;
	this.FormatAsCurrency = FormatAsCurrency;
	this.CompletenessCheckRouting = CompletenessCheckRouting;
	this.DisplayClientError = DisplayClientError;
	this.DisplayClientWarning = DisplayClientWarning;
	
	<% /* MO 13/11/2002 BMIDS00807 : SYSMCP1180 - Get date / dateTime from Application Server */ %>
	this.GetAppServerDate = GetAppServerDate ; 
	<% /* MO 13/11/2002 BMIDS00807 : SYSMCP1180 - End */ %> 
	
}

function SetScreenToReadOnly(frmReference)
{
	for(var nLoop = 0;nLoop < frmReference.elements.length;nLoop++)
		SetFieldState(frmReference, frmReference.elements(nLoop).id, "R");
}

function SetFieldState(frmReference, sFieldId, sInputState)
{
<%	/*	Some general notes:
		If a field is mandatory then when setting that field to readonly or disabled the field background
		colour must be cleared.
		When setting fields to disabled any text/settings must be cleared
	*/
%>	var sState = sInputState.toUpperCase();
	if(sState != "W" && sState != "R" && sState != "D")
	{
		alert("SetFieldState - Invalid State");
		return;
	}

	var thisAll = frmReference.all(sFieldId);

	if(thisAll.tagName == "INPUT")
	{
<%		/*	Text fields
			Because text fields have different classes, some extra processing is required. When being set to
			read only or disabled the attribute TxtClass is set to the current class setting, but only if
			TxtClass has not already been set (this avoids it being set incorrectly). When the field isbeing
			set to writable then the class is set to this attribute.
		*/
%>		if(thisAll.type != "button" && thisAll.type != "radio" && thisAll.type != "checkbox")
		{
			var sTxtClass = thisAll.style.getAttribute("TxtClass");

			if(sState == "W")
			{
				thisAll.readOnly = false;
				thisAll.disabled = false;
				thisAll.tabIndex = 0;
				if(sTxtClass != null) thisAll.className = sTxtClass;
			}
			else
			{
				thisAll.tabIndex = -1;
				if(sTxtClass == null) thisAll.style.setAttribute("TxtClass", thisAll.className);
				thisAll.style.backgroundColor = "";
				thisAll.className = "msgReadOnly";

				if(sState == "R") thisAll.readOnly = true;
				else
				{
					thisAll.disabled = true;
					thisAll.value = "";
				}
			}
		}
<%		/*	Radio button fields
			If a radio button is disabled, the label is coloured as disabled.
			If a radio button is read only, the label is only coloured as disabled if the
			radio button is not set
			For setting a radio button to writable the label has to be reset
		*/
%>		if(thisAll.type == "radio")
		{
			if(sState == "W")
			{
				thisAll.disabled = false;
				thisAll.tabIndex = 0;
			}
			else
			{
				thisAll.disabled = true;
				thisAll.tabIndex = -1;
				if(sState == "D") thisAll.checked = false;
			}

			if((sState == "R" && !thisAll.checked) || sState != "R")
			{
				for(var nLoop = 0;nLoop < frmReference.all.length;nLoop++)
				{
					if(frmReference.all(nLoop).tagName == "LABEL")
					{
						if(thisAll.id == frmReference.all(nLoop).htmlFor)
						{
							if(sState == "W") frmReference.all(nLoop).className="msgLabel";
							else frmReference.all(nLoop).className="msgReadOnlyLabel";
						}
					}
				}
			}
		}
<%		/*	Checkbox fields
			If a checkbox is disabled, the label is coloured as disabled.
			If a checkbox is read only, the label is only coloured as disabled if the
			checkbox is not set
			For setting a checkbox to writable the label has to be reset
		*/
%>		if(thisAll.type == "checkbox")
		{
			if(sState == "W")
			{
				thisAll.disabled = false;
				thisAll.tabIndex = 0;
			}
			else
			{
				thisAll.disabled = true;
				thisAll.tabIndex = -1;
				if(sState == "D") thisAll.checked = false;
			}

			if((sState == "R" && !thisAll.checked) || sState != "R")
			{
				for(var nLoop = 0;nLoop < frmReference.all.length;nLoop++)
				{
					if(frmReference.all(nLoop).tagName == "LABEL")
					{
						if(thisAll.id == frmReference.all(nLoop).htmlFor)
						{
							if(sState == "W") frmReference.all(nLoop).className="msgLabel";
							else frmReference.all(nLoop).className="msgReadOnlyLabel";
						}
					}
				}
			}
		}
	}
<%	/*	Textarea fields
		AY 03/12/99 - If the text field is just read only, selecting it and typing will clear the class
		settings, returning it to default colours.  Setting the disabled attribute as well solves this.
		AY 12/04/00 - This now appears to be fixed in IE5
	*/
%>	if(thisAll.tagName == "TEXTAREA")
	{
		if(sState == "W")
		{
			thisAll.readOnly = false;
			thisAll.disabled = false;
			thisAll.tabIndex = 0;
			thisAll.className = "msgTxt";
		}
		else
		{
			thisAll.tabIndex = -1;
			thisAll.style.backgroundColor = "";
			thisAll.className = "msgReadOnly";

			if(sState == "R") thisAll.readOnly = true;
			else
			{
				thisAll.disabled = true;
				thisAll.value = "";
			}
		}
	}
<%	/*	Combo fields
		To apply a change of class to a select type field we have to hide and redisplay it
	*/
%>	if(thisAll.tagName == "SELECT")
	{
		if(sState == "W")
		{
			thisAll.disabled = false;
			thisAll.tabIndex = 0;
			thisAll.className = "msgCombo";
			thisAll.style.visibility = "hidden";
			thisAll.style.visibility = "visible";

			if(thisAll.selectedIndex == -1) thisAll.selectedIndex = 0;
		}
		else
		{
			thisAll.disabled = true;
			thisAll.tabIndex = -1;
			thisAll.style.backgroundColor = "";
			thisAll.className = "msgReadOnlyCombo";
			thisAll.style.visibility = "hidden";
			thisAll.style.visibility = "visible";

			if(sState == "D") thisAll.selectedIndex = -1;
		}
	}
}

function SetRadioGroupState(frmReference, sGroupName, sInputState)
{
	for(var nLoop = 0;nLoop < frmReference.elements.length;nLoop++)
	{
		if(frmReference.elements(nLoop).name == sGroupName)
			SetFieldState(frmReference, frmReference.elements(nLoop).id,sInputState);
	}
}

function SetCollectionState(refSpnOrDivId,sInputState)
{
	for(var nLoop = 0;nLoop < refSpnOrDivId.all.length;nLoop++)
	{
		if (refSpnOrDivId.all(nLoop).id != "")
			SetFieldState(refSpnOrDivId, refSpnOrDivId.all(nLoop).id,sInputState);
	}
}

function HideCollection(refSpnOrDivId)
{
<%	/*	General notes
		When hiding a field, also clear its contents and ensure that if it has been flagged as
		mandatory the colour is cleared
	*/
%>	refSpnOrDivId.style.visibility = "hidden";
	ClearCollection(refSpnOrDivId);

	for(var nLoop = 0;nLoop < refSpnOrDivId.all.length;nLoop++)
	{
		if(refSpnOrDivId.all(nLoop).tagName == "INPUT")
		{
			refSpnOrDivId.all(nLoop).style.visibility = "hidden";
<%			//Text fields
%>			if(refSpnOrDivId.all(nLoop).type != "button" && refSpnOrDivId.all(nLoop).type != "radio" && refSpnOrDivId.all(nLoop).type != "checkbox")
				refSpnOrDivId.all(nLoop).style.backgroundColor = "";
		}
<%		// Textarea fields
%>		if(refSpnOrDivId.all(nLoop).tagName == "TEXTAREA")
		{
			refSpnOrDivId.all(nLoop).style.visibility = "hidden";
			refSpnOrDivId.all(nLoop).style.backgroundColor = "";
		}
<%		// Combo fields
%>		if(refSpnOrDivId.all(nLoop).tagName == "SELECT")
		{
			refSpnOrDivId.all(nLoop).style.visibility = "hidden";
			refSpnOrDivId.all(nLoop).style.backgroundColor = "";
		}
	}
}

function ShowCollection(refSpnOrDivId)
{
	<%		// APS 10/02/2000 - Added table to list of controls to make visible
	%>
	refSpnOrDivId.style.visibility = "visible";

	for(var nLoop = 0;nLoop < refSpnOrDivId.all.length;nLoop++)
	{
		if(refSpnOrDivId.all(nLoop).tagName == "INPUT") refSpnOrDivId.all(nLoop).style.visibility = "visible";
		if(refSpnOrDivId.all(nLoop).tagName == "TEXTAREA") refSpnOrDivId.all(nLoop).style.visibility = "visible";
		if(refSpnOrDivId.all(nLoop).tagName == "SELECT") refSpnOrDivId.all(nLoop).style.visibility = "visible";
		if(refSpnOrDivId.all(nLoop).tagName == "TABLE") refSpnOrDivId.all(nLoop).style.visibility = "visible";
	}
}

function ClearCollection(refSpnOrDivId)
{
	var bIsChanged = false;

	for(var nLoop = 0;nLoop < refSpnOrDivId.all.length;nLoop++)
	{
		if(refSpnOrDivId.all(nLoop).tagName == "INPUT")
		{
<%			//Text fields
%>			if(refSpnOrDivId.all(nLoop).type != "button" && refSpnOrDivId.all(nLoop).type != "radio" && refSpnOrDivId.all(nLoop).type != "checkbox")
			{
				if(refSpnOrDivId.all(nLoop).value != "")
				{
					bIsChanged = true;
					refSpnOrDivId.all(nLoop).value = "";
				}
			}
<%			/*	Radio buttons
				Only bother if it is actually checked
				If the radio button status is disabled then the label has to be set to disabled
			*/
%>			if(refSpnOrDivId.all(nLoop).type == "radio")
			{
				if(refSpnOrDivId.all(nLoop).checked)
				{
					bIsChanged = true;
					refSpnOrDivId.all(nLoop).checked = false;

					if(refSpnOrDivId.all(nLoop).disabled)
					{
						for(var nButtonLoop = 0;nButtonLoop < refSpnOrDivId.all.length;nButtonLoop++)
						{
							if(refSpnOrDivId.all(nButtonLoop).tagName == "LABEL")
							{
								if(refSpnOrDivId.all(nLoop).id == refSpnOrDivId.all(nButtonLoop).htmlFor)
									refSpnOrDivId.all(nButtonLoop).className="msgReadOnlyLabel";
							}
						}
					}
				}
			}
<%			/*	Check boxes
				Only bother if it is actually checked
				If the check box status is disabled then the label has to be set to disabled
			*/
%>			if(refSpnOrDivId.all(nLoop).type == "checkbox")
			{
				if(refSpnOrDivId.all(nLoop).checked)
				{
					bIsChanged = true;
					refSpnOrDivId.all(nLoop).checked = false;

					if(refSpnOrDivId.all(nLoop).disabled)
					{
						for(nCheckboxLoop = 0;nCheckboxLoop < refSpnOrDivId.all.length;nCheckboxLoop++)
						{
							if(refSpnOrDivId.all(nCheckboxLoop).tagName == "LABEL")
							{
								if(refSpnOrDivId.all(nLoop).id == refSpnOrDivId.all(nCheckboxLoop).htmlFor)
									refSpnOrDivId.all(nCheckboxLoop).className="msgReadOnlyLabel";
							}
						}
					}
				}
			}
		}
<%		//Textarea fields
%>		if(refSpnOrDivId.all(nLoop).tagName == "TEXTAREA")
		{
			if(refSpnOrDivId.all(nLoop).value != "")
			{
				bIsChanged = true;
				refSpnOrDivId.all(nLoop).value = "";
			}
		}
<%		// Combo fields
		// If the combo is disabled, dont reset it to option 1
%>		if(refSpnOrDivId.all(nLoop).tagName == "SELECT")
		{
			if(refSpnOrDivId.all(nLoop).selectedIndex != -1)
			{
				if(refSpnOrDivId.all(nLoop).selectedIndex != 0)
				{
					bIsChanged = true;
					refSpnOrDivId.all(nLoop).selectedIndex = 0;
				}
			}
		}
	}

	return bIsChanged;
}

function DisplayPopup(thisWindow, thisDocument, sPopup, sInput, nPopupWidth, nPopupHeight)
{
<%	/*	Calculate the top left of the main window.  This requires a window event to be captured in order to work
		(Such as hitting a button which displays a popup)
		Then, for a screen with all the framework, we need to work out the top left by a more circuitous route
		We know there are frames to the left and top of of the calling screen, so if we work out the
		difference between the dimensions of the parent window and the calling window we can calculate
		the top left of the parent window
		Once this is done we can then populate the top left co-ordinates for the popup
	*/
%>	//var nWindowTop = thisWindow.event.screenY - thisWindow.event.clientY;
	//var nWindowLeft = thisWindow.event.screenX - thisWindow.event.clientX;
	
	var nWindowTop = thisWindow.screenTop;   //KRW     23/02/2004  BMIDS713
	var nWindowLeft = thisWindow.screenLeft; //KRW     23/02/2004  BMIDS713
	
	var nWindowWidth;
	var nWindowHeight;
	
	if(IsOmigaMenuFrame(thisWindow))
	{
		nWindowWidth = thisWindow.parent.document.body.clientWidth;
		nWindowHeight = thisWindow.parent.document.body.clientHeight;
		nWindowTop = nWindowTop - (nWindowHeight - thisDocument.body.clientHeight);
		nWindowLeft = nWindowLeft - (nWindowWidth - thisDocument.body.clientWidth);
	}
	else
	{
		nWindowWidth = thisDocument.body.clientWidth;
		nWindowHeight = thisDocument.body.clientHeight;
	}
	
	var nPopupTop = Math.floor((nWindowHeight - nPopupHeight) / 2) + nWindowTop;
	var nPopupLeft = Math.floor((nWindowWidth - nPopupWidth) / 2) + nWindowLeft;
	<% /* MAR1587 M Heys 18/04/2006 the popup can be mislocated in limited number of cases
				this is a quick fix to centre the window without changing other processing*/ %>
	if (nPopupTop < 0 || nPopupLeft < 0)
	{	
		nWindowTop = thisWindow.screenTop;
		nWindowLeft = thisWindow.screenLeft;
		nWindowWidth = thisWindow.parent.document.body.clientWidth;
		nWindowHeight = thisWindow.parent.document.body.clientHeight;
		nPopupTop = Math.floor((nWindowHeight - nPopupHeight) / 2) + nWindowTop;
		nPopupLeft = Math.floor((nWindowWidth - nPopupWidth) / 2) + nWindowLeft;
	}
<%	/*	ShowModalDialog currently appears to have a bit of a bug where the popup resizing doesn't
		always work when the size co-ordinates are passed in as the features argument.
		So instead the co-ordinates are passed in as arguments and the resizing done within the popup
		In order to make this as pretty as possible, center is set to true, the left and top co-ordinates
		are also passed in and, within the popup itself, the screen contents are not shown until the
		resizing has occurred.
	*/

	// Build the arguments array to pass into ShowModalDialog
%>	var sArguments = new Array();
	sArguments[0] = nPopupTop + "px";
	sArguments[1] = nPopupLeft + "px";
	sArguments[2] = nPopupWidth + "px";
	sArguments[3] = nPopupHeight + "px";
	sArguments[4] = sInput;
	
	<% // carry forward the base window through multiple popups %>
	if (thisWindow.dialogArguments != null)
	{
		<%/* SG 06/06/02 SYS4828 - typo */%>
		<%/* sArguments[5] = tempWindow.dialogArguments[5]; */%>
		sArguments[5] = thisWindow.dialogArguments[5];	
	}
	else
	{
		sArguments[5] = thisWindow;
	}
	sArguments[6] = thisWindow; // the caller which may be a popup
	<%/*MC -	Dialog Size and Position has been passed as part of third parameter to the showDialog method.
		
				Note: To make remove the dialog flickering change has been made.
	*/%>
	//return (thisWindow.showModalDialog(sPopup,sArguments,"center: 1; help: no;"));
	return (thisWindow.showModalDialog(sPopup,sArguments,"center: 1; help: no;dialogWidth:"+nPopupWidth + "px;dialogHeight:"+ nPopupHeight +"px;dialogTop:"+ nPopupTop +"px;dialogLeft:"+ nPopupLeft  + "px;" ));
}

function GetRadioGroupValue(frmReference,sGroupName)
{
	for(var nLoop = 0;nLoop < frmReference.elements.length;nLoop++)
	{
		if(frmReference.elements(nLoop).name == sGroupName)
		{
			if(frmReference.elements(nLoop).checked) return frmReference.elements(nLoop).value;
		}
	}

	return null;
}

function SetRadioGroupValue(frmReference,sGroupName,sValue)
{
	for(var nLoop = 0;nLoop < frmReference.elements.length;nLoop++)
	{
		if(frmReference.elements(nLoop).name == sGroupName)
		{
			if(frmReference.elements(nLoop).value == sValue) frmReference.elements(nLoop).checked = true;
			else frmReference.elements(nLoop).checked = false;
		}
	}
}

function SetCheckBoxValue(frmReference, sFieldId, sValue)
{
	if (sValue == frmReference.all(sFieldId).value) 
		frmReference.all(sFieldId).checked = true;
	else 
		frmReference.all(sFieldId).checked = false;
}

function GetCheckBoxValue(frmReference, sFieldId)
{
	if (frmReference.all(sFieldId).checked) 
		return frmReference.all(sFieldId).value;
	return null;
}

function CopyComboList(refFieldFrom,refFieldTo)
{
	if(refFieldFrom.tagName == "SELECT" && refFieldTo.tagName == "SELECT")
	{
		while(refFieldTo.options.length > 0) 
			refFieldTo.options.remove(0);

		for(var nLoop = 0;nLoop < refFieldFrom.options.length;nLoop++)
			refFieldTo.add(refFieldFrom.options.item(nLoop));
		refFieldTo.value = ""; <%/* Set to the <SELECT> option */ %>
	}
	else 
		alert("CopyComboList - fields must be combos");
}

function IsValidationType(refField, sValue)
{
	if(refField.tagName == "SELECT") 
		return IsOptionValidationType(refField,refField.selectedIndex,sValue);
	else 
		alert("IsValidationType - field must be a combo");

	return false;
}

function IsOptionValidationType(refField,nIndex,sValue)
{
	if(refField.tagName == "SELECT")
	{
		if (nIndex > -1)
		{
			var TagOption = refField.options.item(nIndex);
			var nCount = 0;
			var sAttribute = null;
<%			// Read all ValidationType attributes and if one of their values matches the value being looked for return true
			// (the format is ValidationType0 = x ValidationType1 = x etc
%>			do
			{
				sAttribute = TagOption.getAttribute("ValidationType" + nCount);
				if(sAttribute == sValue) return true;
				nCount++;
			}
			while(sAttribute != null)
		}
		else
		{
			alert("IsOptionValidationType - specified index invalid");
		}
	}
	else alert("IsOptionValidationType - field must be a combo");

	return false;
}

function IsOmigaMenuFrame(thisWindow)
{
	for(var nLoop = 0; nLoop < thisWindow.parent.frames.length; nLoop++)
	{
		if(thisWindow.parent.frames.item(nLoop).name == "omigamenu")
			return true;
	}

	return false;
}

function GetNameForStageNumber(thisWindow,sStageNumber)
{
	if(IsOmigaMenuFrame(thisWindow))
	{
		var frmContext = thisWindow.parent.frames("omigamenu").document.forms("frmContext");

		for(var nLoop = 0;nLoop < frmContext("idApplicationStage").options.length;nLoop++)
		{
			if(frmContext("idApplicationStage").options.item(nLoop).value == sStageNumber)
				return frmContext("idApplicationStage").options.item(nLoop).text;
		}

		return "";
	}
	
	var sValue = thisWindow.prompt("Enter name for stage number " + sStageNumber, "");
	if(sValue == null) sValue = "";
	return sValue;
}

function GetContextParameter(thisWindow,sContextParameterId,sDefaultValue)
{
<%	/*	General notes
		If null is specified for sDefaultValue, merely passing it into the prompt causes the word
		"null" to appear in the field, so we must ensure that it is replaced by an empty string
	*/
%>	if(IsOmigaMenuFrame(thisWindow))
	{
		var frmContext = thisWindow.parent.frames("omigamenu").document.forms("frmContext");
		return frmContext(sContextParameterId).value;
	}

	var sValueToShow = "";
	if(sDefaultValue != null) 
		sValueToShow = sDefaultValue;
	var sValue = thisWindow.prompt("Enter value for context parameter " + sContextParameterId, sValueToShow);
	if(sValue == null) 
		sValue = "";
	return sValue;
}

function SetContextParameter(thisWindow,sContextParameterId,sValue)
{
<%	/*	General notes
		If null is specified for sValue, merely passing it into the prompt causes the word
		"null" to appear in the field, so we must ensure that it is replaced by an empty string

		AY 08/09/99 Framework Development
		When setting the customer name context fields, set the visible customer name fields as well

		AY 25/02/00 Framework development
		When setting the Application Number show or hide the navigation menu
	*/
%>	if(!IsOmigaMenuFrame(thisWindow)) 
		return;

	var frmContext = thisWindow.parent.frames("omigamenu").document.forms("frmContext");

	var sValueToSet = "";
	if(sValue != null) 
		sValueToSet = sValue;
	frmContext(sContextParameterId).value = sValueToSet;

	var sDivId = null;
	var sLabelId = null;

	switch(sContextParameterId)
	{
		case "idApplicationNumber":
			if(sValueToSet != "")
			{
				thisWindow.parent.frames("navigation").document.all("divNavigation").style.visibility = "visible";
				<% /* BMIDS948 Navigation menu title has been removed from modal_omigamenu */ %>
				<% /* thisWindow.parent.frames("omigamenu").document.all("divNavigation").style.visibility = "visible"; */ %>
			}
			else
			{
				thisWindow.parent.frames("navigation").document.all("divNavigation").style.visibility = "hidden";
				<% /* thisWindow.parent.frames("omigamenu").document.all("divNavigation").style.visibility = "hidden"; */ %>
			}

			sDivId = "spnApplication";
			sLabelId = "lblApplicationNo";
			SetOtherAppDetailsInContext(sValueToSet);
		break;
		case "idOtherSystemAccountNumber":
			sDivId = "spnAccount";
			sLabelId = "lblAccountNo";
		break;	
		case "idMortgageApplicationDescription":
			sDivId = "spnApplicationType";
			sLabelId = "lblApplicationType";
		break;	
		case "idCustomerName1":
			sDivId = "divCustomer1";
			sLabelId = "lblCustomer1";
		break;
		case "idCustomerName2":
			sDivId = "divCustomer2";
			sLabelId = "lblCustomer2";
		break;
		case "idCustomerName3":
			sDivId = "divCustomer3";
			sLabelId = "lblCustomer3";
		break;
		case "idCustomerName4":
			sDivId = "divCustomer4";
			sLabelId = "lblCustomer4";
		break;
		case "idCustomerName5":
			sDivId = "divCustomer5";
			sLabelId = "lblCustomer5";
		break;
		<% /* PSC 13/10/2005 MAR57 - Start */ %>
		case "idCustomerCategory1":
			sDivId = "divCustomerCategory1";
			sLabelId = "lblCustomerCategory1";
		break;
		case "idCustomerCategory2":
			sDivId = "divCustomerCategory2";
			sLabelId = "lblCustomerCategory2";
		break;
		case "idCustomerCategory3":
			sDivId = "divCustomerCategory3";
			sLabelId = "lblCustomerCategory3";
		break;
		case "idCustomerCategory4":
			sDivId = "divCustomerCategory4";
			sLabelId = "lblCustomerCategory4";
		break;
		case "idCustomerCategory5":
			sDivId = "divCustomerCategory5";
			sLabelId = "lblCustomerCategory5";
		break;
		<% /* PSC 13/10/2005 MAR57 - End */ %>
		// IK,15/03/01,SYS1924
		case "idStageId":
			thisWindow.parent.frames("navigation").document.all("btnStageTrigger").click();
		break;
		<% /* MAR1300 GHun */ %>
		case "idPropertyChange":
			sDivId = "divPropertyChange";
			sLabelId = "lblPropertyChange";
		break;
		default:
		break;
	}
<%	/*	If the application number or a customer name context field has been set:

		Clear the title attribute. This will only be set if the label is visible and has overflowed
		N.B. we have to set the title on the background div because it won't work on the label

		If sValueToSet is empty, clear the label and hide the span containing the label and its bullet
		Else make the label and its bullet visible
	*/
%>	if(sDivId != null)
	{
		var thisDiv = thisWindow.parent.frames("omigamenu").document.all(sDivId);
		var thisLabel = thisWindow.parent.frames("omigamenu").document.all(sLabelId);
		thisDiv.title = "";

		if(sValueToSet == "")
		{
			thisDiv.style.visibility = "hidden";
			thisLabel.innerHTML = "";
		}
		else
		{
			thisDiv.style.visibility = "visible";
<%			/*	To format the name field:
				Initialise the label with something and get its offsetHeight since if the label was
				empty, 0 would be returned
				Get the width specified for the label
					
				Then set the label value to the full input value and remember the length of
				the input value at this point

				Keep populating the label with a modified string until it fits on one line within the
				allowable width or until the length of the string is zero (a failsafe only - this 
				shouldn't happen).  This is done by:

					Setting the label to the label value and getting the label height

					If the height has increased then the text has wrapped to more than one line
					or the offsetWidth > available width
					Then decrease the length of the label value and add ... to the end

				If the label value does not equal the full value, set the title attribute
			*/
%>			thisLabel.innerHTML = "X";
			var nStartHeight = thisLabel.offsetHeight;
			var nStartWidth = thisLabel.style.posWidth;
			var sLabelValue = sValueToSet;
			var nLength = sValueToSet.length;
			var nNewHeight;
			var nNewWidth;

			do
			{
				thisLabel.innerHTML = sLabelValue;
				nNewHeight = thisLabel.offsetHeight;
				nNewWidth = thisLabel.offsetWidth;

				if(nNewHeight > nStartHeight || nNewWidth > nStartWidth)
				{
					nLength--;
					sLabelValue = sValueToSet.substr(0,nLength) + "...";
				}
			}
			while((nNewHeight > nStartHeight && nLength > 0 || nNewWidth > nStartWidth) && nLength > 0);

			if(sLabelValue != sValueToSet)
				thisDiv.title = sValueToSet;
		}

		if(thisWindow.parent.frames("omigamenu").document.all("divCustomer1").style.visibility == "visible"
		   || thisWindow.parent.frames("omigamenu").document.all("divCustomer2").style.visibility == "visible"
		   || thisWindow.parent.frames("omigamenu").document.all("divCustomer3").style.visibility == "visible"
		   || thisWindow.parent.frames("omigamenu").document.all("divCustomer4").style.visibility == "visible"
		   || thisWindow.parent.frames("omigamenu").document.all("divCustomer5").style.visibility == "visible")
			thisWindow.parent.frames("omigamenu").document.all("spnCustomers").style.visibility = "visible";
		else
			thisWindow.parent.frames("omigamenu").document.all("spnCustomers").style.visibility = "hidden";
	}
}

function SyncCustomerIndex(thisWindow)
{
	var sCustNo = GetContextParameter(thisWindow,"idCustomerNumber",null);
	var nIndex = 1;
	if(sCustNo == GetContextParameter(thisWindow,"idCustomerNumber1",null))
	{
		SetContextParameter(thisWindow,"idCustomerIndex",nIndex);
		return nIndex;
	}
	nIndex++;
	if(sCustNo == GetContextParameter(thisWindow,"idCustomerNumber2",null))
	{
		SetContextParameter(thisWindow,"idCustomerIndex",nIndex);
		return nIndex;
	}
	nIndex++;
	if(sCustNo == GetContextParameter(thisWindow,"idCustomerNumber3",null))
	{
		SetContextParameter(thisWindow,"idCustomerIndex",nIndex);
		return nIndex;
	}
	nIndex++;
	if(sCustNo == GetContextParameter(thisWindow,"idCustomerNumber4",null))
	{
		SetContextParameter(thisWindow,"idCustomerIndex",nIndex);
		return nIndex;
	}
	nIndex++;
	if(sCustNo == GetContextParameter(thisWindow,"idCustomerNumber5",null))
	{
		SetContextParameter(thisWindow,"idCustomerIndex",nIndex);
		return nIndex;
	}
	return -1;
}

function GetContextCustomerName(thisWindow,sCustomerNumber)
{
	if(IsOmigaMenuFrame(thisWindow))
	{
<%		// Go through each set of customer context fields to find a matching customer number
%>		for(var nLoop = 1;nLoop <= 5;nLoop++)
		{
			var sCustomerNumberId = "idCustomerNumber" + nLoop;
			var sCustomerNameId = "idCustomerName" + nLoop;
			var frmContext = thisWindow.parent.frames("omigamenu").document.forms("frmContext");

			if(frmContext(sCustomerNumberId).value == sCustomerNumber)
				return frmContext(sCustomerNameId).value;
		}
	}
	else return "<NO CONTEXT>";

	return "<ERROR>";
}

function SetFocusToFirstField(frmReference)
{
	for(var nLoop = 0;nLoop < frmReference.all.length;nLoop++)
	{
		if(DoFocusProcessing(frmReference,nLoop))
			return;
	}

	var msgButtons = frmReference.parentElement.all("msgButtons");
	if(msgButtons != null)
	{
		for(var nLoop = 0;nLoop < msgButtons.all.length;nLoop++)
		{
			if(DoFocusProcessing(msgButtons,nLoop))
				return;
		}
	}
}

function SetFocusToLastField(frmReference)
{
	var msgButtons = frmReference.parentElement.all("msgButtons");
	if(msgButtons != null)
	{
		for(var nLoop = msgButtons.all.length - 1;nLoop >= 0;nLoop--)
		{
			if(DoFocusProcessing(msgButtons,nLoop))
				return;
		}
	}

	for(var nLoop = frmReference.all.length - 1;nLoop >= 0;nLoop--)
	{
		if(DoFocusProcessing(frmReference,nLoop))
			return;
	}
}

function DoFocusProcessing(frmReference,nElement)
{
	var bIsFocusSet	= false;
	var nFocusField	= null;
	var thisElement = frmReference.all(nElement);
<%	/*	Check the attributes on the field to decide if it can accept focus	
		
		Buttons and check boxes:
			Check the tabindex, disabled and visibility attributes
			
		Radio buttons:
			As above, but also check if one of the radios is set. If one is set the focus
			to that button, otherwise set to the first in the group

		Text fields:
			Check the tabindex, disabled, readonly and visibility attributes

		Text area and select fields:
			Check the tabindex, disabled and visibility attributes

		If a field is found, set the focus and return true
	*/
%>	if(thisElement.tagName == "INPUT")
	{
		if(thisElement.type == "button" || thisElement.type == "radio" || thisElement.type == "checkbox")
		{
			if(thisElement.tabIndex != -1 && !thisElement.disabled && thisElement.style.visibility != "hidden")
			{
				if(thisElement.type == "radio")
				{
					var sGroupName = thisElement.name;
					for(var nLoop = 0;nLoop < frmReference.all.length;nLoop++)
					{
						if(frmReference.all(nLoop).name == sGroupName)
						{
							if(nFocusField == null) nFocusField = nLoop;
							if(frmReference.all(nLoop).checked == true) nFocusField = nLoop;
						}
					}
				}
				else nFocusField = nElement;
			}
		}
		else
			if(thisElement.tabIndex != -1 && !thisElement.disabled && !thisElement.readonly && thisElement.style.visibility != "hidden")
				nFocusField = nElement;
	}

	if(thisElement.tagName == "TEXTAREA" || thisElement.tagName == "SELECT")
	{
		if(thisElement.tabIndex != -1 && !thisElement.disabled && thisElement.style.visibility != "hidden")
			nFocusField = nElement;
	}

	if(nFocusField != null)
	{
		frmReference.all(nFocusField).focus();
		bIsFocusSet = true;
	}

	return bIsFocusSet;
}

function SizeTextToField(refField,sValue)
{
<%	/*	Clear any existing title attribute
		If the value is an empty string clear the field
		Otherwise, see if the text fits into the field as follows:
		
		Get the starting width and height
		
		Then set the text value to the full input value and remember the length of
		the input value at this point

		Keep populating the text with a modified string until it fits on one line within the
		allowable width or until the length of the string is zero (a failsafe only - this 
		shouldn't happen).  This is done by:

			Setting the text to the text value and getting the height

			If the height has increased then the text has wrapped to more than one line
			or the offsetWidth > available width
			Then decrease the length of the text value and add ... to the end

		If the label value does not equal the full value, set the title attribute
	*/
%>	
	if( sValue == null )
		sValue = "";
	
	if(refField.tagName == "TD" || refField.tagName == "LABEL")
	{
		refField.title = "";

		if(sValue == "")
		{
			if(refField.tagName == "TD") 
				refField.innerText = " ";
			if(refField.tagName == "LABEL")
				refField.innerHTML = "";
			return;
		}
		else
		{
			if(refField.tagName == "LABEL") refField.innerHTML = "X";
			<% /* BM0095 put dummy text into the field before measuring the height */ %>
			if(refField.tagName == "TD") refField.innerText = "x";

			var nStartHeight = refField.offsetHeight;
			var nStartWidth;
			if(refField.tagName == "TD")
				nStartWidth = refField.offsetWidth;
			if(refField.tagName == "LABEL")
				nStartWidth = refField.style.posWidth;
			var sTextValue = sValue;
			var nLength = sValue.length;
			var nNewHeight;
			var nNewWidth;

			do
			{
				if(refField.tagName == "TD")
					refField.innerText = sTextValue;
				if(refField.tagName == "LABEL")
					refField.innerHTML = sTextValue;

				nNewHeight = refField.offsetHeight;

				<% //MAR1289 %>
				<% //Peter Edney - 21/02/2006 %>
				if(refField.tagName == "TD")
					nNewWidth = refField.offsetWidth;
				if(refField.tagName == "LABEL")
					nNewWidth = refField.style.posWidth;

				if(nNewHeight > nStartHeight || nNewWidth > nStartWidth)
				{
					nLength--;
					sTextValue = sValue.substr(0,nLength) + "...";
				}
			}
			while((nNewHeight > nStartHeight && nLength > 0 || nNewWidth > nStartWidth) && nLength > 0);

			if(sTextValue != sValue)
				refField.title = sValue;
		}
	}
}
<% /* MO 13/11/2002 BMIDS00807 : Modified these functions to use GetAppServerDate rather than GetSystemDate - START*/ %>
function CompareDateFieldToSystemDate(refDateField,sComparison)
{
	return CompareDates(GetDateObject(refDateField),sComparison,GetSystemDate(), GetAppServerDate());
}

function CompareDateStringToSystemDate(sDate,sComparison)
{
	return CompareDates(GetDateObjectFromString(sDate),sComparison,GetSystemDate(), GetAppServerDate());
}
function CompareDateStringToSystemDateTime(sDate,sComparison)
{
	return CompareDates(GetDateObjectFromString(sDate),sComparison,GetSystemDateTime(), GetAppServerDate(true));
}
<% /* MO 13/11/2002 BMIDS00807 : Modified these functions to use GetAppServerDate rather than GetSystemDate - END */ %>

function CompareDateFields(refFirstDateField,sComparison,refSecondDateField)
{
	return CompareDates(GetDateObject(refFirstDateField),sComparison,GetDateObject(refSecondDateField));
}

function DateToString(dtDate)
{
	var nYear = dtDate.getFullYear();
	var nMonth = (dtDate.getMonth() + 1);
	var nDay = dtDate.getDate();

	if (nMonth < 10)
		nMonth = "0" + nMonth;
	if (nDay < 10)
		nDay = "0" + nDay;

	var sDate = nDay + "/" + nMonth + "/" + nYear;

	return(sDate);
}
function DateTimeToString(dtDate)
{
	var nYear = dtDate.getFullYear();
	var nMonth = (dtDate.getMonth() + 1);
	var nDay = dtDate.getDate();
	var nHours = dtDate.getHours();
	var nMinutes = dtDate.getMinutes();
	var nSeconds = dtDate.getSeconds();

	if (nMonth < 10)
		nMonth = "0" + nMonth;
	if (nDay < 10)
		nDay = "0" + nDay;
	if (nHours < 10)
		nHours = "0" + nHours;
	if (nMinutes < 10)
		nMinutes = "0" + nMinutes;
	if (nSeconds < 10)
		nSeconds = "0" + nSeconds;
		

	var sDate = nDay + "/" + nMonth + "/" + nYear + " " + nHours + ":" + nMinutes + ":" + nSeconds;

	return(sDate);
}
function GetSystemDate()
{
	var dtTempCurrentDate = new Date();
	var dtCurrentDate = new Date(dtTempCurrentDate.getFullYear(),dtTempCurrentDate.getMonth(),dtTempCurrentDate.getDate());

	return (dtCurrentDate);
}
function GetSystemDateTime()
{
	var dtTempCurrentDate = new Date();
	var dtCurrentDate = new Date(dtTempCurrentDate.getFullYear(),dtTempCurrentDate.getMonth(),dtTempCurrentDate.getDate(),dtTempCurrentDate.getHours(), dtTempCurrentDate.getMinutes(), dtTempCurrentDate.getSeconds());

	return (dtCurrentDate);
}

function CompareDates(dtFirstDate,sComparison,dtSecondDate)
{
<%	// Only compare if both are valid dates
%>	if(dtFirstDate != null && dtSecondDate != null)
	{
		switch(sComparison)
		{
			case "<":
				if(dtFirstDate < dtSecondDate) return true;
			break;
			case "<=":
				if(dtFirstDate <= dtSecondDate) return true;
			break;
			case "=":
				if(dtFirstDate == dtSecondDate) return true;
			break;
			case ">=":
				if(dtFirstDate >= dtSecondDate) return true;
			break;
			case ">":
				if(dtFirstDate > dtSecondDate) return true;
			break;
			default:
				alert("CompareDates - Invalid comparison string");
			break;
		}
	}

	return false;
}

function GetDateObject(refDateField)
{
	if(refDateField.value.length > 0)
	{
		if(refDateField.value.length == 10 && !refDateField.style.textDecorationLineThrough)
			return(GetDateObjectFromString(refDateField.value));
		else
		{
			alert("Invalid Date Format");
			refDateField.focus();
		}
	}

	return null;
}

function GetDateObjectFromString(sDate)
{
	if (sDate.length == 10)
	{
		var sDay = sDate.substr(0,2);
		var sMonth = sDate.substr(3,2);
		var sYear = sDate.substr(6,4);
		return (new Date(sYear,sMonth-1,sDay));
	}
	else if(sDate.length == 19) //date string also has time information
	{
		var sDay = sDate.substr(0,2);
		var sMonth = sDate.substr(3,2);
		var sYear = sDate.substr(6,4);
		var sHour = sDate.substr(11,2);
		var sMinutes = sDate.substr(14,2);
		var sSecs = sDate.substr(17,2);
		return (new Date(sYear,sMonth-1,sDay,sHour,sMinutes,sSecs));
	}
	else
		return null;
}

function CompletenessCheckRouting(thisWindow)
{
	var sContext = GetContextParameter(thisWindow,"idProcessContext",null);
	var bCompletenessCheck = false;
	if(sContext == "CompletenessCheck")
	{
		bCompletenessCheck = true;
	}
	return bCompletenessCheck;
}

function GetComboValidationType(frmReference, sFieldId, iIndex)
{
	var sValidationType = "";

	if (frmReference.all(sFieldId).tagName == "SELECT")
	{		
		if (iIndex == null)
			iIndex = frmReference.all(sFieldId).selectedIndex;

		if (iIndex != -1)
		{
			var TagOption = frmReference.all(sFieldId).options.item(iIndex);
			sValidationType = TagOption.getAttribute("ValidationType0");
		}
	}

	return sValidationType;
}

function SetComboOnValidationType(frmReference, sFieldId, sValue)
{
	var blnFound = false;

	for(var nLoop = 0;nLoop < frmReference.all(sFieldId).options.length && !blnFound;nLoop++)
		blnFound = IsOptionValidationType(frmReference.all(sFieldId), nLoop, sValue);

	if (blnFound) 
		frmReference.all(sFieldId).selectedIndex = nLoop-1;
	else 
		alert("Combo option validation type not found");
}

function EnableDrillDown(refField)
{
	if(refField.tagName == "INPUT")
	{
		if(refField.type == "button")
		{
			if(refField.className == "msgDisabledDDButton")
			{
				refField.className = "msgDDButton";
				refField.disabled = false;
				refField.tabIndex = 0;
			}
		}
	}
}

function DisableDrillDown(refField)
{
	if(refField.tagName == "INPUT")
	{
		if(refField.type == "button")
		{
			if(refField.className == "msgDDButton")
			{
				refField.className = "msgDisabledDDButton";
				refField.disabled = true;
				refField.tabIndex = -1;
			}
		}
	}
}

function SetFieldToReadOnly(frmReference, sFieldId)
{
	SetFieldState(frmReference, sFieldId, "R");
}

function SetFieldToWritable(frmReference, sFieldId)
{
	SetFieldState(frmReference, sFieldId, "W");
}

function SetFieldToDisabled(frmReference, sFieldId)
{
	SetFieldState(frmReference, sFieldId, "D");
}

function SetRadioGroupToReadOnly(frmReference, sGroupName)
{
	SetRadioGroupState(frmReference,sGroupName,"R");
}

function SetRadioGroupToWritable(frmReference, sGroupName)
{
	SetRadioGroupState(frmReference,sGroupName,"W");
}

function SetRadioGroupToDisabled(frmReference, sGroupName)
{
	SetRadioGroupState(frmReference,sGroupName,"D");
}

function EnableCollection(refSpnOrDivId)
{		
	SetCollectionState(refSpnOrDivId,"W");
}

function DisableCollection(refSpnOrDivId)
{
	SetCollectionState(refSpnOrDivId,"D");
}

function SetCollectionToReadOnly(refSpnOrDivId)
{
	SetCollectionState(refSpnOrDivId,"R");
}

function SetCollectionToWritable(refSpnOrDivId)
{
	SetCollectionState(refSpnOrDivId,"W");
}

function SetCurrency(thisWindow,frmReference)
{
	return SetThisCurrency(frmReference,GetContextParameter(thisWindow,"idCurrency",""));
}

function SetThisCurrency(frmReference,sCurrencyParameter)
{
	var sCurrency;
	switch(sCurrencyParameter)
	{
		case "Euro":
			sCurrency = "€";
		break;
		case "Dollar":
			sCurrency = "$";
		break;
		default:
			sCurrency = "£";
		break;
	}

	for(var nLoop = 0;nLoop < frmReference.all.length;nLoop++)
	{
		if(frmReference.all(nLoop).className == "msgCurrency")
		{
			if(frmReference.all(nLoop).tagName == "LABEL")
				frmReference.all(nLoop).innerHTML = sCurrency;
		}
	}

	return sCurrency;
}

function ValidatePostcode(refField)
{
	var arrMasks;
	var text = null;
	var bReturn = false;

	<%// In mask @=[A-Z], #=[0-9]. All others are literals 
	// Build an array of masks, select on length and then check the
	// apropriate ones.%>
	arrMasks = new Array(8);
	arrMasks[0]=new Array();
	<% /* BM0242 Support partial postcodes */ %>

	<% /* arrMasks[1]=new Array(); */ %>
	arrMasks[1]=new Array ("@");
	<% /* arrMasks[2]=new Array(); */ %>
	arrMasks[2]=new Array ("@#");
	arrMasks[3]=new Array ("@@#","@#@","@##","IOM");
	<% /* arrMasks[4]=new Array ("@@#@","@# #","BT74"); */ %>
	arrMasks[4]=new Array ("@@##","@@#@","@# #","BT74");
	<% /* arrMasks[5]=new Array ("@@# #","@#@ #","@## #"); */ %>
	arrMasks[5]=new Array ("@@# #","@#@ #","@## #","@# #@");
	<% /* arrMasks[6]=new Array ("@@#@ #","@@## #","@# #@@"); */ %>
	arrMasks[6]=new Array ("@@#@ #","@@## #","@# #@@","@## #@","@@# #@","@#@ #@");
	<% /* arrMasks[7]=new Array ("@@# #@@","@#@ #@@","@## #@@"); */ %>
	arrMasks[7]=new Array ("@@# #@@","@#@ #@@","@## #@@","@@## #@","@@#@ #@");
	arrMasks[8]=new Array ("@@#@ #@@","@@## #@@");
	<% /* BM0242 End */ %>

	text=new String(refField.value);
	<% /* BM0242 Remove duplicate spaces before validating */ %>
	text=text.replace(/\s+/," ");
	<% /* BM0242 End */ %>
	
	if (text.length==0)
		bReturn=true;
	else if (text.length<9) 
		bReturn= ValidateWithMasks(text,arrMasks[text.length]);
	else
		bReturn=false;
	
	if (!bReturn) {
		alert("Please enter a valid Postcode");
		refField.focus();
	}
	
	return bReturn;
}

function ValidateWithMasks(sText,arrMasks)
{
	var i =null;
	var iPos =null;
	var bResult=false;
	
	for(i=0;i<arrMasks.length;i++)
	{
		bResult=bResult | ValidateWithMask(sText,arrMasks[i]);
	}
	
	return bResult;
}		
	
function ValidateWithMask(sText,sMask)
{
	// Pre-condition - the mask and text are the same length
	var bMatch = true;
	var iLen = null;
	var iPos= null;

	iLen=sMask.length;

	for(iPos=0;iPos<iLen;iPos++) {
		switch (sMask.substr(iPos,1)) 
		{
		case "@":
			bMatch=bMatch & ! (sText.substr(iPos,1)<"A" || sText.substr(iPos,1)>"Z");
			break;
		case "#":
			bMatch=bMatch & ! (sText.substr(iPos,1)<"0" || sText.substr(iPos,1)>"9");
			break;
		default:
			bMatch=bMatch & (sMask.substr(iPos,1)==sText.substr(iPos,1));
			break;
		}
	}
	return bMatch;
}	

function RestrictLength(frmReference, sFieldId, nMaxLength, bMessage)
{
	var thisObj = frmReference.all(sFieldId);
	if(thisObj.tagName != "TEXTAREA" || nMaxLength < 2)
		return false;
		
	var strString = thisObj.value;
	if(strString.length > nMaxLength)
	{
		thisObj.value = strString.slice(0, nMaxLength);
		if(bMessage == true)
		{
			alert("Cannot enter more than " + nMaxLength + " characters");
			thisObj.focus();
		}
		return true;
	}
	else
		return false;
}

function LogonLegallyOpened()
{
	var blnRetVal=true;
	try
	{
		<%
		//Reference the required field directly - If fails then error gets trapped & fn returns false
		%>
		if (top.frames("fraDetails").document.forms("frmDetails").txtUserId.value == "")
		{
			<%
			//Ref OK - but user ID is empty
			%>
			blnRetVal = false;
			
		}
	}
	catch(e)
	{
		<%
		//An error has occured
		//The frame and/or form and or/user id cannot be found
		%>
		
		blnRetVal = false;
	}
	return(blnRetVal);
}

function Omiga4LegallyOpened(blnFirstScreen)
{

	var blnRetVal, strUserId_opened, strUserId_opener;

	blnRetVal= true;

	try
	{			
		<%
		//Reference the user id field on the opener directly - If fails then error gets trapped & fn returns false
		%>
		
		strUserId_opener = top.window.opener.parent.document.frames("fraDetails").document.forms("frmDetails").txtUserId.value;
		if (strUserId_opener == "")
		{
			<%
			//If we've got this far with no error, then check the contents of userid in opener for emptiness
			%>

			blnRetVal = false;
		}
		
		strUserId_opened=top.frames("omigamenu").document.forms("frmContext").idUserId.value;
		if ((!blnFirstScreen) && (strUserId_opened==""))
		{
			blnRetVal = false;
		}
		
	}
	catch(e)
	{
		<%
		//An error has occured
		//The frame and/or form and or/user id cannot be found
		%>
		blnRetVal = false;
	}

	return(blnRetVal);
}

function SetMainScreenToReadOnly(thisWindow, frmReference, sScreenId)
{
	if (this.IsMainScreenReadOnly(thisWindow, sScreenId) == true) 
	{
		this.SetScreenToReadOnly(frmReference);
		return true;
	}
	return false;
}

function IsMainScreenReadOnly(thisWindow, sScreenId)
{
	// Read only 
	var sReadOnly = this.GetContextParameter(thisWindow, "idReadOnly", "0");
	if (sReadOnly == "1") {		
		return true;		
	}
	
	// Processing unit 
	var sProcessingIndicator = this.GetContextParameter(thisWindow, "idProcessingIndicator", "0");
	if (sProcessingIndicator != "1") {		
		return true;		
	}
	
	// Freeze the Data?
	var sDataFreezeIndicator = this.GetContextParameter(thisWindow, "idFreezeDataIndicator", "0");
	if (sDataFreezeIndicator == "1") 
	{
		if (IsDataFreezeScreen(thisWindow, sScreenId) == true)
		{
			return true;
		}		
	}
	
	//CancelDeclineData Freeze the  Data?
	var sCancelDeclineDataFreezeIndicator = this.GetContextParameter(thisWindow, "idCancelDeclineFreezeDataIndicator", "0");
	if (sCancelDeclineDataFreezeIndicator == "1") 
	{
		if (IsCancelDeclineDataFreezeScreen(thisWindow, sScreenId) == true){
			return true;
		}		
	}
	
	return false;
}

function IsDataFreezeScreen(thisWindow, sScreenId)
{
	var iIndex = -1;
	
	if(IsOmigaMenuFrame(thisWindow))
	{
		var frmContext = thisWindow.parent.frames("omigamenu").document.forms("frmContext");
				
		var sDataFreezeScreens = frmContext.idDataFreezeScreens.innerText;
	
		if (sDataFreezeScreens.length > 0) {
			iIndex = sDataFreezeScreens.indexOf(sScreenId.toUpperCase());
		}
	}
	
	return (iIndex >= 0);
}

function IsCancelDeclineDataFreezeScreen(thisWindow, sScreenId)
{
	var iIndex = -1;
	
	if(IsOmigaMenuFrame(thisWindow))
	{
		var frmContext = thisWindow.parent.frames("omigamenu").document.forms("frmContext");
				
		var sCancelDeclineDataFreezeScreens = frmContext.idCancelDeclineDataFreezeScreens.innerText;
	
		if (sCancelDeclineDataFreezeScreens.length > 0) 
		{
			iIndex = sCancelDeclineDataFreezeScreens.indexOf(sScreenId.toUpperCase());
		}
		
	}
	
	return (iIndex >= 0);
}

function IsCancelDeclineStage(thisWindow)
{
	//CancelDeclineData Freeze the  Data?
	var bRet = false;
	var sCancelDeclineDataFreezeIndicator = this.GetContextParameter(thisWindow, "idCancelDeclineFreezeDataIndicator", "0");
	if (sCancelDeclineDataFreezeIndicator == "1") 
	{
		
		bRet = true;
		
	}
	return bRet;
}

function getCancelledStageValue(thisWindow)
{
	return this.GetContextParameter(thisWindow, "idCancelledStageValue", "0");
}

function getDeclinedStageValue(thisWindow)
{
	return this.GetContextParameter(thisWindow, "idDeclinedStageValue", "0");
}

function SetOtherSystemACNoInContext(sApplicationNumber)
{
	var sOtherSystemAccountNo = "";
	if (sApplicationNumber != "")
	{
		var XML = GetApplicationData(sApplicationNumber);
		sOtherSystemAccountNo = XML.GetTagText("OTHERSYSTEMACCOUNTNUMBER");
	} else 

	<%//GD BM0356 START%>
	<%//SetContextParameter(window, "idOtherSystemAccountNumber" , sOtherSystemAccountNo) ;%>
	SetContextParameter(top.frames("main"), "idOtherSystemAccountNumber" , sOtherSystemAccountNo);
	<%//GD BM0356 END%>
}

<% /* BMIDS821 GHun Set OtherSystemAccountNumber and ApplicationDate in context */ %>
function SetOtherAppDetailsInContext(sApplicationNumber)
{
	var sOtherSystemAccountNo = "";
	var sApplicationDate = "";
	if (sApplicationNumber != "")
	{
		var XML = GetApplicationData(sApplicationNumber);
		sOtherSystemAccountNo = XML.GetTagText("OTHERSYSTEMACCOUNTNUMBER");
		sApplicationDate = XML.GetTagText("APPLICATIONDATE");
	}
	SetContextParameter(top.frames("main"), "idOtherSystemAccountNumber", sOtherSystemAccountNo);
	SetContextParameter(top.frames("main"), "idApplicationDate", sApplicationDate);
}
<% /* BMIDS821 End */ %>

<% /* BMIDS821 GHun Changed GetOtherSystemAccountNumber to return all the XML from the
	GetApplicationData call, rather than just one field.
function GetOtherSystemAccountNumber(sApplicationNumber)
*/ %>
function GetApplicationData(sApplicationNumber)
{
	var XML = new scXMLFunctions.XMLObject();
	<%/* MV - 21/08/2002 - BMIDS0355 - Code Error*/%>
	XML.CreateRequestTag(top.frames("main") ,"") ;
	XML.CreateActiveTag("APPLICATION");
	XML.CreateTag("APPLICATIONNUMBER", sApplicationNumber);
	
	XML.RunASP(document, "GetApplicationData.asp");
	if (!XML.IsResponseOK())
		return " ";
	else
	{
		XML.CreateTagList('APPLICATION');
		XML.SelectTagListItem(0);
		<% /* return XML.GetTagText("OTHERSYSTEMACCOUNTNUMBER"); */ %>
		return XML;
	}
}
<% /* BMIDS821 End */ %>

function FormatAsCurrency(value)
{
	if (value < 1000) 
	{ 
		return new String(value);
	}
	else
	{
		var sString = new String(value);
		var sDecimalPart = sString.substr(sString.length-3, sString.length);
		var intPart = new String(sString.substr(0, sString.length-3));
		var targetStr = new String();
	
		if (intPart.length > 3)
		{
			var leadNums = intPart.length % 3;
			if (leadNums != 0) 
			{
				targetStr = targetStr.concat(intPart.substr(0,leadNums)) + ",";
			}
			intPart = intPart.substr(leadNums, intPart.length - leadNums);
			for (var i = 0; i < intPart.length -3 ; i+=3)
			{
				targetStr = targetStr.concat(intPart.substr(i,3)) + ",";
			}
			targetStr = targetStr.concat(intPart.substr(i,3)) + sDecimalPart;
		}	
		return targetStr;
	}
}

function DisplayClientWarning(sMessage, sICON, sOKLegend, sCancelLegend)
{
	var sReturn = null;

	// Build the arguments array to pass into ShowModalDialog
	var sArguments = new Array(6);

	sArguments[0] = window;
	sArguments[1] = sMessage; 	// Message to be displayed
	sArguments[2] = 1;		// Type = Warning
	sArguments[3] = sICON;		// ICON to display
	sArguments[4] = sOKLegend;	// Legend for 'OK' button
	sArguments[5] = sCancelLegend;		// Legend for 'Cancel' button
	
	//SD 12/10/2005 - Start
	//sReturn = window.showModalDialog("MsgBox.asp",sArguments);
	sReturn = window.showModalDialog("MsgBox.asp",sArguments,"status=no;scroll=no");
	//SD 12/10/2005 - End

	if(sReturn != null)
	{
		return (sReturn[0]);
	}
}

function DisplayClientError(sMessage, sICON)
{
	var sReturn = null;

	// Build the arguments array to pass into ShowModalDialog
	var sArguments = new Array(4);

	sArguments[0] = window;
	sArguments[1] = sMessage; 	// Message to be displayed
	sArguments[2] = 2;		// Type = Error
	sArguments[3] = sICON;		// ICON to display

	//SD 12/10/2005 - Start
	//sReturn = window.showModalDialog("MsgBox.asp",sArguments);
	sReturn = window.showModalDialog("MsgBox.asp",sArguments,"status=no;scroll=no");
	//SD 12/10/2005 - End

	if(sReturn != null)
	{
		return (sReturn[0]);
	}
}

<%/* MO 13/11/2002 BMIDS00807 : SYSMCP1180 - use date / dateTime from application server  */ %>
function GetAppServerDate(bFullDateAndTime)
{
	var XML = new scXMLFunctions.XMLObject();
	XML.RunASP(document, "DateTimeOnAppServer.asp");
	XML.SelectTag(null, "RESPONSE");
	
	var dtTempCurrentDate = new Date(Date.parse(XML.ActiveTag.text));
	
	if(arguments.length < 1 || bFullDateAndTime == false)
	{
		return new Date(dtTempCurrentDate.getFullYear(),dtTempCurrentDate.getMonth(),dtTempCurrentDate.getDate());
	}
	else
	{
		return dtTempCurrentDate ;
	}
}
<%/*  MO 13/11/2002 BMIDS00807 : SYSMCP1180 - End */ %>
<% /* PJO 08/03/2006 MAR1359 Progress Message */ %>

var m_ProgWindow ;

function progressOn (sText, iWidth) 
{
    var iHeight = 140; <% /* MAR1453 GHun */ %>
    var iLeft = (window.screen.width - iWidth) / 2 ;
    var iTop = (window.screen.height - iHeight) / 2 ;
 

	var sAttributes = "toolbar=No," +
					  "location=No," +
					  "directories=No," +
					  "menubar=No," +
					  "titlebar=No," +
		              "scrollbars=No," +
		              "status=No," + 
		              "titlebar=No," + 
		              "resizable=No," +
		              "width=" + iWidth + "," +
		              "height=" + iHeight + "," +
		              "left=" + iLeft + "," +
		              "top=" + iTop ;
    m_ProgWindow = window.open("progress.asp?" + sText, "Progress", sAttributes);
} 

function progressOff() 
{ 
    if (m_ProgWindow) 
    {
        m_ProgWindow.close(); 
    }
    m_ProgWindow = null;
}

</script>
