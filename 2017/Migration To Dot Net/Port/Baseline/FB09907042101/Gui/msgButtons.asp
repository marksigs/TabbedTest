<%
/*
Workfile:      msgButtons.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Control file for the screen buttons
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		15/11/99	Addition of the Loans button
AY		19/01/00	Comments set as server side script and
					unnecessary characters removed
					N.B. this file does not contain the server side
					language setting because it is being included
					in files which already contain it
AY		11/02/00	Buttons changed from img to input type=button
AS		21/11/00	CORE000010: Added in btnBack button
GD		25/01/01	Added Confirm button as part of SYS1833
PSC		05/08/02	BMIDS00006 Added Arrears and Features
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

To use this file:

In the window.onload of the calling file set up an array of the buttons to be displayed.  The array should
consist of a mixture of the following strings (please amend as buttons are added/deleted):

Submit (the OK button - because the button performs submit events it is referred to as submit)
Cancel
Undo
Another
Loans
Previous
Next
N.B. The strings are case insensitive.

Pass this array into ShowMainButtons(sButtonArray).
The specified buttons will be displayed left to right in the order in which they are listed.

To then Enable or Disable any of these buttons, call EnableMainButton(sButton) or DisableMainButton(sButton)
passing as a string one of the values listed above.
*/
%>
<div id="divSubmit" style="TOP: 0px; LEFT: 0px;HEIGHT:1px; WIDTH: 1px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE">
	<input id="btnSubmit" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnOK.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" accesskey="O">
	<input id="btnSubmitDisabled" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnOK_dis.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" disabled tabindex="-1">
</div>

<div id="divPrevious" style="TOP: 0px; LEFT: 0px;HEIGHT:1px; WIDTH: 1px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE">
	<input id="btnPrevious" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnPrev.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" accesskey="P">
	<input id="btnPreviousDisabled" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnPrev_dis.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" disabled tabindex="-1">
</div>

<div id="divNext" style="TOP: 0px; LEFT: 0px;HEIGHT:1px; WIDTH: 1px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE">
	<input id="btnNext" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnNext.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" accesskey="N">
	<input id="btnNextDisabled" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnNext_dis.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" disabled tabindex="-1">
</div>

<div id="divFinish" style="TOP: 0px; LEFT: 0px;HEIGHT:1px; WIDTH: 1px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE">
	<input id="btnFinish" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnFinish.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" accesskey="F">
	<input id="btnFinishDisabled" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnFinish_dis.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" disabled tabindex="-1">
</div>

<div id="divCancel" style="TOP: 0px; LEFT: 0px;HEIGHT:1px; WIDTH: 1px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE">
	<input id="btnCancel" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnCancel.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" accesskey="C">
	<input id="btnCancelDisabled" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnCancel_dis.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" disabled tabindex="-1">
</div>

<div id="divUndo" style="TOP: 0px; LEFT: 0px;HEIGHT:1px; WIDTH: 1px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE">
	<input id="btnUndo" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnUndo.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" accesskey="U">
	<input id="btnUndoDisabled" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnUndo_dis.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" disabled tabindex="-1">
</div>

<div id="divAnother" style="TOP: 0px; LEFT: 0px;HEIGHT:1px; WIDTH: 1px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE">
	<input id="btnAnother" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnAnother.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" accesskey="A">
	<input id="btnAnotherDisabled" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnAnother_dis.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" disabled tabindex="-1">
</div>

<div id="divLoans" style="TOP: 0px; LEFT: 0px;HEIGHT:1px; WIDTH: 1px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE">
	<input id="btnLoans" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnLoans.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" accesskey="L">
	<input id="btnLoansDisabled" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnLoans_dis.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" disabled tabindex="-1">
</div>

<div id="divBack" style="TOP: 0px; LEFT: 0px;HEIGHT:1px; WIDTH: 1px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE">
	<input id="btnBack" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnBack.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" accesskey="B">
	<!--<input id="btnBackDisabled" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnLoans_dis.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" disabled tabindex="-1"> -->
</div>

<div id="divConfirm" style="TOP: 0px; LEFT: 0px;HEIGHT:1px; WIDTH: 1px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE">
	<input id="btnConfirm" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnConfirm.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton">
	<input id="btnConfirmDisabled" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnConfirm_dis.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" disabled tabindex="-1">
</div>

<div id="divFeatures" style="TOP: 0px; LEFT: 0px;HEIGHT:1px; WIDTH: 1px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE">
	<input id="btnFeatures" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnFeatures.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton">
	<input id="btnFeaturesDisabled" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnFeatures_dis.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" disabled tabindex="-1">
</div>

<div id="divArrears" style="TOP: 0px; LEFT: 0px;HEIGHT:1px; WIDTH: 1px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE">
	<input id="btnArrears" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnArrears.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton">
	<input id="btnArrearsDisabled" value="" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 72px; HEIGHT: 31px; BACKGROUND-IMAGE: url(images/btnArrears_dis.gif); VISIBILITY: hidden; POSITION: ABSOLUTE" class="msgMainButton" disabled tabindex="-1">
</div>

<%
/*
<div id="msglogo" style="HEIGHT: 32px; LEFT: 461px; POSITION: absolute; TOP: -8px; WIDTH: 20px; Z-INDEX: 45"> 
	<img height="40" src="images/dolph.gif" width="152">
</div>
*/
%>

<script LANGUAGE="JScript">
<!--
var divButton;
var btnEnable;
var btnDisable;
	
function private_SetFieldVariables(sButton)
{
<%	// Central function to set the variables which will be used to manipulate the buttons required
	// (Not meant for use outside this file)
	// When adding or deleting buttons in this file this switch must be changed
%>	var bIsFound = true;
	var sToUpperButton = sButton.toUpperCase();

	switch(sToUpperButton)
	{
		case "SUBMIT":
			divButton	= divSubmit;
			btnEnable	= btnSubmit;
			btnDisable	= btnSubmitDisabled;
		break;
		case "CANCEL":
			divButton	= divCancel;
			btnEnable	= btnCancel;
			btnDisable	= btnCancelDisabled;
		break;
		case "UNDO":
			divButton	= divUndo;
			btnEnable	= btnUndo;
			btnDisable	= btnUndoDisabled;
		break;
		case "ANOTHER":
			divButton	= divAnother;
			btnEnable	= btnAnother;
			btnDisable	= btnAnotherDisabled;
		break;
		case "LOANS":
			divButton	= divLoans;
			btnEnable	= btnLoans;
			btnDisable	= btnLoansDisabled;
		break;
		case "PREVIOUS":
			divButton	= divPrevious;
			btnEnable	= btnPrevious;
			btnDisable	= btnPreviousDisabled;
		break;
		case "NEXT":
			divButton	= divNext;
			btnEnable	= btnNext;
			btnDisable	= btnNextDisabled;
		break;
		case "FINISH":
			divButton	= divFinish;
			btnEnable	= btnFinish;
			btnDisable	= btnFinishDisabled;
		break;
		case "BACK":
			divButton	= divBack;
			btnEnable	= btnBack;
			//btnDisable	= btnBackDisabled;
		break;
		case "CONFIRM":
			divButton	= divConfirm;
			btnEnable	= btnConfirm;
			btnDisable	= btnConfirmDisabled;
		break;

		case "ARREARS":
			divButton	= divArrears;
			btnEnable	= btnArrears;
			btnDisable	= btnArrearsDisabled;
		break;

		case "FEATURES":
			divButton	= divFeatures;
			btnEnable	= btnFeatures;
			btnDisable	= btnFeaturesDisabled;
		break;

		default:
			divButton	= null;
			btnEnable	= null;
			btnDisable	= null;
			bIsFound	= false;
		break;
	}

	return bIsFound;
}
	
function ShowMainButtons(sButtonArray)
{
<%	// Show all the buttons specified in the list
	// Starting position for first button
%>	var nLeft = 0;

	if(sButtonArray != null)
		for(var nLoop = 0; nLoop < sButtonArray.length; nLoop++)
		{
			if(private_SetFieldVariables(sButtonArray[nLoop]))
			{
<%				// Make the button group visible at the correct position
				// and increment that position ready for the next button
%>				divButton.style.visibility = "visible";
				btnEnable.style.visibility = "visible";
				divButton.style.left = nLeft + "px";
				nLeft += 75; 
			}
		}
}

function EnableMainButton(sButton)
{
<%	// Enable the specified button
	// (show the button to which the onclick event should be attached)
%>	private_SetButtonState(sButton, true);
}
	
function DisableMainButton(sButton)
{
<%	// Disable the specified button
	// (hide the button to which the onclick event should be attached and show the grey one instead)
%>	private_SetButtonState(sButton, false);
}

function MoveMainButtonsVertical(sButtonArray, nRelPos)
{
<%	// Move the main buttons vertically by the given amount
%>	for(var nLoop = 0; nLoop < sButtonArray.length; nLoop++)
		if(private_SetFieldVariables(sButtonArray[nLoop])) divButton.style.top = nRelPos;
}
	
function private_SetButtonState(sButton, bIsEnabled)
{
<%	// Show the appropriate button image
%>	if(private_SetFieldVariables(sButton))
		if(divButton.style.visibility == "visible")
		{
			if(bIsEnabled)
			{
				btnEnable.style.visibility	= "visible";
				btnDisable.style.visibility	= "hidden";
			}
			else
			{
				btnEnable.style.visibility	= "hidden";
				btnDisable.style.visibility	= "visible";
			}
		}
}
-->
</script>
 