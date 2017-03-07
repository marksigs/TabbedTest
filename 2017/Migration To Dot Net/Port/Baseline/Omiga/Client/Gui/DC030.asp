<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<% /* 
Workfile:      DC030.asp
Copyright:     Copyright © 1999 Marlborough Stirling
Description:   PersonalDetails
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		18/11/99	Created
JLD		10/12/1999	DC/025 - removed customer salutation
JLD		13/12/1999	DC/042 - moved screen around a bit
					DC/033 - Added Customer Name field
					DC/038 - corrected routing to buttons
JLD		14/12/1999	DC/034 - make calling frame for dc032 bigger
JLD		16/12/1999	SYS0076 - removed idHasDependents and missing out of DC040 if no dependents
					SYS0064 - removed car owner question.
JLD		17/12/1999	SYS0079 - check validity of date of birth
AD		30/01/2000	Rework
AY		14/02/00	Change to msgButtons button types
IW		22/02/00	SYS0290 - Change Scriptlet: scMathFunctions.htm to scMathFunctions.asp
IW		23/02/00	SYS0092 - Added CustomerNameRefresh() JavaScript (with calls in Forename & Surname Blur events)
IW		24/02/00	SYS0162 - Changed Resident indicator back end mapping
IW		22/03/00	SYS0112 - Misc Cosmetic changes 
IW		24/03/00	Homezone - Added Pref Name & SufficientRepaymentIndicator
AY		30/03/00	New top menu/scScreenFunctions change
MC		03/05/00	SYS0583 - Save Mailshot indicator value
MH      04/05/00    SYS0636 - Tab order
MH      05/05/00	SYS0577 - Screen layout
SR		15/05/00    SYS0189 - Age should be greater than 17 and less than 100
BG		17/05/00	SYS0752 Removed Tooltips
IW		25/05/00	SYS0776 Added EmailPreffered option
MC		21/06/00	SYS0845 Make various fields mandatory.
BG		25/07/00	SYS0971 - Made customer name field longer to handle max length of customer name
IK		04/03/01	critical data test
CL		12/03/01	SYS1920 Read only functionality added
SA		23/05/01	SYS0907 Changed maxlength of surname from 35 to 40	
					SYS1070 Altered CustomerNameRefresh so preferred name has no impact on name displayed.							
					SYS1053 Changed maxlength of Mothers maiden name from 35 to 40	
							optSufficientRepaymentIndicatorYes had incorrect Name causing inproper highlighting of label.
							Maxlength of txtElectoralRoll changed from 10 to 4 and mask added.
							New function - CheckNames added to catch first char of surname + forenames being numeric.
TW      09/10/02    SYS5115 Modified to incorporate client validation
DF		19/06/02	BMIDS00077 - Altering code to make this compatible with
							with version 7.0.2 of Omiga 4, the following comment was in Core. 
							SYS4727 Use cached versions of frame functions

MC		19/04/04	BMIDS751	Popup dialog Height incr by 10px
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		Description
MV		30/05/2002	BMIDS00013 - BM044 - Amended New Field "No Of Dependants" and Modified btnSubmit.onclick()
MV		30/05/2002	BMIDS00013 - BM045 - Disabled Health and smoker Radio Buttons. 
MV		07/06/2002	BMIDS00013 - BM044 - Modified Routing Screen in btnSubmit.onclick()
MV		10/06/2002	BMIDS00013 - BM044 - Modified Routing Screen in btnSubmit.onclick()
MV		11/06/2002	BMIDS00013 - BM044 - Code Review Errors
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MV		20/08/2002	BMIDS00204	Restricted the NoOfDependents to two Digits
MDC		27/08/2002	BMIDS00336 - BM062 Credit Check and Bureau Download
CL		16/10/2002	BMIDS00639 - Change to txtDateOfBirth.onblur
MV		31/10/2002	BMIDS00355	Modified HTML
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client
GD		17/11/2002	BMIDS00376 Ensure readonly correctly passed to contact details popup.
JR		23/01/2003	BM0271	Duplicate ReadOnly action for non-processing unit.
BS		28/01/2003	BM0272	Initialise MemberOfStaff for additional applicants
HMA     17/09/2003  BM0063  Amend HTML text for radio buttons
PJO     18/09/2003  BMIDS00619 Comment out hidden fields
KRW      25/09/03   BM0063 Screen alignment corrections
MC		24/05/2004	BMIDS747	DC035 New Screen Added. Reference to Verification summary screen
								Form element added to route the screen.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History

Prog    Date		AQR		Description
MF		22/07/2005			IA_WP01 process flow change. Addition of special needs combo.
							Removal of Mothers maiden name. 
GHun	02/08/2005	MAR14	Prevent combos from being hidden when using the menu
MF		08/08/2005	MAR20	Populate new fields Time at Bank years & Time at bank months
MF		09/09/2005	MAR20	Enabled certain fields where data is null: Marital Status,
								Nationality, Resident in UK, Member of Staff, No Of Dependants,
								Life cover question
MF		22/09/2005	MAR19	Modifications to fix multiple client bug, setting readonly state
							state of controls.
JJ		07/10/2005  MAR119  txtNumberOfDependants maxlength changed to 1.							
PE		25/01/2006	MAR939	Parameterised the disabling of the customer screens
PSC		30/01/2006	MAR939	Correct logic
GHun	31/01/2006	MAR1158	Removed redundant OtherSystemReadOnlyCustomer logic
GHun	10/04/2006	MAR1602	Display OtherSystemCustomerNumber
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History

Prog    Date		AQR		Description
AW		08/05/2006	MAR390  Make Other title mandatory
PB		11/05/2006	EP529	Merged MAR1274 - Enabled MiddleName field where data is null: txtSecondForename
PB		23/05/2006	EP586	Fields re-enabling when viewing second applicant - changed Initialise() to window.onload()
PB		25/05/2006	EP612	Replaced window.onload() with Initialse() and added re-check for readonly fields.
							Also applied EP586-612 to the Cancel button (not previously reported)
HMA     06/07/2006  EP903   Use application date for age verification.	
                            Pass age on application through to Critaical Data Check.
PB		14/07/2006	EP995	Enables marketing fields, i.e.: "Do you wish to receive this information?"
MAH		16/11/2006	E2CR35	Add Residency questions and Housing benefit
MCh		28/11/2006	EP2_127 Disable txtNumberOfDependants if DCDependantsAtApplicationLevel = True
MAH		11/12/2006	EP2_353	Amended validation of Housing Benefit option
PSC		11/01/2007	EP2_741	Add Mother's Maiden Name back in
INR		22/01/2007	EP2_677 move HOUSINGBENEFIT to new screen DC037.
AShaw	25/01/2007	EP2_1007 Resize DC037 screen (smaller).
AShaw	29/01/2007  EP2_610 - Default cboSpecialNeeds combo to SELECT.
AShaw	29/01/2007	EP2_1007 - Benefit ques form is mandatory.
AShaw	19/02/2007	EP2_1474  Add m_blnReadOnly test to CommitData method.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 
*/
%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object id="scClientFunctions"
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex="-1" 
data="scClientFunctions.asp" width="1" height="1" type="text/x-scriptlet"
viewastext></object>
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript" type="text/javascript"></script>

<% /* the following has been removed as per Core V7 - DPF 19/06/02 - BMIDS00077
<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scMathFunctions.asp height=1 id=scMathFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
*/ %>
<%/* FORMS */%>
<form id="frmToDC010" method="post" action="DC010.asp" style="DISPLAY: none"></form>
<form id="frmToDC020" method="post" action="DC020.asp" style="DISPLAY: none"></form>
<form id="frmToDC031" method="post" action="DC031.asp" style="DISPLAY: none"></form>
<form id="frmToDC032" method="post" action="DC032.asp" style="DISPLAY: none"></form>
<form id="frmToDC033" method="post" action="DC033.asp" style="DISPLAY: none"></form>
<form id="frmToDC035" method="post" action="DC035.asp" style="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="gn300.asp" style="DISPLAY: none"></form>
<form id="frmToDC060" method="post" action="DC060.asp" style="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span><!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange" year4>
<div id="divCustomerName" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 40px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 15px" class="msgLabel">
		Customer Name
		<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
			<input id="txtCustName" name="CustomerName" maxlength="70" style="WIDTH: 430px; POSITION: absolute" class="msgTxt">
		</span>
	</span>
</div>
<div id="divDetailsOne" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 110px; HEIGHT: 180px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel">
		<strong>Personal Details</strong>
		<span id="spnOtherSysCustomerNo" style="LEFT: 376px; POSITION: absolute; TOP: 4px" class="msgLabel">
			Customer Number
			<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
				<input id="txtOtherSysCustomerNo" type="text" maxlength="12" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
			</span>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 30px" class="msgLabel">
		Surname
		<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
			<input id="txtSurname" name="Surname" maxlength="40" style="WIDTH: 150px; POSITION: absolute" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 252px; POSITION: absolute; TOP: 30px" class="msgLabel">
		Contact Details
		<span style="LEFT: 90px; POSITION: absolute; TOP: -8px">
			<input id="btnContact" type ="button" style="WIDTH: 26px; HEIGHT: 26px" class="msgDDButton">
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 54px" class="msgLabel">
		Forenames
		<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
			<input id="txtFirstForename" name="FirstForename" maxlength="30" style="WIDTH: 150px; POSITION: absolute" class="msgTxt">
		</span>

		<span style="LEFT: 248px; POSITION: absolute; TOP: -3px">
			<input id="txtSecondForename" name="SecondForename" maxlength="30" style="WIDTH: 150px; POSITION: absolute" class="msgTxt">
		</span>

		<span style="LEFT: 416px; POSITION: absolute; TOP: -3px">
			<input id="txtOtherForenames" name="OtherForenames" maxlength="30" style="WIDTH: 150px; POSITION: absolute" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 78px" class="msgLabel">
		Preferred Name
		<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
			<input id="txtPreferredname" name="PreferredName" maxlength="30" style="WIDTH: 150px; POSITION: absolute" class="msgTxt">
		</span>
	</span>	

	<span style="LEFT: 4px; POSITION: absolute; TOP: 102px" class="msgLabel">
		Title
		<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
			<select id="cboTitle" name="Title" style="WIDTH: 150px" class="msgCombo" onchange="cboTitle_OnChange()" menusafe="true">
			</select>
		</span>

		<span id="spnOtherTitle" style="LEFT: 248px; POSITION: absolute; TOP: -3px">
			<input id="txtOtherTitle" name="OtherTitle" maxlength="20" style="WIDTH: 150px; POSITION: absolute" class="msgTxt">
		</span>
	</span>



	<span style="LEFT: 4px; POSITION: absolute; TOP: 130px" class="msgLabel">
		Date of Birth
		<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
			<input id="txtDateOfBirth" name="DateOfBirth" maxlength ="10" style="WIDTH: 70px;" class="msgTxt" >
		</span>
	</span>


	<span style="LEFT: 4px; POSITION: absolute; TOP: 130px">
		<span style="LEFT: 204px; POSITION: absolute; TOP: 0px" class="msgLabel">
			Age
			<span style="LEFT: 44px; POSITION: absolute; TOP: -3px">
				<input id="txtAge" name="Age" maxlength="3" style="WIDTH: 40px; POSITION: absolute" class="msgReadOnly" readonly tabindex="-1">
			</span>
		</span>

		<span style="LEFT: 376px; POSITION: absolute; TOP: 0px" class="msgLabel">
			Gender
			<span style="LEFT: 43px; POSITION: absolute; TOP: -3px">
				<select id="cboGender" name="Gender" style="WIDTH: 154px" class="msgCombo" menusafe="true">
				</select>
			</span>
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 153px" class="msgLabel">
		Mother's<br>Maiden Name
		<span style="LEFT: 80px; POSITION: absolute; TOP: 3px">
			<input id="txtMotherMaidenName" name="MotherMaidenName" maxlength="40" style="WIDTH: 150px; POSITION: absolute" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 188px" class="msgLabel">
		Special Needs
		<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
			<select id="cboSpecialNeeds" style="WIDTH: 120px" disabled readonly class="msgReadOnlyCombo" menusafe="true">
			</select>
		</span>
	</span>
	<span style="LEFT: 260px; POSITION: absolute; TOP: 185px" class="msgLabel">
		Additional Personal Details
		<span style="LEFT: 160px; POSITION: absolute; TOP: -8px">
			<input id="btnPersonalDetails" type ="button" style="WIDTH: 26px; HEIGHT: 26px" class="msgDDButton">
		</span>
	</span>

</div>

<div id="divDetails2" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 331px; HEIGHT: 190px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Nationality
		<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
			<select id="cboNationality" name="Nationality" style="WIDTH: 120px" class="msgCombo" menusafe="true">
			</select>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 38px" class="msgLabel">
		Marital Status
		<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
			<select id="cboMaritalStatus" name="MaritalStatus" style="WIDTH: 120px" class="msgCombo" menusafe="true">
			</select>
		</span>
	</span>
	
	<% /* MV - 24/05/2002 - BMIDS00013 - BM045 - Disable RelationShip by making visibility "HIDDEN" */ %>
	<% /* PJO 18/05/2003 - BMIDS00619 - Comment out unused fields
	<span id="spnRelationship" style="LEFT: 4px; POSITION: absolute; TOP: 50px;VISIBILITY=hidden " class="msgLabel">
		Relationship to<BR>Applicant2
		<span style="LEFT: 110px; POSITION: absolute; TOP: 3px">
			<select id="cboRelationship" name="ApplicantRelationship" style="WIDTH: 120px" class="msgCombo">
			</select>
		</span>
	</span>
	*/ %>
	
	<% /* MV - 25/05/2002 - BMIDS00013 - BM044 - Amended a New Field "No Of Dependants" */ %> 
	<span style="LEFT: 4px; POSITION: absolute; TOP: 64px" class="msgLabel">
		No of Dependants
		<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
			<input id="txtNumberOfDependants" name="NumberOfDependants" maxlength="1" style= "WIDTH: 120px" class="msgTxt">
		</span> 
	</span>	
	
	<% /* MV - 24/05/2002 - BMIDS00013 - BM045 - Disable Health Remove Health and smoker questions by making visibility "HIDDEN" */ %>
	<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
	<% /* PJO 18/05/2003 - BMIDS00619 - Comment out unused fields
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 82px ;VISIBILITY=hidden " class="msgLabel">
		Smoker?
		<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
			<input id="optSmokerYes" name="SmokerGroup" type="radio" value="1"><label for="optSmokerYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="optSmokerNo" name="SmokerGroup" type="radio" value="0"><label for="optSmokerNo" class="msgLabel">No</label>
		</span> 
	</span>		

	<span style="LEFT: 4px; POSITION: absolute; TOP: 112px;VISIBILITY=hidden" class="msgLabel">
		Good Health?
		<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
			<input id="optGoodHealthYes" name="GoodHealthGroup" type="radio" value="1"><label for="optGoodHealthYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="optGoodHealthNo" name="GoodHealthGroup" type="radio" value="0"><label for="optGoodHealthNo" class="msgLabel">No</label>
		</span> 
	</span>
	*/ %>
	<% /* MF MARS20 WP06 Hide field */ %>
	<span style="DISPLAY:none; LEFT: 4px; WIDTH: 230px; POSITION: absolute; TOP: 92px">
		Do you or will you have sufficient life cover in place at the time of taking out the mortgage?
		<span style="LEFT: 110px; WIDTH: 50px; POSITION: absolute; TOP: 32px; HEIGHT: 20px">
			<input id="optSufficientRepaymentIndicatorYes" name="SufficientRepaymentIndicatorGroup" readonly disabled type="radio" value="1"><label for="optSufficientRepaymentIndicatorYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 170px; WIDTH: 40px; POSITION: absolute; TOP: 32px; HEIGHT: 21px">
			<input id="optSufficientRepaymentIndicatorNo" name="SufficientRepaymentIndicatorGroup" readonly disabled type="radio" value="0"><label for="optSufficientRepaymentIndicatorNo" class="msgLabel">No</label>
		</span> 
	</span>
	<% /* MF MARS20 WP06 New fields. */ %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 90px" class="msgLabel">
		Time at Bank
		<span style="LEFT: 110px; TOP:-3px; POSITION: absolute;">
			<input id="txtBankYears" type="text" maxlength="2" style= "WIDTH: 20px" class="msgTxt">
		</span>
		<span style="LEFT: 133px; TOP:0px; POSITION: absolute;">
			Years
		</span>
		<span style="LEFT: 171px; TOP:-3px; POSITION: absolute;">
			<input id="txtBankMonths" type="text" maxlength="2" style= "WIDTH: 20px" class="msgTxt" NAME="txtBankMonths">
		</span>
		<span style="LEFT: 195px; TOP:0px; POSITION: absolute;">
			Months
		</span>
	</span>
	
	<% /*	BMIDS00336 MDC 27/08/2002 */ %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 112px" class="msgLabel">
		Credit Search
		<span style="LEFT: 37px; POSITION: relative; TOP: 0px">
			<input id="optCreditSearchYes" name="CreditSearchGroup" readonly disabled type="radio" value="1"><label for="optCreditSearchYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 55px; POSITION: relative; TOP: 0px">
			<input id="optCreditSearchNo" name="CreditSearchGroup" readonly disabled type="radio" value="0"><label for="optCreditSearchNo" class="msgLabel">No</label>
		</span> 
	</span>
	<% /*	BMIDS00336 MDC 27/08/2002 - End */ %>

	<span style="LEFT: 260px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Year added to Electoral Roll
		<span style="LEFT: 163px; POSITION: absolute; TOP: -3px">
			<input id="txtElectoralRoll" name="ElectoralRoll" maxlength="4" style="WIDTH: 100px; POSITION: absolute" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 260px; POSITION: absolute; TOP: 38px" class="msgLabel">
		Resident in the U.K.?
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="optUKResidentYes" name="UKResidentGroup" type="radio" value="1"><label for="optUKResidentYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 230px; POSITION: absolute; TOP: -3px">
			<input id="optUKResidentNo" name="UKResidentGroup" type="radio" value="0"><label for="optUKResidentNo" class="msgLabel">No</label>
			<%/* E2CR35 Add residency button*/%>
			<span style="LEFT: 50px; POSITION: absolute; TOP: -3px">
				<input id="btnResidency" name="Residency" type="button" style="WIDTH: 26px; HEIGHT: 26px" class="msgDDButton"> 
			</span>
		</span>
	</span>		

	<span style="LEFT: 260px; POSITION: absolute; TOP: 64px" class="msgLabel">
		Member of Staff?
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="optStaffYes" name="StaffGroup" type="radio" value="1"><label for="optStaffYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 230px; POSITION: absolute; TOP: -3px">
			<input id="optStaffNo" name="StaffGroup" type="radio" value="0"><label for="optStaffNo" class="msgLabel">No</label>
		</span> 
	</span>		

	<span style="LEFT: 260px; POSITION: absolute; TOP: 87px" class="msgLabel">
		<strong>It is important to keep customers informed of new services and products which we consider may be relevant to their needs.</strong>
	</span>

	<span style="LEFT: 260px; POSITION: absolute; TOP: 131px" class="msgLabel">
		Do you wish to receive this information?
		<span style="LEFT: 200px; POSITION: absolute; TOP: -3px"><% /*
			PB 14/07/2006 EP995 Enabled radio buttons */ %>
			<input id="optMailshotRequiredYes" name="MailshotGroup" type="radio" value="1"><label for="optMailshotRequiredYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 260px; POSITION: absolute; TOP: -3px">
			<input id="optMailshotRequiredNo" name="MailshotGroup" type="radio" value="0"><label for="optMailshotRequiredNo" class="msgLabel">No</label><% /*
			PB EP995 End */ %>
		</span> 
	</span>

	<span style="LEFT: 259px; POSITION: absolute; TOP: 160px">
		<input id="btnVerification" value="Verification" type="button" style="WIDTH: 100px" class="msgButton">
	</span>

	<span style="LEFT: 366px; POSITION: absolute; TOP: 160px">
		<input id="btnAlias" value="Previous Names" type="button" style="WIDTH: 100px" class="msgButton">
	</span>
</div>
</form><!-- Main Buttons -->
<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 520px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!-- File containing field attributes (remove if not required) --><!--  #include FILE="attribs/DC030attribs.asp" --><!-- Specify Code Here -->
<script language="JScript" type="text/javascript">
<!--
var m_sMetaAction = "";
var m_sContext = null;
var m_sReadOnly	= "";	
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_sCustomer1Number = "";
var m_sCustomer1Version = "";
var m_sCustomer2Number = "";
var m_sCustomer2Version = "";
var m_sCustomer3Number = "";
var m_sCustomer3Version = "";
var m_sCustomer4Number = "";
var m_sCustomer4Version = "";
var m_sCustomer5Number = "";
var m_sCustomer5Version = "";
var m_sCurrentCustomerNumber = "";
var m_sCurrentCustomerVersionNumber = "";
var m_sDateOfBirth = "";
<% /* PJO 18/05/2003 - BMIDS00619 - Unused
var m_nOriginalRelationshipSel = -1;
*/ %>
var m_sNextCustNum = "";
var m_sNextCustVersion = "";
var m_sSurname = "";
var m_sFirstForename = "";
var m_sPreferredname = "";
var CustomerXML = null;
var scScreenFunctions;
var m_blnReadOnly = false;
var m_sProcessInd = ""; //JR BM0271
var m_sApplicationDate = ""; <% /* EP903 */ %>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	//next line replaced by line below as per Core V7 - DPF 19/06/02 - BMIDS00077
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Personal Details","DC030",scScreenFunctions);

	RetrieveContextData();

	//This function is contained in the field attributes file (remove if not required)
	SetMasks();
	Validation_Init();
	
	Initialise();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC030");
		
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();	
	
	// EP2_1007 - Can't make button Mandatory in Attribs so Just make it Red (with just a subtle hint of Red).
	frmScreen.btnPersonalDetails.parentElement.parentElement.style.color = "red";
	
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function frmScreen.txtDateOfBirth.onblur()
{
	//if the value has changed set the age field
	if(m_sDateOfBirth != frmScreen.txtDateOfBirth.value)
	{
		m_sDateOfBirth = frmScreen.txtDateOfBirth.value;
		frmScreen.txtAge.value = "";
		var nAge = CalculateCustomerAge();
		
		<% /* EP903 Calculate age on the date the application was received */ %>
		var nAppAge = CalculateCustomerAgeOnApplication();
		
		<% /* EP903 Verify age against max and min global parameters */ %>
		if(nAppAge != -1)
		{   
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			var sMinAge = XML.GetGlobalParameterAmount(document,"MinimumAge");
			var sMaxAge = XML.GetGlobalParameterAmount(document,"MaximumAge");
			
			if(nAppAge < sMinAge ) alert("Age on Application Date is below accepted age");
			else if (nAppAge > sMaxAge) alert("Age on Application Date is greater than accepted age");
			else if (nAge != -1)
				frmScreen.txtAge.value = nAge.toString();
		}
	}
}

function CustomerNameRefresh()
{
	// SYS1070 Preferred name should have no impact on customer name displayed
	//if (frmScreen.txtPreferredname.value == "")
	//{
		var sCustomerName = frmScreen.txtFirstForename.value+' '+frmScreen.txtSurname.value;
	//}
	//else
	//{
	//	var sCustomerName = frmScreen.txtPreferredname.value+' '+frmScreen.txtSurname.value;
	//}	
		scScreenFunctions.SetContextParameter(window,"idCustomerName" + GetNumberOfCustomer(m_sCurrentCustomerNumber), sCustomerName);
		frmScreen.txtCustName.value  = sCustomerName
}


function frmScreen.txtSurname.onblur()
{
	CustomerNameRefresh()	
}

function frmScreen.txtFirstForename.onblur()
{
	CustomerNameRefresh()		
}

function frmScreen.txtPreferredname.onblur()
{
	CustomerNameRefresh()		
}

function frmScreen.cboTitle.onchange()
{
	DefaultGenderFromTitle();
	HideOtherTitle();
}

function frmScreen.btnContact.onclick()
{
	var sReturn = null;
	var ArrayArguments = CustomerXML.CreateRequestAttributeArray(window, null)

<% // GD BMIDS00376 START %>
	ArrayArguments[4] = (m_blnReadOnly || m_sReadOnly);
<% // GD BMIDS00376 END %>

	ArrayArguments[5] = CustomerXML.XMLDocument.xml;
	ArrayArguments[6] = scScreenFunctions.GetContextCustomerName(window, m_sCurrentCustomerNumber);

	sReturn = scScreenFunctions.DisplayPopup(window, document, "DC032.asp", ArrayArguments, 630, 395);

	if(sReturn != null)
	{
		FlagChange(sReturn[0]);
		var sXML = sReturn[1];
		CustomerXML.LoadXML(sXML);
	}
}

function frmScreen.btnVerification.onclick()
{
	var bSuccess = CommitData();
	if(bSuccess == true)
	{
		// set context info
		scScreenFunctions.SetContextParameter(window,"idCustomerNumber",m_sCurrentCustomerNumber);
		scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber",m_sCurrentCustomerVersionNumber);
		//frmToDC031.submit();
		frmToDC035.submit();
		
	}
}
	
function frmScreen.btnAlias.onclick()
{
	var bSuccess = CommitData();
	if(bSuccess == true)
	{
		// set context info
		scScreenFunctions.SetContextParameter(window,"idCustomerNumber",m_sCurrentCustomerNumber);
		scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber",m_sCurrentCustomerVersionNumber);
		frmToDC033.submit();
	}
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sContext = scScreenFunctions.GetContextParameter(window,"idProcessContext",null);
	m_sReadOnly	= scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","APP0001");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","2");
	m_sCustomer1Number = scScreenFunctions.GetContextParameter(window,"idCustomerNumber1",null);
	m_sCustomer1Version = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber1",null);
	m_sCustomer2Number = scScreenFunctions.GetContextParameter(window,"idCustomerNumber2",null);
	m_sCustomer2Version = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber2",null);
	m_sCustomer3Number = scScreenFunctions.GetContextParameter(window,"idCustomerNumber3",null);
	m_sCustomer3Version = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber3",null);
	m_sCustomer4Number = scScreenFunctions.GetContextParameter(window,"idCustomerNumber4",null);
	m_sCustomer4Version = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber4",null);
	m_sCustomer5Number = scScreenFunctions.GetContextParameter(window,"idCustomerNumber5",null);
	m_sCustomer5Version = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber5",null);
	m_sProcessInd = scScreenFunctions.GetContextParameter(window,"idProcessingIndicator",null); //JR BM0271
	m_sApplicationDate = scScreenFunctions.GetContextParameter(window,"idApplicationDate","");  <%/* EP903 */%>

	if(m_sMetaAction == "FromDC031" || m_sMetaAction == "FromDC033" || m_sContext == "CompletenessCheck")
	{
		//Get the current customer
		m_sCurrentCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber",null);
		m_sCurrentCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber",null);
	}	
}

function PopulateCombos()
{
	var XMLTitle = null;
	var XMLSex = null;
	var XMLNationality = null;
	var XMLMaritalStatus = null;
	<% /* PJO 18/05/2003 - BMIDS00619 - Unused
	var XMLApplicantRelationship = null;
	*/ %>
	//next line replaced by line below as per Core V7 - DPF 19/06/02 - BMIDS00077
	//var XML = new scXMLFunctions.XMLObject();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("Title", "Sex", "Nationality", "MaritalStatus", "SpecialNeeds"<% /*, "ApplicantRelationship"*/ %>);

	if(XML.GetComboLists(document, sGroupList))
	{
		XMLTitle = XML.GetComboListXML("Title");
		XMLSex = XML.GetComboListXML("Sex");
		XMLNationality = XML.GetComboListXML("Nationality");
		XMLMaritalStatus = XML.GetComboListXML("MaritalStatus");
		XMLSpecialNeeds = XML.GetComboListXML("SpecialNeeds");
		<% /* PJO 18/05/2003 - BMIDS00619 - Unused
		XMLApplicantRelationship = XML.GetComboListXML("ApplicantRelationship");
		*/ %>
		var blnSuccess = true;
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document, frmScreen.cboTitle,XMLTitle,true);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document, frmScreen.cboGender,XMLSex,true);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document, frmScreen.cboNationality,XMLNationality,true);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document, frmScreen.cboMaritalStatus,XMLMaritalStatus,true);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document, frmScreen.cboSpecialNeeds,XMLSpecialNeeds,true);
		<% /* PJO 18/05/2003 - BMIDS00619 - Unused
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document, frmScreen.cboRelationship,XMLApplicantRelationship,true);
		*/ %>
		if(!blnSuccess)
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}

	XML = null;		
}

<% /* PJO 18/05/2003 - BMIDS00619 - Unused
function SetUpRelationship()
{
	if(m_sCurrentCustomerNumber == m_sCustomer1Number && m_sCurrentCustomerVersionNumber == m_sCustomer1Version &&
	  (m_sCustomer2Number != "" && m_sCustomer2Version != ""))
	{
		//We are the first customer and there is a second customer so 
		//show the Relationship combo and set it
		
		//MV - 30/05/2002 - BMIDS00013 - BM044 - Hide Relationship field
		//scScreenFunctions.ShowCollection(spnRelationship);

		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window, null)
		XML.CreateActiveTag("CUSTOMERRELATIONSHIP");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("OWNERCUSTOMERNUMBER", m_sCurrentCustomerNumber);
		XML.CreateTag("OWNERCUSTOMERVERSIONNUMBER", m_sCurrentCustomerVersionNumber);
		// 		XML.RunASP(document, "GetCustomerRelationship.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "GetCustomerRelationship.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}


		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = XML.CheckResponse(ErrorTypes);

		if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
		{
			//Use this record to set the combo
			frmScreen.cboRelationship.value = XML.GetTagText("RELATIONSHIPTOOWNER");
			//save initial value 
			m_nOriginalRelationshipSel = frmScreen.cboRelationship.selectedIndex;
		}

		ErrorTypes = null;
		ErrorReturn = null;				
		XML = null;
	}
	else
		scScreenFunctions.HideCollection(spnRelationship);
}
*/ %> 

function GetCustomer()
{
	//next line replaced by line below as per Core V7 - DPF 19/06/02 - BMIDS00077
	//CustomerXML = new scXMLFunctions.XMLObject();
	CustomerXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	with (CustomerXML)
	{
		CreateRequestTag(window, null)
		CreateActiveTag("CUSTOMERVERSION");
		CreateTag("CUSTOMERNUMBER", m_sCurrentCustomerNumber);
		CreateTag("CUSTOMERVERSIONNUMBER", m_sCurrentCustomerVersionNumber);
		RunASP(document, "GetPersonalDetails.asp");
		IsResponseOK();
	}
}

function CalculateCustomerAge()
{
	var nAge = -1;
	var dteBirthdate = scScreenFunctions.GetDateObject(frmScreen.txtDateOfBirth);
	if(dteBirthdate != null)
	{
		<% /* MO - BMIDS00807 */ %>
		var dteToday = scScreenFunctions.GetAppServerDate(true);

		<% /* var dteToday = new Date(); */ %>
		//next line replaced by line below as per Core V7 - DPF 19/06/02 - BMIDS00077
		//nAge = scMathFunctions.GetYearsBetweenDates(dteBirthdate, dteToday);
		nAge = top.frames[1].document.all.scMathFunctions.GetYearsBetweenDates(dteBirthdate, dteToday);

	}

	return(nAge);
}

<% /* EP903 New function */ %>
function CalculateCustomerAgeOnApplication()
{
	var nAppAge = -1;
	var dteBirthdate = scScreenFunctions.GetDateObject(frmScreen.txtDateOfBirth);
	if(dteBirthdate != null)
	{
		var dtApplicationDate = scScreenFunctions.GetDateObjectFromString(m_sApplicationDate);
		nAppAge = top.frames[1].document.all.scMathFunctions.GetYearsBetweenDates(dteBirthdate, dtApplicationDate);
	}

	return(nAppAge);
}

function HideOtherTitle()
{
	var selIndex = frmScreen.cboTitle.selectedIndex;
	if(selIndex != -1 && scScreenFunctions.IsOptionValidationType(frmScreen.cboTitle,selIndex,"O"))
		scScreenFunctions.ShowCollection(spnOtherTitle);
	else
		scScreenFunctions.HideCollection(spnOtherTitle);
}

function DefaultGenderFromTitle()
{
	var selIndex = frmScreen.cboTitle.selectedIndex;
	if(selIndex != -1 && scScreenFunctions.IsValidationType(frmScreen.cboTitle,"M"))
		frmScreen.cboGender.value = "1";
	else
	{
		if(selIndex != -1 && scScreenFunctions.IsValidationType(frmScreen.cboTitle,"F"))
			frmScreen.cboGender.value = "2";
		else
			//default back to <SELECT>
			frmScreen.cboGender.selectedIndex = 0;
	}
}

function CommitData()
{
	// EP2_1474 Add m_blnReadOnly test.
	if( m_sReadOnly == "1" || m_sProcessInd == "0" || m_blnReadOnly == true)
		return true;

	var bSuccess = frmScreen.onsubmit();
	if (bSuccess)
	{
		<% /* PJO BMIDS00619 - Unused
		if(scScreenFunctions.GetRadioGroupValue(frmScreen,"SmokerGroup") == null)
		{
			alert("Please enter smoker status");
			return false;
		}
		else if(scScreenFunctions.GetRadioGroupValue(frmScreen,"GoodHealthGroup") == null)
		{
			alert("Please enter health status");
			return false;
		}
		else */ %>
		<% /* MF MAR019 WP01 Fields now readonly
		if(scScreenFunctions.GetRadioGroupValue(frmScreen,"UKResidentGroup") == null)
		{
			alert("Please enter UK Residency status");
			return false;
		}
		else if(scScreenFunctions.GetRadioGroupValue(frmScreen,"StaffGroup") == null)
		{
			alert("Please enter Staff status");
			return false;
		}
		*/ %>
		<% /* MF MAR019 Make DOB Mandatory */ %>		
		if (frmScreen.txtDateOfBirth.value == ""){
			alert("Please enter Date of Birth");
			frmScreen.txtDateOfBirth.focus();
			return false;
		}

		CustomerXML.SelectTag(null,"CUSTOMERVERSION");

		// EP2_1007 - Alert if Housing Benefit NOT answered.
		if (CustomerXML.GetTagText("HOUSINGBENEFIT") == null || CustomerXML.GetTagText("HOUSINGBENEFIT") == "" )
		{
			alert("Please complete Additional Personal Details");
			return false;
		}
		
		var bResidencyDetailsExist = ((CustomerXML.GetTagText("COUNTRYOFRESIDENCY") != "") || 
											(CustomerXML.GetTagText("REASONFORCOUNTRYOFRESIDENCY") != "") || 
											(CustomerXML.GetTagText("RIGHTTOWORK") != "") || 
											(CustomerXML.GetTagText("DIPLOMATICIMMUNITY") != ""));
											
		var bResidencyDetailsIncomplete = ((CustomerXML.GetTagText("COUNTRYOFRESIDENCY") == "") || 
											(CustomerXML.GetTagText("REASONFORCOUNTRYOFRESIDENCY") == "") || 
											(CustomerXML.GetTagText("RIGHTTOWORK") == "") || 
											(CustomerXML.GetTagText("DIPLOMATICIMMUNITY") == ""));
											
		if((frmScreen.optUKResidentNo.checked) && bResidencyDetailsIncomplete)
			{
			alert("Please complete details of customers country of residence");
			frmScreen.btnResidency.focus();
			return false;
			}
		else if((frmScreen.optUKResidentYes.checked) && bResidencyDetailsExist)
			{
			var bUkResident = confirm("Please confirm customer is a UK resident by pressing \"OK\" otherwise press \"Cancel\"");
			if(bUkResident) <%/* remove non resident details */%>
				{
				CustomerXML.SetTagText("COUNTRYOFRESIDENCY", "");
				CustomerXML.SetTagText("REASONFORCOUNTRYOFRESIDENCY", "");
				CustomerXML.SetTagText("RIGHTTOWORK", "");
				CustomerXML.SetTagText("DIPLOMATICIMMUNITY", "");
				}
			else
				{
				alert("Please amend the residency status and check the residency details");
				frmScreen.optUKResidentNo.click();
				return false;
				}
			}
		<%/* MAH 15/11/2006 E2CR35 END*/%>
		<% /* AW MAR390 Other title mandatory */ %>
		if(scScreenFunctions.IsValidationType(frmScreen.cboTitle,"O") && frmScreen.txtOtherTitle.value == "") 
		{
			alert("Please enter the 'Other' Title or Change Title selection ");
			frmScreen.cboTitle.focus();
			return false;
		}
		<% /* SYS1053 Prevent the user from inputting number as first character of surname + forenames */ %>
		<% /* MF MAR019 WP01 Fields now readonly
		if (CheckNames(frmScreen.txtSurname.value) == false)
		{
			alert("First character of surname must be alpha");
			frmScreen.txtSurname.focus();
			return false;
		}
		if (CheckNames(frmScreen.txtFirstForename.value) == false)
		{
			alert("First character of first forename must be alpha");
			frmScreen.txtFirstForename.focus(); 
			return false;
		}
		if (CheckNames(frmScreen.txtSecondForename.value) == false)
		{
			alert("First character of second forename must be alpha");
			frmScreen.txtSecondForename.focus(); 
			return false;
		}
		if (CheckNames(frmScreen.txtOtherForenames.value) == false)
		{
			alert("First character of other forenames must be alpha");
			frmScreen.txtOtherForenames.focus(); 
			return false;
		}
		*/ %>
		if (IsChanged())
		{
			if(frmScreen.txtAge.value == "")
			{
				// the date of birth is invalid
				alert("Date of birth invalid");
				bSuccess = false;
			}
			
			<% /* PJO 18/05/2003 - BMIDS00619 - Unused			
			if(bSuccess &&
			   m_nOriginalRelationshipSel != -1 &&
			   m_nOriginalRelationshipSel != frmScreen.cboRelationship.selectedIndex)
			{
				//save changed relationship details
				//next line replaced by line below as per Core V7 - DPF 19/06/02 - BMIDS00077
				//var XML = new scXMLFunctions.XMLObject();
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				with (XML)
				{
					CreateRequestTag(window, null)
					CreateActiveTag("CUSTOMERRELATIONSHIP");
					CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
					CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
					CreateTag("OWNERCUSTOMERNUMBER", m_sCurrentCustomerNumber);				
					CreateTag("OWNERCUSTOMERVERSIONNUMBER", m_sCurrentCustomerVersionNumber);
					CreateTag("RELTOCUSTOMERNUMBER",m_sCustomer2Number);
					CreateTag("RELTOCUSTOMERVERSIONNUMBER",m_sCustomer2Version);
					CreateTag("RELATIONSHIPTOOWNER", frmScreen.cboRelationship.value);
					// 					RunASP(document, "SaveCustomerRelationship.asp");
					// Added by automated update TW 09 Oct 2002 SYS5115
					switch (ScreenRules())
						{
						case 1: // Warning
						case 0: // OK
											RunASP(document, "SaveCustomerRelationship.asp");
							break;
						default: // Error
							SetErrorResponse();
						}

				}

				if(!XML.IsResponseOK())
					bSuccess = false;

				XML = null;
			}			
			*/ %>
			if(bSuccess)
			{
				// Save CUSTOMER info
				//next line replaced by line below as per Core V7 - DPF 19/06/02 - BMIDS00077
				//var XML = new scXMLFunctions.XMLObject();
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				XML.CreateRequestTag(window, null)
				// add OPERATION IK_04/03/2001
				XML.SetAttribute("OPERATION","CriticalDataCheck");
				XML.CreateActiveTag("CUSTOMER");
				//get this info from the CustomerXML as it hasn't changed
				CustomerXML.SelectTag(null,"CUSTOMER");
				XML.CreateTag("CUSTOMERNUMBER", CustomerXML.GetTagText("CUSTOMERNUMBER"));
				XML.CreateTag("OTHERSYSTEMCUSTOMERNUMBER", CustomerXML.GetTagText("OTHERSYSTEMCUSTOMERNUMBER"));
				XML.CreateTag("OTHERSYSTEMTYPE",CustomerXML.GetTagText("OTHERSYSTEMTYPE"));
				XML.CreateTag("CHANNELID",CustomerXML.GetTagText("CHANNELID"));
				XML.CreateTag("INTERMEDIARYGUID",CustomerXML.GetTagText("INTERMEDIARYGUID"));

				// Save CUSTOMERVERSION info
				XML.CreateActiveTag("CUSTOMERVERSION");
				XML.CreateTag("CUSTOMERNUMBER",m_sCurrentCustomerNumber);
				XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCurrentCustomerVersionNumber);
				XML.CreateTag("SURNAME", frmScreen.txtSurname.value);
				XML.CreateTag("FIRSTFORENAME", frmScreen.txtFirstForename.value);
				XML.CreateTag("SECONDFORENAME", frmScreen.txtSecondForename.value);
				XML.CreateTag("OTHERFORENAMES", frmScreen.txtOtherForenames.value);
				XML.CreateTag("PREFERREDNAME", frmScreen.txtPreferredname.value);
				XML.CreateTag("TITLE", frmScreen.cboTitle.value);
				XML.CreateTag("TITLEOTHER", frmScreen.txtOtherTitle.value);
				XML.CreateTag("DATEOFBIRTH", frmScreen.txtDateOfBirth.value);
				XML.CreateTag("TIMEATBANKYEARS", frmScreen.txtBankYears.value);
				XML.CreateTag("TIMEATBANKMONTHS", frmScreen.txtBankMonths.value);
				XML.CreateTag("AGE", frmScreen.txtAge.value);
				<% /* PSC 11/01/2007 EP2_741 */ %>
				XML.CreateTag("MOTHERSMAIDENNAME", frmScreen.txtMotherMaidenName.value);
				XML.CreateTag("GENDER", frmScreen.cboGender.value);				
				XML.CreateTag("NATIONALITY", frmScreen.cboNationality.value);
				XML.CreateTag("MARITALSTATUS", frmScreen.cboMaritalStatus.value);
				<% /* MF MARS WP01 add special needs XML */ %>
				XML.CreateTag("SPECIALNEEDS", frmScreen.cboSpecialNeeds.value);			
				<% /* PJO 18/05/2003 - BMIDS00619 - Unused
				if(frmScreen.optSmokerYes.checked)
					XML.CreateTag("SMOKERSTATUS", "1");
				else if(frmScreen.optSmokerNo.checked)
					XML.CreateTag("SMOKERSTATUS", "0");
				else
					XML.CreateTag("SMOKERSTATUS", "");

				if(frmScreen.optGoodHealthYes.checked)
					XML.CreateTag("GOODHEALTH", "1");
				else if(frmScreen.optGoodHealthNo.checked)
					XML.CreateTag("GOODHEALTH", "0");
				else
					XML.CreateTag("GOODHEALTH", "")
				*/ %>				
				<% /* MV - 29/05/2002 - BMIDS00013 - BM044 - Amend NoOfDependants */ %>
				XML.CreateTag("NUMBEROFDEPENDANTS",frmScreen.txtNumberOfDependants.value);
				<% /* MF 09/08/2005 MARS20 WP06 New fields */ %>
				XML.CreateTag("TIMEATBANKYEARS",frmScreen.txtBankYears.value);
				XML.CreateTag("TIMEATBANKMONTHS",frmScreen.txtBankMonths.value);
				XML.CreateTag("YEARADDEDTOVOTERSROLL", frmScreen.txtElectoralRoll.value);
				if(frmScreen.optUKResidentYes.checked)
					XML.CreateTag("UKRESIDENTINDICATOR", "1");
				else if(frmScreen.optUKResidentNo.checked)
					XML.CreateTag("UKRESIDENTINDICATOR", "0");
				else
					XML.CreateTag("UKRESIDENTINDICATOR", "");
				<%/* MAH 15/11/2006 E2CR35 START */%>
				
				CustomerXML.SelectTag(null,"CUSTOMERVERSION");
				XML.CreateTag("COUNTRYOFRESIDENCY", CustomerXML.GetTagText("COUNTRYOFRESIDENCY")); 
				XML.CreateTag("REASONFORCOUNTRYOFRESIDENCY", CustomerXML.GetTagText("REASONFORCOUNTRYOFRESIDENCY"));
				XML.CreateTag("RIGHTTOWORK", CustomerXML.GetTagText("RIGHTTOWORK")); 
				XML.CreateTag("DIPLOMATICIMMUNITY", CustomerXML.GetTagText("DIPLOMATICIMMUNITY"));
				
				<%/* MAH 14/11/2006 E2CR35 END */%>
				if(frmScreen.optStaffYes.checked)
					XML.CreateTag("MEMBEROFSTAFF", "1");
				else if(frmScreen.optStaffNo.checked)
					XML.CreateTag("MEMBEROFSTAFF", "0");
				else
					XML.CreateTag("MEMBEROFSTAFF", "");
				<%/* MAH 14/11/2006 E2CR35 END */%>
				if(frmScreen.optSufficientRepaymentIndicatorYes.checked)
					XML.CreateTag("SUFFICIENTREPAYMENTINDICATOR", "1");
				else if(frmScreen.optSufficientRepaymentIndicatorNo.checked)
					XML.CreateTag("SUFFICIENTREPAYMENTINDICATOR", "0");
				else
					XML.CreateTag("SUFFICIENTREPAYMENTINDICATOR", "");

				if(frmScreen.optMailshotRequiredYes.checked)
					XML.CreateTag("MAILSHOTREQUIRED", "1");
				else if(frmScreen.optMailshotRequiredNo.checked)
					XML.CreateTag("MAILSHOTREQUIRED", "0");
				else
					XML.CreateTag("MAILSHOTREQUIRED", "");

				XML.CreateTag("CONTACTEMAILADDRESS", CustomerXML.GetTagText("CONTACTEMAILADDRESS"));
				XML.CreateTag("EMAILPREFERRED", CustomerXML.GetTagText("EMAILPREFERRED"));
				XML.CreateTag("CORRESPONDENCESALUTATION", CustomerXML.GetTagText("CORRESPONDENCESALUTATION"));
				
				<% /* EP903 */ %>
				XML.CreateTag("APPLICATIONDATE", m_sApplicationDate);
				<% /* EP2_677 */ %>
				XML.CreateTag("HOUSINGBENEFIT", CustomerXML.GetTagText("HOUSINGBENEFIT"));
								
				// Save CUSTOMERTELEPHONE info
				//Lift it straight out of CustomerXML
				var tagTelephone = CustomerXML.SelectTag(null,"CUSTOMERTELEPHONENUMBERLIST");
				XML.SelectTag(null,"CUSTOMER");
				XML.CreateActiveTag("CUSTOMERTELEPHONENUMBERLIST");
				XML.AddXMLBlock(tagTelephone);

				// add CRITICALDATACONTEXT IK_04/03/2001
				XML.SelectTag(null,"REQUEST");
				XML.CreateActiveTag("CRITICALDATACONTEXT");
				XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
				XML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
				XML.SetAttribute("SOURCEAPPLICATION","Omiga");
				XML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
				XML.SetAttribute("ACTIVITYINSTANCE","1");
				XML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
				XML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
				XML.SetAttribute("COMPONENT","omCust.CustomerBO");
				XML.SetAttribute("METHOD","UpdateCustomerPersonalDetails");
				
				window.status = "Critical Data Check - please wait";
				
				// 				XML.RunASP(document, "OmigaTMBO.asp");
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
					{
					case 1: // Warning
					case 0: // OK
									XML.RunASP(document, "OmigaTMBO.asp");
						break;
					default: // Error
						XML.SetErrorResponse();
					}

				
				window.status = "";
				
				/* IK_04/03/2001
				XML.RunASP(document, "SavePersonalDetails.asp");
				*/
				
				bSuccess = XML.IsResponseOK();
				XML = null;
			}
		}
	}

	return(bSuccess);
}

<% /* MF 09/09/2005 Created rather than use scSCreenFunctions.html method as this
		did not give the desired appearance. This call gives a different look 
		to the background:
			scScreenFunctions.SetFieldToReadOnly(frmScreen,sField);  
*/ %>
function MakeFieldReadOnly(sField, bReadOnly){
	document.all(sField).disabled=bReadOnly;			
	switch (document.all(sField).tagName){
		case "select":
			if(bReadOnly){
				document.all(sField).className="msgReadOnlyCombo";
			}else{
				document.all(sField).className="msgCombo";
			}
			break;				
		default:
			<% /* PSC 30/01/2006 MAR939 - Start */ %>	
			if (document.all(sField).type == "radio")
			{
				if (bReadOnly)
				{
					document.all(sField).disabled = true;
				}
			}
			else
			{			
				if(bReadOnly){
					document.all(sField).className="msgReadOnly";
				}else{
					document.all(sField).className="msgTxt";
				}
			}
			<% /* PSC 30/01/2006 MAR939 - End */ %>	
			break;
	}						
}

function MakeFieldsReadOnly(bReadOnly){
	var aFields = new Array ("txtCustName",
		<% /* EP529 - MAR1274 */ %>
		//"txtSurname", "txtFirstForename", "txtSecondForename",
		"txtSurname", "txtFirstForename",
		<% /* EP529 - MAR1274 End */ %>
		"txtOtherForenames", "txtPreferredname", 
		"cboTitle", "txtOtherTitle",
		"txtAge", "cboGender", "cboSpecialNeeds", 
		"txtElectoralRoll", "MailshotGroup",
		"CreditSearchGroup");
		
	for(var i=0;i<aFields.length;i++){						
		MakeFieldReadOnly(aFields[i],bReadOnly);						
	}
}

function PopulateScreen()
{
	//set the Customer Name
	var sCustomerName = scScreenFunctions.GetContextCustomerName(window, m_sCurrentCustomerNumber);
	frmScreen.txtCustName.value = sCustomerName;
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCustName");
	<% /* MF MARS WP01 add Disable all fields apart from DOB */ %>
	<% /* EP529 / MAR1274 ...apart from DOB and Middle name if a middle name has not been input  */ %>
	
	<% /* PE - 25/01/2006 - MAR939 - Parameterised the disabling of the customer screens	*/ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bDisable = XML.GetGlobalParameterBoolean(document,"DisableCustomerDetails");	
	MakeFieldsReadOnly(bDisable);	
	
	<% /* MAR1602 GHun */ %>
	CustomerXML.SelectTag(null,"CUSTOMER");
	frmScreen.txtOtherSysCustomerNo.value = CustomerXML.GetTagText("OTHERSYSTEMCUSTOMERNUMBER");
	if (frmScreen.txtOtherSysCustomerNo.value == "")
	{
		scScreenFunctions.HideCollection(spnOtherSysCustomerNo);
	}
	MakeFieldReadOnly("txtOtherSysCustomerNo", true);
	<% /* MAR1602 End */ %>
	
	CustomerXML.SelectTag(null,"CUSTOMERVERSION");

	with (frmScreen)
	{
		<% /* MF MARS WP01 add special needs combo */ %>
		// EP2_610 - Default to Select.
		if( CustomerXML.GetTagText("SPECIALNEEDS")!= null && CustomerXML.GetTagText("SPECIALNEEDS")!= "") 		
			cboSpecialNeeds.value = CustomerXML.GetTagText("SPECIALNEEDS"); 
		else
			cboSpecialNeeds.selectedIndex = 0;	
		txtSurname.value = CustomerXML.GetTagText("SURNAME");
		txtFirstForename.value = CustomerXML.GetTagText("FIRSTFORENAME");
		txtSecondForename.value = CustomerXML.GetTagText("SECONDFORENAME");
		<% /* EP529 / MAR1274 */ %>
		MakeFieldReadOnly("txtSecondForename",CustomerXML.GetTagText("SECONDFORENAME")!="" && bDisable);
		<% /* EP529 / MAR1274 End */ %>
		txtOtherForenames.value = CustomerXML.GetTagText("OTHERFORENAMES");
		txtPreferredname.value = CustomerXML.GetTagText("PREFERREDNAME");
		cboTitle.value = CustomerXML.GetTagText("TITLE");
		txtOtherTitle.value = CustomerXML.GetTagText("TITLEOTHER");
		txtBankYears.value = CustomerXML.GetTagText("TIMEATBANKYEARS");
		txtBankMonths.value = CustomerXML.GetTagText("TIMEATBANKMONTHS");
		
		txtDateOfBirth.value = CustomerXML.GetTagText("DATEOFBIRTH");
		<% /* PSC 30/01/2006 MAR939 */ %>	
		MakeFieldReadOnly("txtDateOfBirth",CustomerXML.GetTagText("DATEOFBIRTH")!="" && bDisable);
//		EP2_610 - Now Mandatory so no messing about with colour schemes. You can have any colour as long as it's RED.
//		if (txtDateOfBirth.value == ""){					
//			frmScreen.txtDateOfBirth.parentElement.parentElement.style.color = "red";
//		}else{
//			<% /* MF Modification to MAR19. Use next element to return default color */ %>
//			frmScreen.txtDateOfBirth.parentElement.parentElement.style.color
//				= frmScreen.txtAge.parentElement.parentElement.style.color;			
//		}
		//set age up
		var nAge = CalculateCustomerAge();
		if(nAge != -1)
			txtAge.value = nAge.toString();
		cboGender.value = CustomerXML.GetTagText("GENDER");	
		
		<% /* PSC 11/01/2007 EP2_741 - Start */ %>
		txtMotherMaidenName.value = CustomerXML.GetTagText("MOTHERSMAIDENNAME");
		MakeFieldReadOnly("txtMotherMaidenName",CustomerXML.GetTagText("MOTHERSMAIDENNAME")!="" && bDisable);
		<% /* PSC 11/01/2007 EP2_741 - End */ %>	
		
		var bButtonEnabled = false;
		btnAlias.disabled = bButtonEnabled;

		<%/* MAH 14/11/2006 E2CR35 */%>
		cboNationality.value = CustomerXML.GetTagText("NATIONALITY");		
		<% /* PSC 30/01/2006 MAR939 */ %>	
		MakeFieldReadOnly("cboNationality",CustomerXML.GetTagText("NATIONALITY")!="" && bDisable);
		cboMaritalStatus.value = CustomerXML.GetTagText("MARITALSTATUS");				
		<% /* PSC 30/01/2006 MAR939 */ %>	
		MakeFieldReadOnly("cboMaritalStatus",CustomerXML.GetTagText("MARITALSTATUS")!="" && bDisable);
		scScreenFunctions.SetRadioGroupValue(frmScreen, "UKResidentGroup", CustomerXML.GetTagText("UKRESIDENTINDICATOR"));		
		<% /* PSC 30/01/2006 MAR939 */ %>	
		MakeFieldReadOnly("optUKResidentNo",CustomerXML.GetTagText("UKRESIDENTINDICATOR")!="" && bDisable);
		<% /* PSC 30/01/2006 MAR939 */ %>	
		MakeFieldReadOnly("optUKResidentYes",CustomerXML.GetTagText("UKRESIDENTINDICATOR")!="" && bDisable);
		<%/* MAH 14/11/2006 E2CR35 */%>
		if(CustomerXML.GetTagText("UKRESIDENTINDICATOR")=="1")
			frmScreen.btnResidency.disabled = true;
		else
			frmScreen.btnResidency.disabled = false;
		<%/* MAH 14/11/2006 E2CR35 */%>
		scScreenFunctions.SetRadioGroupValue(frmScreen, "StaffGroup", CustomerXML.GetTagText("MEMBEROFSTAFF"));
		<% /* PSC 30/01/2006 MAR939 */ %>	
		MakeFieldReadOnly("optStaffNo",CustomerXML.GetTagText("MEMBEROFSTAFF")!="" && bDisable);
		<% /* PSC 30/01/2006 MAR939 */ %>	
		MakeFieldReadOnly("optStaffYes",CustomerXML.GetTagText("MEMBEROFSTAFF")!="" && bDisable);
		<% /*  MV - 24/05/2002 - BMIDS00013 - BM044 - Amend NOOFDEPENDANTS */%>
		txtNumberOfDependants.value = CustomerXML.GetTagText("NUMBEROFDEPENDANTS");	
		<% /* PSC 30/01/2006 MAR939 */ %>	
		MakeFieldReadOnly("txtNumberOfDependants",CustomerXML.GetTagText("NUMBEROFDEPENDANTS")!="" && bDisable);
		
		txtElectoralRoll.value = CustomerXML.GetTagText("YEARADDEDTOVOTERSROLL");		
		
		<% /* MF 02/08/2005 MARS20 WP06 New fields */ %>
		txtBankYears.value = CustomerXML.GetTagText("TIMEATBANKYEARS");
		txtBankMonths.value = CustomerXML.GetTagText("TIMEATBANKMONTHS");		
		<% /* PJO 18/05/2003 - BMIDS00619 - Unused
		scScreenFunctions.SetRadioGroupValue(frmScreen, "SmokerGroup", CustomerXML.GetTagText("SMOKERSTATUS"));
		scScreenFunctions.SetRadioGroupValue(frmScreen, "GoodHealthGroup", CustomerXML.GetTagText("GOODHEALTH"));
		*/ %>		
		scScreenFunctions.SetRadioGroupValue(frmScreen, "MailshotGroup", CustomerXML.GetTagText("MAILSHOTREQUIRED"));
		scScreenFunctions.SetRadioGroupValue(frmScreen, "SufficientRepaymentIndicatorGroup", CustomerXML.GetTagText("SUFFICIENTREPAYMENTINDICATOR"));
		<% /* BMIDS00336 MDC 27/08/2002 */ %>
		var sCreditSearch = CustomerXML.GetTagText("CREDITSEARCH");
		if(sCreditSearch == "")
			scScreenFunctions.SetRadioGroupValue(frmScreen, "CreditSearchGroup", "0");
		else	
			scScreenFunctions.SetRadioGroupValue(frmScreen, "CreditSearchGroup", CustomerXML.GetTagText("CREDITSEARCH"));
		<% /* BMIDS00336 MDC 27/08/2002 - End */ %>

		<% /* PJO 18/05/2003 - BMIDS00619 - Unused
		if (IsSingleApplicant()) scScreenFunctions.HideCollection(spnRelationship)
		*/ %>	

		<% /* EP2_127 Disable if dependants collected at Application level */ %>
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();		
		frmScreen.txtNumberOfDependants.disabled = (XML.GetGlobalParameterBoolean(document, "DCDependantsAtApplicationLevel") == true) ? true : false;
		XML = null; //Kill			
	}
}

function IsSingleApplicant()
// Returns true if there is only a single applicant
{
	var nApplicantCount = 0;
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);
		if(sCustomerRoleType == "1") nApplicantCount++;
	}
	return (nApplicantCount < 2);
}


function Initialise()
{
	PopulateCombos();

	if(m_sCurrentCustomerNumber == "" &&
	   m_sCurrentCustomerVersionNumber == "")
	{
		// set up to point to the first customer from context
		m_sCurrentCustomerNumber = m_sCustomer1Number;
		m_sCurrentCustomerVersionNumber = m_sCustomer1Version;
	}

	GetCustomer();
	<% /* PJO 18/05/2003 - BMIDS00619 - Unused
	SetUpRelationship()
	*/ %>		

	if(m_sReadOnly == "1")
		scScreenFunctions.SetScreenToReadOnly(frmScreen);

	<% /* Set label colour for mandatory option groups */ %>
	<% /* PJO 18/05/2003 - BMIDS00619 - Unused
	frmScreen.optGoodHealthNo.parentElement.parentElement.style.color = "red";
	frmScreen.optSmokerNo.parentElement.parentElement.style.color = "red";
	*/ %>
	<% /* MF MAR019 WP01 
	frmScreen.optStaffNo.parentElement.parentElement.style.color = "red";
	frmScreen.optUKResidentNo.parentElement.parentElement.style.color = "red";
	*/ %>	
	PopulateScreen();
	//save the date of birth for comparison later
	m_sDateOfBirth = frmScreen.txtDateOfBirth.value;

	HideOtherTitle();
	
	<% /*MV - 29/05/2002 - BMIDS00013 -  Default Good Health and smoker status */ %>
	<% /* PJO 18/05/2003 - BMIDS00619 - Unused
	SetDefaultToRadioButtons();
	*/ %>		
	<% /*	BMIDS00336 MDC 27/08/2002 */ %>
	<% /* MF MARS019 WP01 statically set in HTML
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "optCreditSearchYes");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "optCreditSearchNo");
	*/ %>
	<% /*	BMIDS00336 MDC 27/08/2002 - End */ %>	
	<% /* MF MARS WP01 Fields are now always read only 
	if(m_sReadOnly != "1")	
		frmScreen.txtSurname.focus();
	*/ %>
}

function GetCustomerNumber(nApplicantNumber)
{
	var bSuccess = true;

	switch (nApplicantNumber)
	{
		case 1:
			m_sNextCustNum		= m_sCustomer1Number;
			m_sNextCustVersion	= m_sCustomer1Version;
			break;
		case 2:
			m_sNextCustNum		= m_sCustomer2Number;
			m_sNextCustVersion	= m_sCustomer2Version;
			break;
		case 3:
			m_sNextCustNum		= m_sCustomer3Number;
			m_sNextCustVersion	= m_sCustomer3Version;
			break;
		case 4:
			m_sNextCustNum		= m_sCustomer4Number;
			m_sNextCustVersion	= m_sCustomer4Version;
			break;
		case 5:
			m_sNextCustNum		= m_sCustomer5Number;
			m_sNextCustVersion	= m_sCustomer5Version;
			break;
		default:
			m_sNextCustNum		= "";
			m_sNextCustVersion	= "";
			break;
	}

	if(m_sNextCustNum == "") bSuccess = false;
	return(bSuccess);
}

function GetNumberOfCustomer(sCurrentCustomerNumber)
{
	var nNumber;
	if(sCurrentCustomerNumber == m_sCustomer1Number) nNumber = 1;
	else if(sCurrentCustomerNumber == m_sCustomer2Number) nNumber = 2;
	else if(sCurrentCustomerNumber == m_sCustomer3Number) nNumber = 3;
	else if(sCurrentCustomerNumber == m_sCustomer4Number) nNumber = 4;
	else nNumber = 5;

	return(nNumber);
}

function GetNextCustomerNumber(sCurrentCustomerNumber, nDirection)
{
	var bNextCustomerIsValid = false;
			
	var nThisCustApplicantNumber = GetNumberOfCustomer(sCurrentCustomerNumber);
	nThisCustApplicantNumber += nDirection;
	bNextCustomerIsValid = GetCustomerNumber(nThisCustApplicantNumber);
			
	return(bNextCustomerIsValid);
}

<% /* SYS1053 SA 24/5/01 New function to check 1st character of names */ %>
function CheckNames(sNameToCheck)
{
	var CheckStr = sNameToCheck.substr(0,1);
	if (_validation.isNum(CheckStr))
	{
		return false;
	}
	
	return true;	
	
}

function btnCancel.onclick()
{
	CustomerXML = null;
	
	if(m_sContext == "CompletenessCheck")
	{
		frmToGN300.submit();
		return;
	}

	if(GetNextCustomerNumber(m_sCurrentCustomerNumber, -1))
	{
		m_sCurrentCustomerNumber = m_sNextCustNum;
		m_sCurrentCustomerVersionNumber = m_sNextCustVersion;
		Initialise();
		<% /* PB 25/05/2006 EP612 Recheck for readonly fields */ %>
		m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC030");
		<% /* EP612 End */ %>
	}
	else
<% /*MF 22/07/2005 IA_WP01 process flow change
		frmToDC020.submit();
*/ %>
		frmToDC010.submit();
}

function btnSubmit.onclick()
{
	var bSuccess = CommitData();
			
	if(bSuccess == true)
	{
		CustomerXML = null;

		if(m_sContext == "CompletenessCheck")
		{
			frmToGN300.submit();
			return;
		}
					
		if(GetNextCustomerNumber(m_sCurrentCustomerNumber, 1))
		{
			m_sCurrentCustomerNumber = m_sNextCustNum;
			m_sCurrentCustomerVersionNumber = m_sNextCustVersion;
			Initialise();
			<% /* PB 23/05/2006 EP586/612 */ %>
			m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC030");
			<% /* EP586/612 End */ %>
			<% /* BS 28/01/2003 BM0271 */ %>
			ClientPopulateScreen();
		}
		else
			frmToDC060.submit();	

	}
}		

<% /* PJO 18/05/2003 - BMIDS00619 - Unused
function SetDefaultToRadioButtons()
{
	frmScreen.optGoodHealthYes.checked = true ;
	frmScreen.optSmokerNo.checked = true;  
}
*/ %>
<%/* MAH 14/11/2006 E2CR35 START*/%>

function frmScreen.optUKResidentYes.onclick()
{
	frmScreen.btnResidency.disabled = true;
}
function frmScreen.optUKResidentNo.onclick()
{
	frmScreen.btnResidency.disabled = false;
}
function frmScreen.btnResidency.onclick()
{
	var sReturn = null;
	var ArrayArguments = CustomerXML.CreateRequestAttributeArray(window, null)

<% // GD BMIDS00376 START %>
	ArrayArguments[4] = (m_blnReadOnly || m_sReadOnly);
<% // GD BMIDS00376 END %>

	ArrayArguments[5] = CustomerXML.XMLDocument.xml;
	ArrayArguments[6] = scScreenFunctions.GetContextCustomerName(window, m_sCurrentCustomerNumber);

	sReturn = scScreenFunctions.DisplayPopup(window, document, "DC036.asp", ArrayArguments, 630, 395);

	if(sReturn != null)
	{
		FlagChange(sReturn[0]);
		var sXML = sReturn[1];
		CustomerXML.LoadXML(sXML);
	}
}
<% // EP2_677 %>
function frmScreen.btnPersonalDetails.onclick()
{
	var sReturn = null;
	var ArrayArguments = CustomerXML.CreateRequestAttributeArray(window, null)

	ArrayArguments[4] = (m_blnReadOnly || m_sReadOnly);
	ArrayArguments[5] = CustomerXML.XMLDocument.xml;
	ArrayArguments[6] = scScreenFunctions.GetContextCustomerName(window, m_sCurrentCustomerNumber);

	// EP2_1007 - Resize screen - Now less silly!
	sReturn = scScreenFunctions.DisplayPopup(window, document, "DC037.asp", ArrayArguments, 530, 175);

	if(sReturn != null)
	{
		FlagChange(sReturn[0]);
		var sXML = sReturn[1];
		CustomerXML.LoadXML(sXML);
	}
}
-->
</script>
</body>
</html>
