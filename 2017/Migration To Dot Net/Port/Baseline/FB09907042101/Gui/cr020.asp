<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<%
/*
Workfile:      cr020.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Customer Add/Edit screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		04/02/00	Optimised version
AY		11/02/00	Change to msgButtons button types
AY		29/02/00	SYS0318 - Clear the ReadOnly and CustomerReadOnly context fields
AY		01/03/00	SYS0048 - Check for Other title now looks at the validation type, not combo.value
AY		08/03/00	Changes to AddCustomerToPackage processing
AY		14/03/00	SYS0050 - Display customer name in framework if necessary
					SYS0189 - Age < 18 now validated on lose focus
					SYS0220 - OtherSystemNumber processing removed
SR		15/03/00    SYS0393 - Call AddCustomerToApplication instead of AddCustomerToPackage 
AY		24/03/00	SYS0322 - Flag introduced to prevent submit processing being run twice
AY		29/03/00	New top menu/scScreenFunctions change
MH      02/05/00    SYS00618 Postcode validation
IW		08/05/00	SYS0584	 Added prefered contact option buttons
IW		09/05/00	SYS0085  Changes to Telephone Validation.
IW		10/05/00	SYS0730  Changed order of validation routines.
IW		15/05/00	SYS0201  Customer Name Populated for CR060
IW		15/05/00	SYS0482  Disable Areas of Interest (until such time as there are any).
IW`		22/05/00	SYS0745	 MISC Cosmetic changes.
IW		22/05/00	SYS0712	 Preffered Contact not updating correctly.
IW		25/05/00	SYS0776	 Added "EmailPreferred Indicator" option.
MC		14/06/00	SYS0935  Amended to use PREFERREDMETHODOFCONTACT
BG		24/08/00	SYS0745	 Check full date has been entered and that it is a valid date
CL		25/01/01	SYS1756	 Adjusted to call Contact history CR025	
GD		20/02/01	SYS1752  Child AQR to SYS1748
							 Added OTHERSYSTEMCUSTOMERNUMBER
CL		05/03/01	SYS1920 Read only functionality added
IK		15/03/01	SYS1924 Completeness Check routing
SA		24/05/01	SYS1053 Check Surname + forename fields for first char being numeric
					New function CheckNames added.
SR		13/06/01	SYS2362 - add a new text field to display 'OtherSystemCustomerNumber' 
JR		14/06/01	Omiplus24 - Amend TelephoneDetails to cater for Country and Area Code
					and added VBScript to line that calls CalculateCustomerAge sub.
JLD		4/12/01		SYS2806 use scScreenfunctions.CompletenessCheckRouting
DRC     12/04/02    SYS3964 Changed Age Calc to JScript for ambiguous dates
STB		30/04/02	SYS0003 Store Customer/VersionNo as module-level variables when sync'ing downloaded customers.
MEVA	22/05/02	SYS2583 Remove Areas of Interest
LD		23/05/02	SYS4727 Use cached versions of frame functions
SG		29/05/02	SYS4767 MSMS to Core integration
SG		05/06/02	SYS4804 Error from SYS4767
JLD		12/06/02	sys4728 use stylesheet at all times
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMids History:

Prog	Date		Description
MDC		21/05/2002	BMIDS00004 - BM073 Versions of employment
MV		22/05/2002  BMIDS00005 - BM052 Customer Searches Modified btnCancel.onclick()
GD		26/06/02	BMIDS0077	applied SG 25/06/02 SYS4930 
MV		02/07/2002	BMIDS00120 - Amended btnSubmit.onclick()  with new scXMLFunctions Declaration
MV		04/07/2002  BMIDS00100 - Reseting the values 
MV		05/07/2002	BMIDS00118 - Added  frmScreen.txtOtherTitle.onblur() and Amended DefaultGenderFromTitle()
MV		05/07/2002	BMIDS00096 - Modified CopyAddress() and PopulateScreen();
MV		09/07/2002  BMIDS00100 - Reseting the values - Amended btnCancel.onclick() 
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
ASu		02/09/2002	BMIDS00139 - Remove 'Areas Of Interest' from GUI	
MDC		03/09/2002	BMIDS00393 - Attempt to get customer details for metaAction
								 CreateNewCustomerForNewApplication (required when cancelling from CR025)
ASu		26/09/2002	BMIDS00453 - Remove mandatory validation from 'Other' textbox,and test on screen submit.
TW      09/10/2002	Modified to incorporate client validation - SYS5115
ASu		23/10/2002	BMIDS00690 - Add additional checks to enable/disable 'Copy' button and make changes to copy function.
PSC		23/10/2002	BMIDS00465 - Route to CR030 if updating existing customer
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client
SR		29/11/2002	BMIDS01048 - modifed the logic for clearing CustomerName1 from context
GHun	03/11/2002	BM0077	   - There should be no validation when clicking OK when the screen is readonly
MV		06/12/2002	BMIDS0020	Modified the PAF Search tab Order
PJO     03/07/2003  BM0006      Allow guarantor to be added after 4 applicants
KRW     13/05/2004  BMIDS745    cboCurrentAddressCountry no longer defaults to England 
HMA     01/07/2004  BMIDS758    Set indicator if new customer is being added to a Transfer of Equity
HMA     29/07/2004  BMIDS758    Make sure the indicator is set for existing customers.
HMA     04/08/2004  BMIDS836    Use "CC" as validation type for CC069
HMA     05/08/2004  BMIDS836    Make sure NewToEInd is set correctly
HMA     12/08/2004  BMIDS836    Check that NewToEInd is set on new customer versions.
HMA     08/12/2004  BMIDS957    Unused VBScript code removed.
TLiu	02/09/2005	MAR38		Changed layout for Flat No., House Name & No.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS History
Prog	Date		AQR		Description
MF		20/07/2005	MAR019	Cosmetic Changes for IA_WP01
MF		07/09/2005	MAR20	Enabled title change event
MF		22/09/2005	MAR19	Modifications to fix multiple client bug, setting readonly state
							state of controls. Also Set title, sex, Special Needs mandatory.
							Renamed Sex to Gender	
PSC		11/10/2005	MAR57	Set up other system customer number on customer creation						
HM		08/11/2005	MAR160	Copy button behavior
SD		14/11/2005	MAR258	Critical Data check changes
PSC		30/11/2005	MAR729	Disable copy button
GHun	01/12/2005	MAR777  Temporarily enable country combo
PE		13/12/2005	MAR831	Reduced surname max length to 32, to fit into Oracle database.
PE		25/01/2006	MAR939	Parameterised the disabling of the customer screens
PSC		30/01/2006	MAR939	Correct logic
PAK		16/03/2006	MAR1443	Reverse MAR777 - disable country combo


~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM History
Prog	Date		AQR		Description
pct		13/03/2006	EP8		Behavioural adjustments for BFPO Addresses
pct		17/03/2006	EP225	Removal of read-only attributes on 2nd Forenames and DOB
LH		25/05/2006	EP594	Current address must be a UK address
HMA     12/07/20076	EP903   Pass Application Date through for customer age verification.
                            Use application date for age verification.	
GHun	19/12/2006	EP2_56	Changed CommitToPackage for TOE
DS		18/01/2007  EP2_610 Removed SpecialNeeds combo as mandatory field as per new requirement
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
	<head>
		<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4" />
		<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
		<title></title>
	</head>
	<body>
		<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
		<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden"
			tabIndex="-1" type="text/x-scriptlet" width="1" viewastext>
		</object>
		<% /* Validation script - Controls Soft Coded Field Attributes */ %>
		<script src="validation.js" language="JScript"></script>
		<% /* Scriptlets */ %>
		<% /* CORE UPGRADE 702 <object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object> */ %>
		<% /* Specify Forms Here */ %>
		<form id="frmToCR030" method="post" action="cr030.asp" STYLE="DISPLAY: none"></form>
		<form id="frmToCR025" method="post" action="cr025.asp" STYLE="DISPLAY: none"></form>
		<form id="frmToCR010" method="post" action="cr010.asp" STYLE="DISPLAY: none"></form>
		<form id="frmToGN300" method="post" action="gn300.asp" STYLE="DISPLAY: none"></form>
		<% /* Span to keep tabbing within this screen */ %>
		<span id="spnToLastField" tabindex="0"></span>
		<% /* Specify Screen Layout Here */ %>
		<form id="frmScreen" mark validate="onchange" year4>
			<% /* Personal Details */ %>
			<div style="HEIGHT: 144px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
				<span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel">
					<strong>Personal Details</strong>
				</span>
				<span style="LEFT: 4px; POSITION: absolute; TOP: 24px" class="msgLabel">
		Surname
		<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
						<input id="txtSurname" maxlength="32" style="POSITION: absolute; WIDTH: 200px" class="msgTxt" type="text">
					</span>
	</span>
				<span id="spnOtherSysCustomerNo" style="LEFT: 360px; POSITION: absolute; TOP: 24px" class="msgLabel">
		Customer Number
		<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
						<input id="txtOtherSysCustomerNo" maxlength="20" style="POSITION: absolute; WIDTH: 100px"
							class="msgTxt" type="text">
					</span>
	</span>
				<span style="LEFT: 4px; POSITION: absolute; TOP: 48px" class="msgLabel">
		Forenames
		<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
						<input id="txtFirstForename" maxlength="30" style="POSITION: absolute; WIDTH: 150px" class="msgTxt" type="text">
					</span>

		<span style="LEFT: 248px; POSITION: absolute; TOP: -3px">
						<input id="txtSecondForename" maxlength="30" style="POSITION: absolute; WIDTH: 150px" class="msgTxt" type="text">
					</span>

		<span style="LEFT: 416px; POSITION: absolute; TOP: -3px">
						<input id="txtOtherForenames" maxlength="30" style="POSITION: absolute; WIDTH: 150px" class="msgTxt" type="text">
					</span>
	</span>
				<span style="LEFT: 4px; POSITION: absolute; TOP: 72px" class="msgLabel">
		Title
		<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
						<select id="cboTitle" style="WIDTH: 150px" class="msgCombo">
						</select>
					</span>

		<span id="spnOtherTitle" style="LEFT: 248px; POSITION: absolute; TOP: -3px">
						<input id="txtOtherTitle" maxlength="20" style="POSITION: absolute; WIDTH: 150px" class="msgTxt" type="text">
					</span>
	</span>
				<span style="LEFT: 4px; POSITION: absolute; TOP: 96px" class="msgLabel">
		Date of Birth
		<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
						<% /* CORE UPGRADE 702 */ %>
						<input id="txtDateOfBirth" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgTxt" type="text">
					</span>

		<span style="LEFT: 204px; POSITION: absolute; TOP: 0px" class="msgLabel">
			Age
			<span style="LEFT: 44px; POSITION: absolute; TOP: -3px">
							<input id="txtAge" maxlength="3" style="POSITION: absolute; WIDTH: 40px" class="msgReadOnly"
								readOnly tabindex="-1" type="text">
						</span>
		</span>

		<span style="LEFT: 376px; POSITION: absolute; TOP: 0px" class="msgLabel">
			Gender
			<span style="LEFT: 40px; POSITION: absolute; TOP: -3px">
							<select id="cboGender" style="WIDTH: 100px" class="msgCombo">
							</select>
						</span>
		</span>
	</span								>
		<span class="msgLabel" style="LEFT: 4px; POSITION: absolute; TOP: 120px">
			Special Needs
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<select class="msgCombo" id="cboSpecialNeeds" style="WIDTH: 120px" name="cboSpecialNeeds">
				</select>
			</span>
		</span>
	</div>
			<% /* Address Details */ %>
			<div style="HEIGHT: 234px; LEFT: 10px; POSITION: absolute; TOP: 210px; WIDTH: 604px"
				class="msgGroup">
				<% /* Current Address */ %>
				<span id="spnCurrentAddress" style="LEFT: 4px; POSITION: absolute; TOP: 4px">
					<span style="LEFT: 0px; POSITION: absolute; TOP: 0px" class="msgLabel">
						<strong>Current Address</strong>
					</span> <!-- SG 29/05/02 SYS4767 START --> <!-- SG 28/02/02 SYS4186 -->
					<span style="LEFT: 120px; POSITION: absolute; TOP: 6px">
						<input id="btnCopy" value="Copy" type="button" style="WIDTH: 75px" class="msgButton">
					</span> <!-- SG 29/05/02 SYS4767 END -->
					<span style="LEFT: 204px; POSITION: absolute; TOP: 6px">
						<input id="btnCurrentAddressClearAddress" value="Clear" type="button" style="WIDTH: 75px"
							class="msgButton">
					</span>
					<span style="LEFT: 0px; POSITION: absolute; TOP: 40px" class="msgLabel">
			Postcode
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtCurrentAddressPostcode" maxlength="8" style="POSITION: absolute; WIDTH: 70px"
								class="msgTxtUpper" onchange="ClearCurrentAddressPAFIndicator()">
						</span>
		</span>
					<span style="LEFT: 0px; POSITION: absolute; TOP: 64px" class="msgLabel">
			Flat No./ Name
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtCurrentAddressFlatNumber" maxlength="10" style="POSITION: absolute; WIDTH: 70px"
								class="msgTxt" onchange="ClearCurrentAddressPAFIndicator()">
						</span>
		</span>

		<span style="LEFT: 0px; POSITION: absolute; TOP: 82px" class="msgLabel">
			House<br>No. &amp; Name
			<span style="LEFT: 80px; POSITION: absolute; TOP: 3px">
				<input id="txtCurrentAddressHouseNumber" maxlength="10" style="POSITION: absolute; WIDTH: 45px"	class="msgTxt" onchange="ClearCurrentAddressPAFIndicator()">
			</span>
			<span style="LEFT: 130px; POSITION: absolute; TOP: 3px">
				<input id="txtCurrentAddressHouseName" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgTxt" onchange="ClearCurrentAddressPAFIndicator()">
			</span>
		</span>

		<span style="LEFT: 179px; POSITION: absolute; TOP: 35px">
						<input id="btnCurrentAddressPAFSearch" value="P.A.F. Search" type="button" style="WIDTH: 100px"
							class="msgButton">
					</span>
					<span style="LEFT: 0px; POSITION: absolute; TOP: 112px" class="msgLabel">
			Street
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtCurrentAddressStreet" maxlength="40" style="POSITION: absolute; WIDTH: 200px"
								class="msgTxt" onchange="ClearCurrentAddressPAFIndicator()">
						</span>
		</span>
					<span style="LEFT: 0px; POSITION: absolute; TOP: 136px" class="msgLabel">
			District
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtCurrentAddressDistrict" maxlength="40" style="POSITION: absolute; WIDTH: 200px"
								class="msgTxt" onchange="ClearCurrentAddressPAFIndicator()">
						</span>
		</span>
					<span style="LEFT: 0px; POSITION: absolute; TOP: 160px" class="msgLabel">
			Town
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtCurrentAddressTown" maxlength="40" style="POSITION: absolute; WIDTH: 200px"
								class="msgTxt" onchange="ClearCurrentAddressPAFIndicator()">
						</span>
		</span>
					<span style="LEFT: 0px; POSITION: absolute; TOP: 184px" class="msgLabel">
			County
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtCurrentAddressCounty" maxlength="40" style="POSITION: absolute; WIDTH: 200px"
								class="msgTxt" onchange="ClearCurrentAddressPAFIndicator()">
						</span>
		</span>
					<span style="LEFT: 0px; POSITION: absolute; TOP: 208px" class="msgLabel">
			Country
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<select id="cboCurrentAddressCountry" style="WIDTH: 200px" class="msgReadOnly" 
							readonly tabindex="-1" onchange="ClearCurrentAddressPAFIndicator()">
							</select>
						</span>
		</span>
				</span>
				<% /* Correspondence Address */ %>
				<span id="spnCorrespondenceAddress" style="LEFT: 320px; POSITION: absolute; TOP: 4px">
					<span style="LEFT: 0px; POSITION: absolute; TOP: 0px" class="msgLabel">
						<strong>Correspondence Address<br>
							(if different from Current Address)</strong>
					</span>
					<span style="LEFT: 204px; POSITION: absolute; TOP: 6px">
						<input id="btnCorrespondenceClearAddress" value="Clear" type="button" style="WIDTH: 75px"
							class="msgButton">
					</span>
					<span style="LEFT: 0px; POSITION: absolute; TOP: 40px" class="msgLabel">
			Postcode
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtCorrespondenceAddressPostcode" maxlength="8" style="POSITION: absolute; WIDTH: 70px"
								class="msgTxtUpper" onchange="ClearCorrespondenceAddressPAFIndicator()">
						</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 64px" class="msgLabel">
			Flat No./ Name
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtCorrespondenceAddressFlatNumber" maxlength="10" style="POSITION: absolute; WIDTH: 70px"
								class="msgTxt" onchange="ClearCorrespondenceAddressPAFIndicator()">
						</span>
		</span>
		
		
		
	<% /* EP8 pct 13/03/2006 */ %>
	<span style="LEFT: 200px; POSITION: absolute; TOP: 64px" class="msgLabel">
		BFPO
		<span style="LEFT: 10px; POSITION: absolute; TOP: -3px">
			<input id="chkCorrespondenceAddressBFPO" type="checkbox" name="BFPO" tabIndex="-1" style="POSITION: absolute; WIDTH: 70px" onclick="EnableDisableCorrespondenceAddressBFPO()">
		</span>
	</span>
	<% /* EP8 pct 13/03/2006 - End */ %>
		
		
		
		
	
		
		
		
		<span style="LEFT: 0px; POSITION: absolute; TOP: 82px" class="msgLabel">
			House<br>No. &amp; Name
			<span style="LEFT: 80px; POSITION: absolute; TOP: 3px">
							<input id="txtCorrespondenceAddressHouseNumber" maxlength="10" style="POSITION: absolute; WIDTH: 45px"
								class="msgTxt" onchange="ClearCorrespondenceAddressPAFIndicator()">
						</span>
			<span style="LEFT: 130px; POSITION: absolute; TOP: 3px">
							<input id="txtCorrespondenceAddressHouseName" maxlength="40" style="POSITION: absolute; WIDTH: 150px"
								class="msgTxt" onchange="ClearCorrespondenceAddressPAFIndicator()">
						</span>

		</span>
					<span style="LEFT: 179px; POSITION: absolute; TOP: 35px">
						<input id="btnCorrespondenceAddressPAFSearch" value="P.A.F. Search" type="button" style="WIDTH: 100px"
							class="msgButton">
					</span>
					<span style="LEFT: 0px; POSITION: absolute; TOP: 112px" class="msgLabel">
			Street
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtCorrespondenceAddressStreet" maxlength="40" style="POSITION: absolute; WIDTH: 200px"
								class="msgTxt" onchange="ClearCorrespondenceAddressPAFIndicator()">
						</span>
		</span>
					<span style="LEFT: 0px; POSITION: absolute; TOP: 136px" class="msgLabel">
			District
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtCorrespondenceAddressDistrict" maxlength="40" style="POSITION: absolute; WIDTH: 200px"
								class="msgTxt" onchange="ClearCorrespondenceAddressPAFIndicator()">
						</span>
		</span>
					<span style="LEFT: 0px; POSITION: absolute; TOP: 160px" class="msgLabel">
			Town
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtCorrespondenceAddressTown" maxlength="40" style="POSITION: absolute; WIDTH: 200px"
								class="msgTxt" onchange="ClearCorrespondenceAddressPAFIndicator()">
						</span>
		</span>
					<span style="LEFT: 0px; POSITION: absolute; TOP: 184px" class="msgLabel">
			County
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtCorrespondenceAddressCounty" maxlength="40" style="POSITION: absolute; WIDTH: 200px"
								class="msgTxt" onchange="ClearCorrespondenceAddressPAFIndicator()">
						</span>
		</span>
					<span style="LEFT: 0px; POSITION: absolute; TOP: 208px" class="msgLabel">
			Country
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<select id="cboCorrespondenceAddressCountry" style="WIDTH: 200px" class="msgCombo" onchange="ClearCorrespondenceAddressPAFIndicator()">
							</select>
						</span>
		</span>
				</span>
			</div>
			<% /* Contact Details */ %>
			<div style="HEIGHT: 170px; LEFT: 10px; POSITION: absolute; TOP: 450px; WIDTH: 604px"
				class="msgGroup">
				<span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel">
					<strong>Contact Details</strong>
				</span>
				<span style="LEFT: 4px; POSITION: absolute; TOP: 24px" class="msgLabel">
		Type
		<span style="LEFT: 90px; POSITION: absolute; TOP: 0px" class="msgLabel"> 
			Country
		</span>
		<span style="LEFT: 90px; POSITION: absolute; TOP: 10px" class="msgLabel">
			Code
		</span>
		<span style="LEFT: 150px; POSITION: absolute; TOP: 0px" class="msgLabel">
			Area Code
		</span>
		<span style="LEFT: 210px; POSITION: absolute; TOP: 0px" class="msgLabel">
			Telephone Number
		</span>
		<span style="LEFT: 305px; POSITION: absolute; TOP: 0px" class="msgLabel">
			&nbsp;Work Extension
		</span>
		<span style="LEFT: 395px; POSITION: absolute; TOP: 0px" class="msgLabel">
			Contact Time
		</span>
		<span style="LEFT: 545px; POSITION: absolute; TOP: -14px" class="msgLabel">
			Preferred<BR>Contact ?
		</span>
	</span>
				<span style="LEFT: 4px; POSITION: absolute; TOP: 48px">
					<select id="cboType1" style="WIDTH: 80px" class="msgCombo">
					</select>
					<span style="LEFT: 90px; POSITION: absolute; TOP: 0px">
						<input id="txtCountryCode1" maxlength="3" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
					</span>
					<span style="LEFT: 148px; POSITION: absolute; TOP: 0px">
						<input id="txtAreaCode1" maxlength="6" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
					</span>
					<span style="LEFT: 208px; POSITION: absolute; TOP: 0px">
						<input id="txtTelNumber1" maxlength="30" style="POSITION: absolute; WIDTH: 90px" class="msgTxt">
					</span>
					<span style="LEFT: 305px; POSITION: absolute; TOP: 0px">
						<input id="txtExtensionNumber1" maxlength="8" style="POSITION: absolute; WIDTH: 80px" class="msgTxt">
					</span>
					<span style="LEFT: 395px; POSITION: absolute; TOP: 0px">
						<input id="txtTime1" maxlength="30" style="POSITION: absolute; WIDTH: 140px" class="msgTxt">
					</span>
					<span style="LEFT: 558px; POSITION: absolute; TOP: 0px">
						<input id="optPREFERREDCONTACT1" name="RadioGroup" type="radio" value="1">
					</span>
				</span>
				<span style="LEFT: 4px; POSITION: absolute; TOP: 72px">
					<select id="cboType2" style="WIDTH: 80px" class="msgCombo">
					</select>
					<span style="LEFT: 90px; POSITION: absolute; TOP: 0px">
						<input id="txtCountryCode2" maxlength="3" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
					</span>
					<span style="LEFT: 148px; POSITION: absolute; TOP: 0px">
						<input id="txtAreaCode2" maxlength="6" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
					</span>
					<span style="LEFT: 208px; POSITION: absolute; TOP: 0px">
						<input id="txtTelNumber2" maxlength="30" style="POSITION: absolute; WIDTH: 90px" class="msgTxt">
					</span>
					<span style="LEFT: 305px; POSITION: absolute; TOP: 0px">
						<input id="txtExtensionNumber2" maxlength="8" style="POSITION: absolute; WIDTH: 80px" class="msgTxt">
					</span>
					<span style="LEFT: 395px; POSITION: absolute; TOP: 0px">
						<input id="txtTime2" maxlength="30" style="POSITION: absolute; WIDTH: 140px" class="msgTxt">
					</span>
					<span style="LEFT: 558px; POSITION: absolute; TOP: 0px">
						<input id="optPREFERREDCONTACT2" name="RadioGroup" type="radio" value="1">
					</span>
				</span>
				<span style="LEFT: 4px; POSITION: absolute; TOP: 96px">
					<select id="cboType3" style="WIDTH: 80px" class="msgCombo">
					</select>
					<span style="LEFT: 90px; POSITION: absolute; TOP: 0px">
						<input id="txtCountryCode3" maxlength="3" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
					</span>
					<span style="LEFT: 148px; POSITION: absolute; TOP: 0px">
						<input id="txtAreaCode3" maxlength="6" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
					</span>
					<span style="LEFT: 208px; POSITION: absolute; TOP: 0px">
						<input id="txtTelNumber3" maxlength="30" style="POSITION: absolute; WIDTH: 90px" class="msgTxt">
					</span>
					<span style="LEFT: 305px; POSITION: absolute; TOP: 0px">
						<input id="txtExtensionNumber3" maxlength="8" style="POSITION: absolute; WIDTH: 80px" class="msgTxt">
					</span>
					<span style="LEFT: 395px; POSITION: absolute; TOP: 0px">
						<input id="txtTime3" maxlength="30" style="POSITION: absolute; WIDTH: 140px" class="msgTxt">
					</span>
					<span style="LEFT: 558px; POSITION: absolute; TOP: 0px">
						<input id="optPREFERREDCONTACT3" name="RadioGroup" type="radio" value="1">
					</span>
				</span>
				<span style="LEFT: 4px; POSITION: absolute; TOP: 120px">
					<select id="cboType4" style="WIDTH: 80px" class="msgCombo">
					</select>
					<span style="LEFT: 90px; POSITION: absolute; TOP: 0px">
						<input id="txtCountryCode4" maxlength="3" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
					</span>
					<span style="LEFT: 148px; POSITION: absolute; TOP: 0px">
						<input id="txtAreaCode4" maxlength="6" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
					</span>
					<span style="LEFT: 208px; POSITION: absolute; TOP: 0px">
						<input id="txtTelNumber4" maxlength="30" style="POSITION: absolute; WIDTH: 90px" class="msgTxt">
					</span>
					<span style="LEFT: 305px; POSITION: absolute; TOP: 0px">
						<input id="txtExtensionNumber4" maxlength="8" style="POSITION: absolute; WIDTH: 80px" class="msgTxt">
					</span>
					<span style="LEFT: 395px; POSITION: absolute; TOP: 0px">
						<input id="txtTime4" maxlength="30" style="POSITION: absolute; WIDTH: 140px" class="msgTxt">
					</span>
					<span style="LEFT: 558px; POSITION: absolute; TOP: -3px">
						<input id="optPREFERREDCONTACT4" name="RadioGroup" type="radio" value="1">
					</span>
				</span>
				<span style="LEFT: 4px; POSITION: absolute; TOP: 148px" class="msgLabel">
		E-Mail Address
		<span style="LEFT: 90px; POSITION: absolute; TOP: 0px">
						<input id="txtContactEMailAddress" maxlength="100" style="POSITION: absolute; WIDTH: 445px"
							class="msgTxt">
					</span>
		<span style="LEFT: 558px; POSITION: absolute; TOP: -3px">
						<input id="optPREFERREDCONTACT5" name="RadioGroup" type="radio" value="1">
					</span>
	</span>
			</div>
			<%
/* ASu Removed As per BMIDS00139 Start
 // Areas of Interest //
<div style="TOP: 626px; LEFT: 10px; HEIGHT: 40px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 13px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Please enter your Areas of Interest
		<span style="TOP: -9px; LEFT: 180px; POSITION: ABSOLUTE">
			<input id="ddAreasOfInterest" value="" type="button" style="WIDTH: 32px; HEIGHT: 32px" class="msgDDButton">
		</span>
	</span>
</div>
 ASu Finish
*/ 
%>
		</form>
		<%/* Main Buttons */ %>
		<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 623px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" -->
		</div>
		<% /* Span to keep tabbing within this screen */ %>
		<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->  <!-- #include FILE="includes/pafsearch.asp" -->
		<% /*File containing field attributes (remove if not required) - see xx999attribs.asp for example code */ %>
		<!--  #include FILE="attribs/cr020attribs.asp" -->
		<% /* Specify Code Here */ %>
		<script language="JScript">
<!--
<%//GD Added 20.02.2001 SYS1752
%>
var m_sOtherSystemCustomerNumber = null;
var m_sDateOfBirth = null;
<%
/* ASu Removed As per BMIDS00139 Start
var m_sXMLAreasOfInterest = null;  Stores XML string for Areas Of Interest  
ASu End */
%>
var m_bCurrentAddressPAFIndicator = false;
var m_bCorrespondenceAddressPAFIndicator = false;
var m_sCustomerNumber = null;
var m_sCustomerVersionNumber = null;
var m_sMetaAction = null;
var m_sContext = null;
var m_sPackageNumber = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sTypeOfApplicationValue = null;
var m_sReadOnly = null; <% /* APS UNIT TEST REF 23 */ %>
var m_sCustomerLockIndicator = null; <% /* APS UNIT TEST REF 102 */ %>
var m_sOtherSystemCustomerNo ="";
var m_sCurrentAddressGUID = null;
var m_sCorrespondenceAddressGUID = null;
var m_sCurrentCustAddressSeqNo = null
var m_sCorrespondenceCustAddressSeqNo = null
var m_bIsSubmit = false;
var scScreenFunctions;
var m_blnReadOnly = false;
<% /* ASu BMIDS00690 - Start. Set trigger to check for 'Copy'Function in get telephone details */ %>
var m_sCopy = false;
<% /* ASu - End */ %>
<% /* PJO BM0006 */ %>
var m_iMaxApplicants = null ;
var m_iMaxCustomers = null ;
var m_bNewToEIndicator = 0;         // BMIDS758
var m_bNewCustomer = false;

//SG 29/05/02 SYS4767 START
//SG 20/03/02 MSMS0013
var m_sCustNo1 = null;
var m_sCustVersNo1 = null;
var m_sCustRoleType1 = null;
//SG 20/03/02 MSMS0013
//SG 29/05/02 SYS4767 END
<%/* SG 25/06/02 SYS4930 */%>
var m_bIsPopup = false;

<% /* PSC 13/10/2005 MAR57 */ %>
var ComboXML = null
var m_sCustomerCategory = "";

<% /* EP903 */ %>
var m_sApplicationDate = "";

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	/* CORE UPGRADE 702 scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();*/
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Customer Details","CR020",scScreenFunctions);
	RetrieveContextData();
<%	//This function is contained in the field attributes file (remove if not required)
%>	SetMasks(m_bNewCustomer);
	GetComboLists();	
	PopulateScreen();
	HideOtherTitle();	
	Initialise();
	<% /* PJO BM0006 - Read info from Global Pars */ %>
	GetMaximumApplicantsData() ;
	HideCustomerNumber(); <%/* Hide label and text field of CustomerNumber, if CustomerNumber is null */%>
	Validation_Init();	
	FocusFirstField();	
	if (ContinueOnReadOnly()) SetScreenToReadOnly();
	else
	{
		ResetCustomerContext();
		frmToCR010.submit();
	}
m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "CR020");	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

<% /* MF 26/07/2005 choose a field to focus first. Surname if new customer
DOB if existing customer but value not entered, otherwise first phone no. entry */ %>		
function FocusFirstField()
{
	if(m_bNewCustomer){
		frmScreen.txtSurname.focus();
	}else if(frmScreen.txtDateOfBirth.value == ""){
		frmScreen.txtDateOfBirth.focus();
	}else {
		frmScreen.cboType1.focus();
	}
}
<% /* MF 26/07/2005 Disable most fields if this is an existing customer */ %>	
function MakeFieldsReadOnly()
{		
	var aFields = new Array (
	<% /* Personal details fields */ %>
	"txtSurname", "txtOtherSysCustomerNo", "txtFirstForename", "cboTitle",
	<% /* PSC 30/01/2006 MAR939 */ %>
	"txtOtherTitle", "txtAge", "cboGender",	"cboSpecialNeeds", 		
	<% /*Adress fields */ %>		
	"txtCurrentAddressPostcode", "txtCurrentAddressFlatNumber",
	"txtCurrentAddressFlatNumber", "txtCurrentAddressFlatNumber",
	"txtCurrentAddressHouseName", "txtCurrentAddressHouseNumber",
	"txtCurrentAddressStreet", "txtCorrespondenceAddressCounty", 
	"txtCurrentAddressDistrict", "txtCurrentAddressTown",
	"txtCurrentAddressCounty", "cboCurrentAddressCountry",
	"txtCorrespondenceAddressPostcode", "cboCorrespondenceAddressCountry", 
	"txtCorrespondenceAddressFlatNumber", "txtCorrespondenceAddressHouseName",
	"txtCorrespondenceAddressHouseNumber", "txtCorrespondenceAddressTown",
	"txtCorrespondenceAddressStreet", "txtCorrespondenceAddressDistrict");
	<% /* Disable buttons */ %>
	frmScreen.btnCorrespondenceAddressPAFSearch.disabled = true;
	frmScreen.btnCurrentAddressPAFSearch.disabled = true;
	frmScreen.btnCorrespondenceClearAddress.disabled = true;
	frmScreen.btnCurrentAddressClearAddress.disabled = true;
	
	for(var i=0;i<aFields.length;i++){						
		document.all(aFields[i]).disabled=true;			
		switch (document.all(aFields[i]).tagName){
			case "select":
				document.all(aFields[i]).className="msgReadOnlyCombo";
				break;				
			default:
				document.all(aFields[i]).className="msgReadOnly";
				break;
		}						
	}			
}

function Initialise()
{
	frmScreen.cboType1.onchange()
	frmScreen.cboType2.onchange()
	frmScreen.cboType3.onchange()
	frmScreen.cboType4.onchange()
	frmScreen.txtContactEMailAddress.onchange();
	frmScreen.txtCurrentAddressPostcode.setAttribute("required", "true");
<%
/* ASu Removed As per BMIDS00139 Start
	scScreenFunctions.DisableDrillDown(frmScreen.ddAreasOfInterest) 
 ASu End */
%>
}

<% /* SR 11/06/01 : SYS2362- Hide CustomerNumber field, if it is null  */ %>
function HideCustomerNumber()
{
	if(frmScreen.txtOtherSysCustomerNo.value == "")
	{
		scScreenFunctions.HideCollection(spnOtherSysCustomerNo);	
	}
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function SetScreenToReadOnly()
{
<%	// Set all controls on the screen to read only
	// The buttons are treated separately
%>	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);

		frmScreen.btnCorrespondenceAddressPAFSearch.disabled = true;
		frmScreen.btnCurrentAddressPAFSearch.disabled = true;
		frmScreen.btnCorrespondenceClearAddress.disabled = true;
		frmScreen.btnCurrentAddressClearAddress.disabled = true;
		
		//SG 28/02/02 SYS4767
		frmScreen.btnCopy.disabled = true;	

		if (m_sMetaAction == "CreateExistingCustomerForExistingApplication" || m_sMetaAction == "UpdateExistingCustomer")
			DisableMainButton("Submit");
	}
}

function GetComboLists()
{
	/* CORE UPGRADE 702 var XML = new scXMLFunctions.XMLObject(); */
	ComboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("Title", "Sex", "Country", "TelephoneUsage","SpecialNeeds", "CustomerCategory");
	var bSuccess = false;

	if(ComboXML.GetComboLists(document,sGroups))
	{
		bSuccess = true;
		bSuccess = bSuccess & ComboXML.PopulateCombo(document,frmScreen.cboTitle,"Title",true);
		bSuccess = bSuccess & ComboXML.PopulateCombo(document,frmScreen.cboGender,"Sex",true);
		bSuccess = bSuccess & ComboXML.PopulateCombo(document,frmScreen.cboCurrentAddressCountry,"Country",false); // PAK MAR1443
		//EP594: current address must be a UK address
		frmScreen.cboCurrentAddressCountry.value = 1;
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboCurrentAddressCountry");
//	    bSuccess = bSuccess & ComboXML.PopulateCombo(document,frmScreen.cboCurrentAddressCountry,"Country",true);
//		bSuccess = bSuccess & ComboXML.PopulateCombo(document,frmScreen.cboCorrespondenceAddressCountry,"Country",false);
		bSuccess = bSuccess & ComboXML.PopulateCombo(document,frmScreen.cboCorrespondenceAddressCountry,"Country",true);
		bSuccess = bSuccess & ComboXML.PopulateCombo(document,frmScreen.cboType1,"TelephoneUsage",true);
		bSuccess = bSuccess & ComboXML.PopulateCombo(document,frmScreen.cboType2,"TelephoneUsage",true);
		bSuccess = bSuccess & ComboXML.PopulateCombo(document,frmScreen.cboType3,"TelephoneUsage",true);
		bSuccess = bSuccess & ComboXML.PopulateCombo(document,frmScreen.cboType4,"TelephoneUsage",true);
		bSuccess = bSuccess & ComboXML.PopulateCombo(document,frmScreen.cboSpecialNeeds,"SpecialNeeds",true);
	}

	if(!bSuccess)
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		DisableMainButton("Submit");
	}
}

function frmScreen.btnCurrentAddressClearAddress.onclick()
{
	FlagChange(scScreenFunctions.ClearCollection(spnCurrentAddress));
	frmScreen.txtCurrentAddressPostcode.focus();
}

function frmScreen.btnCurrentAddressPAFSearch.onclick()
{
	with (frmScreen)
		m_bCurrentAddressPAFIndicator = PAFSearch(txtCurrentAddressPostcode,txtCurrentAddressHouseName,
			txtCurrentAddressHouseNumber,txtCurrentAddressFlatNumber,txtCurrentAddressStreet,txtCurrentAddressDistrict,
			txtCurrentAddressTown,txtCurrentAddressCounty,cboCurrentAddressCountry);
}

function frmScreen.btnCorrespondenceClearAddress.onclick()
{
<%	// APS 09/09/1999 - UNIT TEST REF 73
	// Add clear button to clear out possible wrong PAF Search
	// AY 15/09/99 - clearing of correspondence address fields now handled by a common function
	// Make sure that any changes of field content are flagged
%>	FlagChange(scScreenFunctions.ClearCollection(spnCorrespondenceAddress));
	frmScreen.txtCorrespondenceAddressPostcode.focus();
}

function frmScreen.btnCorrespondenceAddressPAFSearch.onclick()
{
	with (frmScreen)
		m_bCorrespondenceAddressPAFIndicator = PAFSearch(txtCorrespondenceAddressPostcode,
			txtCorrespondenceAddressHouseName,
			txtCorrespondenceAddressHouseNumber,txtCorrespondenceAddressFlatNumber,txtCorrespondenceAddressStreet,
			txtCorrespondenceAddressDistrict,txtCorrespondenceAddressTown,txtCorrespondenceAddressCounty,
			cboCorrespondenceAddressCountry);
}

<%
/* ASu Removed As per BMIDS00139 Start.
function frmScreen.ddAreasOfInterest.onclick()
{
	 APS UNIT TEST REF 11 & 12 - Passing through the CustomerNumber and CustomerVersionNumber
	 we cannot use the context form as this asp page is a popup
	 APS UNIT TEST REF 72 - Read Only processing

	If a value is returned from cr060, OK was pressed so we need to change the Areas Of Interest string
	 AY 15/09/99 - CR060 now returns null if cancelled or an array if OK'd
	               Array[0] returns whether there was a change made on CR060
	               Array[1] returns the XML string
	var sReturn = null;

	var ArrayArguments = new Array(3);
	ArrayArguments[0] = m_sXMLAreasOfInterest;
	ArrayArguments[1] = m_sCustomerNumber;
	ArrayArguments[2] = m_sCustomerVersionNumber;
	ArrayArguments[3] = m_sReadOnly;
	ArrayArguments[4] = frmScreen.txtFirstForename.value + " " + frmScreen.txtSurname.value;

	sReturn = scScreenFunctions.DisplayPopup(window, document, "cr060.asp", ArrayArguments, 630, 394);
	if(sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_sXMLAreasOfInterest = sReturn[1];
	}
}
ASu End */
%>

function frmScreen.cboTitle.onchange()
{
	HideOtherTitle();
	DefaultGenderFromTitle();
}

function HideOtherTitle()
{
	if(scScreenFunctions.IsValidationType(frmScreen.cboTitle,"O")) scScreenFunctions.ShowCollection(spnOtherTitle);
	else scScreenFunctions.HideCollection(spnOtherTitle);
}

function DefaultGenderFromTitle()
{
	if(scScreenFunctions.IsValidationType(frmScreen.cboTitle,"M")) frmScreen.cboGender.value = "1";
	else if(scScreenFunctions.IsValidationType(frmScreen.cboTitle,"F")) frmScreen.cboGender.value = "2";
	else frmScreen.cboGender.value = "";
	
	if(scScreenFunctions.IsValidationType(frmScreen.cboTitle,"O")) 
		frmScreen.txtOtherTitle.focus();
}

function frmScreen.txtOtherTitle.onblur()
{
	if(scScreenFunctions.IsValidationType(frmScreen.cboTitle,"O") && frmScreen.txtOtherTitle.value == "") 
	{
//		frmScreen.txtOtherTitle.focus(); 
		alert("Please enter the 'Other' Title or Change Title selection ");
		frmScreen.cboTitle.focus();
		return;
	}
}

function btnSubmit.onclick()
{
<%	// APS UNIT TEST REF 34
	// AY 15/09/99 - don't perform commit if no changes have been made
	// AY 23/09/99 - Make sure an existing customer is added to the package if
	// no details are changed on the screen
	// AY 14/03/00 SYS0050 - if there is nothing in CustomerName1, populate it
	// i.e. we're not in an application yet
%>		
	var bTelephoneDetailsOk;
	
	<% /* BM0077 Added to skip over validation but still do routing if the screen is readonly */ %>
	var blnDoRouting = false;
	if (m_sReadOnly == "1")
		blnDoRouting = true;
	else
	{
	<% /* BM0077 End */ %>

		<% /* SR 29/11/2002 : BMIDS01048 */ %>
		if(frmScreen.onsubmit()) 
		{
			<% /* SYS1053 Prevent the user from inputting number as first character of surname + forenames */ %>
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
			<% /*ASu BMIDS00453 - Start */ %>
			if(scScreenFunctions.IsValidationType(frmScreen.cboTitle,"O") && frmScreen.txtOtherTitle.value == "") 
			{
				alert("Please enter the 'Other' Title or Change Title selection ");
				frmScreen.cboTitle.focus();
				return false;
			}
			<% /*ASu End */ %>
			if(m_bIsSubmit) return;
			m_bIsSubmit = true;	

			<% /* BMIDS00004 MDC 21/05/2002 - If required created new customer version */ %>
			if (m_sMetaAction == "CreateExistingCustomerForExistingApplication")
			{
				<% /* Create new customer version... */ %>
				<% /* MV - 02/07/2002 - Core Upgrade Code Error - BMIDS000120 
				var CustXML = new scXMLFunctions.XMLObject(); */ %>
				var CustXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				CustXML.CreateRequestTag(window,null);
				CustXML.CreateActiveTag("SEARCH");
				CustXML.CreateActiveTag("CUSTOMER");
				CustXML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
				CustXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
				CustXML.CreateTag("NEWTOECUSTOMERIND", m_bNewToEIndicator);  // BMIDS836

				// 			CustXML.RunASP(document, "CreateNewCustomerVersion.asp");
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules()) //!!
				{
					case 1: // Warning
					case 0: // OK
							CustXML.RunASP(document, "CreateNewCustomerVersion.asp");
							break;
					default: // Error
						CustXML.SetErrorResponse();
				}

				if(CustXML.IsResponseOK())
				{
					<% /* ...and update context with the new customer version number */ %>
					var CustVerTag = CustXML.SelectTag(null,"CUSTOMERKEY")
					m_sCustomerVersionNumber = CustXML.GetTagText("CUSTOMERVERSIONNUMBER");
				}
				else
				{
					alert("Unable to create new customer version. Please contact Help Desk.");
					return false;
				}
			}
		
			<% /* BMIDS00004 MDC 21/05/2002 - End */ %>
			if (m_sReadOnly != "1" && IsChanged())
			{
	
				<% /* SYS00618 Validate postcodes */ %>
				if (frmScreen.cboCurrentAddressCountry.value == 10 ? scScreenFunctions.ValidatePostcode(frmScreen.txtCurrentAddressPostcode) : true)
				//if (scScreenFunctions.ValidatePostcode(frmScreen.txtCurrentAddressPostcode))
				   if ((frmScreen.cboCorrespondenceAddressCountry.value == 10) && !frmScreen.chkCorrespondenceAddressBFPO.checked ? scScreenFunctions.ValidatePostcode(frmScreen.txtCorrespondenceAddressPostcode) : true)
				   //if (scScreenFunctions.ValidatePostcode(frmScreen.txtCorrespondenceAddressPostcode))
				   {
						bTelephoneDetailsOk = ValidateTelephoneDetails() ;
						if (bTelephoneDetailsOk)
						{
							if(!CommitScreen())
							{
								m_bIsSubmit = false;
								return;
							}
						}
						else
						{	
							m_bIsSubmit = false;
							return;
						}
					}
					else 
					{	
						m_bIsSubmit = false;
						return;
					}
				else 
				{	
					m_bIsSubmit = false;
					return;
				}		
			}
			<% /* BM0077 */ %>
			blnDoRouting = true;
			<% /* BM0077 End */ %>
		}
	}
	
	<% /* BM0077 */ %>
	if (blnDoRouting)
	{
	<% /* BM0077 End */ %>
		if(scScreenFunctions.CompletenessCheckRouting(window))
		{
			frmToGN300.submit();
			return;
		}

		if(CommitToPackage())
		{
			<% /* PSC 13/10/2005 MAR57 - Start */ %>
			if(scScreenFunctions.GetContextParameter(window,"idCustomerName1","") == "")
			{
				scScreenFunctions.SetContextParameter(window,"idCustomerName1",frmScreen.txtFirstForename.value + " " + frmScreen.txtSurname.value);
				scScreenFunctions.SetContextParameter(window,"idCustomerCategory1", m_sCustomerCategory); 
			}
			<% /* PSC 13/10/2005 MAR57 - End */ %>

			<%
			//GD 20.02.2001 SYS1752 added these lines.....
			%>
			<% /* PSC 23/10/2002 BMIDS00465 */ %>
			if (m_sMetaAction == "CreateNewCustomerForExistingApplication" || m_sMetaAction == "CreateExistingCustomerForExistingApplication" || m_sMetaAction == "UpdateExistingCustomer") frmToCR030.submit();

			<% /* PSC 23/10/2002 BMIDS00465 */ %>

			if (m_sMetaAction == "CreateNewCustomerForNewApplication" || m_sMetaAction == "CreateExistingCustomerForNewApplication")
				frmToCR025.submit();
			<%
			//GD 20.02.2001 SYS1752  The following 2 IFs can now be removed....
			//if (m_sMetaAction == "CreateNewCustomerForExistingApplication" || m_sMetaAction == "CreateExistingCustomerForExistingApplication"
			 //  || m_sMetaAction == "UpdateExistingCustomer") frmToCR030.submit();
			//if (m_sMetaAction == "CreateNewCustomerForNewApplication" || m_sMetaAction == "CreateExistingCustomerForNewApplication")
			//	frmToCR040.submit();
			%>
			
		}
		else m_bIsSubmit = false;
	} <% /* BM0077 */ %>
}

function ResetCustomerContext()
{
<%	/*	AY 29/02/00 SYS0318 - clear the read only flags */
%>	m_sCustomerNumber = null;
	m_sCustomerVersionNumber = null;
	m_sReadOnly = null;
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber",m_sCustomerNumber);
	scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber",m_sCustomerVersionNumber);
	scScreenFunctions.SetContextParameter(window,"idCustomerReadOnly",m_sReadOnly);
	scScreenFunctions.SetContextParameter(window,"idReadOnly",m_sReadOnly);
	<% /*MV - 04/07/2002 - BMIDS00100 - Resting the values  */%>

	<% /* SR 28/11/2002: BMIDS01048 - Clear CustomerName1 from context only when it corresponds to CustomerNumber 
						 but not to CustomerNumber1, i.e., only when CustomerNumber1 is empty 	*/
	 %>
	<% /* PSC 13/10/2005 MAR57 - Start */ %>
	if(scScreenFunctions.GetContextParameter(window,"idCustomerNumber1")== "")
	{
		scScreenFunctions.SetContextParameter(window,"idCustomerName1", ""); 
		scScreenFunctions.SetContextParameter(window,"idCustomerCategory1", ""); 
	}
	<% /* PSC 13/10/2005 MAR57 - Start */ %>

	scScreenFunctions.SetContextParameter(window,"idCustomerName", "");
	<% /* SR 28/11/2002 BMIDS01048 - End */ %>
}

function btnCancel.onclick()
{
<%	// APS UNIT TEST REF 23 - set the current customer context to null because 
	// we are not adding this customer to an application so we forget him/her
	// APS UNIT TEST REF 34 - Delete the Customer lock iff not in read only mode
%>	
	if(scScreenFunctions.CompletenessCheckRouting(window))
	{
		frmToGN300.submit();
		return;
	}

	if(m_sMetaAction == "UpdateExistingCustomer") frmToCR030.submit();
	else
	{
		<% /* BMIDS00393 MDC 05/09/2002 */ %>
		<% /* if(m_sMetaAction == "CreateExistingCustomerForNewApplication" || m_sMetaAction == "CreateExistingCustomerForExistingApplication") */ %>
		if(m_sMetaAction == "CreateExistingCustomerForNewApplication" || 
			m_sMetaAction == "CreateExistingCustomerForExistingApplication" ||
			(m_sMetaAction == "CreateNewCustomerForNewApplication" && m_sCustomerNumber != ""))
		{
			if (m_sReadOnly != "1") DeleteCustomerLock();
			ResetCustomerContext();
		}
		else
		{<% /* SR 28/11/2002: BMIDS01048 - Clear CustomerName1 from context only when it corresponds to CustomerNumber 
								 but not to CustomerNumber1, i.e., only when CustomerNumber1 is empty 	*/
		 %>
			if(scScreenFunctions.GetContextParameter(window,"idCustomerNumber1")== "")
				scScreenFunctions.SetContextParameter(window,"idCustomerName1", ""); 
		
			scScreenFunctions.SetContextParameter(window,"idCustomerName", "");
		} <% /* SR 28/11/2002: BMIDS01048 - End */ %>

		<% /* MV : 22/05/2002 :  BMIDS00005 - BM052 Customer Searches */ %>
		scScreenFunctions.SetContextParameter(window,"idSearchFlag", "1");
		<% /*MV - 04/07/2002 - BMIDS00100 - Resting the values */%>
			
		frmToCR010.submit();
	}
}

function DeleteCustomerLock()
{
<%	// APS UNIT TEST REF 34 - Delete the Customer lock by calling
	// CustomerlockBO.Delete
%>	var blnReturn = false;
	/* CORE UPGRADE 702 var XML = new scXMLFunctions.XMLObject(); */
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	XML.CreateRequestTag(window,"DELETE");

	XML.CreateActiveTag("CUSTOMERLOCK");

	m_sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber");
	
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.RunASP(document,"DeleteCustomerLock.asp")
	
	if (XML.IsResponseOK()) blnReturn = true;
	XML = null;

	return blnReturn;
}

function PopulateScreen()
{
	//SG 29/05/02 SYS4767 START
	//SG 20/03/02 MSMS0013 - Changed IF statement
	//SG 28/02/02 SYS4186
	<% /* MV - 05/07/2002 - BMIDS00096 - Copy button shud be disables if there is only 1 customer */ %>
	<% /* ASu BMIDS00690 - Start. Add additional checking */ %>
	<% /* PSC 30/11/2005 MAR729 */ %>
	//SG 28/02/02 SYS4186
	//SG 29/05/02 SYS4767 END
	<% /* MF 26/07/2005 MARS MAR019 Is this a new customer */ %>	

	<% /* PE - 25/01/2006 - MAR939 - Parameterised the disabling of the customer screens	*/ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bEnable = !XML.GetGlobalParameterBoolean(document,"DisableCustomerDetails");			
	if(m_sMetaAction == "CreateNewCustomerForNewApplication" || m_sMetaAction == "CreateNewCustomerForExistingApplication" || bEnable){		
		<% /* PSC 30/11/2005 MAR729 */ %>
		frmScreen.btnCopy.disabled = false;	
		<% /* MF 27/07/2005 Make Surname and first forename mandatory for new customers*/ %>
		frmScreen.txtSurname.setAttribute("required", "true");
		frmScreen.txtFirstForename.setAttribute("required", "true");
		<% /* MF 22/09/2005 Make title, gender, 'Special Needs' mandatory for new customers*/ %>
		frmScreen.cboGender.setAttribute("required", "true");
		<% /* DS :  EP2_610
		frmScreen.cboSpecialNeeds.setAttribute("required", "true");
		*/ %>
		<% /* MF 23/09/2005 MAR19 WP01 modifications. default special needs combo */ %>
		// EP2_610 - Remove this default value - Now SELECT.
		//scScreenFunctions.SetComboOnValidationType(frmScreen, "cboSpecialNeeds", "N"); 

		frmScreen.cboTitle.setAttribute("required", "true");

	} else {
		<% /* PSC 30/11/2005 MAR729 */ %>
		frmScreen.btnCopy.disabled = true;	
		MakeFieldsReadOnly();	
	}		
	<% /* BMIDS00393 MDC 03/09/2002 */ %>
	if(m_sMetaAction == "CreateExistingCustomerForNewApplication" || 
		m_sMetaAction == "CreateExistingCustomerForExistingApplication" ||
	   m_sMetaAction == "UpdateExistingCustomer" ||
	   m_sContext == "CompletenessCheck" ||
		(m_sMetaAction == "CreateNewCustomerForNewApplication" && m_sCustomerNumber != ""))
	{

		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window,"SEARCH");
		XML.CreateActiveTag("SEARCH");
		XML.CreateActiveTag("CUSTOMER");
		XML.CreateTag("CUSTOMERNUMBER",m_sCustomerNumber);

		<%
		//GD Added OTHERSYSTEMCUSTOMERNUMBER SYS1752 20.02.2001
		%>
		XML.CreateTag("OTHERSYSTEMCUSTOMERNUMBER",m_sOtherSystemCustomerNumber);
		XML.CreateTag("CUSTOMERVERSIONNUMBER",m_sCustomerVersionNumber);
		
		if (((m_sMetaAction == "CreateExistingCustomerForNewApplication") || 
				(m_sMetaAction == "CreateExistingCustomerForExistingApplication"))
			&& (m_sCustomerLockIndicator != "1")) 
		{	
			XML.RunASP(document,"GetAndSyncCustomerDetails.asp");
			
			<%//GD ADDED 21.02.2001  SYS1752 %>
			if(XML.SelectTag(null,"CUSTOMER") != null)
			{

				if (scScreenFunctions.GetContextParameter(window,"idCustomerNumber")=="")
				{
					scScreenFunctions.SetContextParameter(window,"idCustomerNumber",XML.GetTagText("CUSTOMERNUMBER"));
					
					<% /* SYS0003 - Store the CustomerNo as a module-level as that's whats saved. */ %>
					m_sCustomerNumber = XML.GetTagText("CUSTOMERNUMBER")
				}
				if (scScreenFunctions.GetContextParameter(window,"idOtherSystemCustomerNumber")=="")
				{
					scScreenFunctions.SetContextParameter(window,"idOtherSystemCustomerNumber",XML.GetTagText("OTHERSYSTEMCUSTOMERNUMBER"));
				}
				if (scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber")=="")
				{
					scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber",XML.GetTagText("CUSTOMERVERSIONNUMBER"));
					
					<% /* SYS0003 - Store the CustomerVersionNo as a module-level as that's whats saved. */ %>
					m_sCustomerVersionNumber = XML.GetTagText("CUSTOMERVERSIONNUMBER");
				}
			}
		}
		else XML.RunASP(document,"GetCustomerDetails.asp");
		
		if(XML.IsResponseOK())
		{
			GetReadOnlyFlag(XML);
			GetCustomerDetails(XML);
			GetAddressDetails(XML);
			GetTelephoneDetails(XML);

<%
/*ASu Removed As per BMIDS00139 Start
			XML.ActiveTag = null;
			XML.CreateTagList("AREASOFINTERESTLIST");
			if(XML.ActiveTagList.length > 0)
			{
				XML.SelectTagListItem(0);
				m_sXMLAreasOfInterest = XML.ActiveTag.xml;
			}
ASu End */
%>
		}

		XML = null;
	}
<%
/* ASu Removed As per BMIDS00139 Start.
	else
	{ */
%>
<%		// APS UNIT TEST REF 12
		// Generate the default XML for new customers for the Areas of Interest
		// before the defaulted values would never have been saved to the database
		// if the user had not visited CR060
%>
<%
/*		var XMLAreasOfInterest = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		//CORE UPGRADE 702 var XMLAreasOfInterest = new scXMLFunctions.XMLObject();
		var tagAreasOfInterestList = XMLAreasOfInterest.CreateActiveTag("AREASOFINTERESTLIST");

		for (var iLoop=1; iLoop <= 3; iLoop++)
		{
			XMLAreasOfInterest.ActiveTag = tagAreasOfInterestList;
			XMLAreasOfInterest.CreateActiveTag("AREASOFINTEREST");
			XMLAreasOfInterest.CreateTag("INTERESTAREA",iLoop.toString());
		}

		m_sXMLAreasOfInterest = tagAreasOfInterestList.xml;
		XMLAreasOfInterest = null;
	}
 ASu End */
%>
}

function ContinueOnReadOnly()
{
<%	// APS UNIT TEST 72 - ReadOnly processing
	// APS UNIT TEST REF 86
	// APS 10/09/99 - UNIT TEST REF 102 Setting CustomerLockIndicator = true
%>	var blnContinue = true;

	if ((m_sMetaAction == "CreateExistingCustomerForNewApplication") || (m_sMetaAction == "CreateExistingCustomerForExistingApplication"))
	{
		if (m_sReadOnly == "1")
		{
			var sMessage = "Customer " + m_sCustomerNumber + " is locked to another user. "

			if (m_sMetaAction == "CreateExistingCustomerForExistingApplication")
			{
				sMessage = sMessage + "You cannot add this customer to the application until they have been released."
				alert(sMessage);
			}

			if (m_sMetaAction == "CreateExistingCustomerForNewApplication")
			{											
				sMessage = sMessage + "Do you wish to continue?"
				blnContinue = confirm(sMessage);
			}
		}
		else
		{
			m_sCustomerLockIndicator = "1";
			scScreenFunctions.SetContextParameter(window,"idCustomerLockIndicator","1");
		}
	}

	return blnContinue;
}

function GetReadOnlyFlag(XML)
{
<%	// Sets the read only flag from the attribute in the response block
	// which is set on a call to GetAndSynchroniseCustomerDetails
%>	if((m_sMetaAction == "CreateExistingCustomerForNewApplication") || (m_sMetaAction == "CreateExistingCustomerForExistingApplication"))
	{
		XML.ActiveTag = null;
		XML.CreateTagList("RESPONSE");
		if (XML.SelectTagListItem(0))
		{
			m_sReadOnly = XML.GetAttribute("READONLY");
			scScreenFunctions.SetContextParameter(window,"idCustomerReadOnly",m_sReadOnly);
			scScreenFunctions.SetContextParameter(window,"idReadOnly",m_sReadOnly);
		}
	}
}

function GetCustomerDetails(XML)
{
<%	// Retrieves the details from the <CUSTOMERVERSION> section of the returned XML
	// APS UNIT TEST REF 13	- Not storing Mothers MaidenName or email address
%>
<%
	// SR 11/06/2001 : SYS2362 - Populate OtherSystemCustomerNumber
%>
	
	XML.ActiveTag = null;
	XML.CreateTagList("CUSTOMER");
	if(XML.SelectTagListItem(0))
	{
		frmScreen.txtOtherSysCustomerNo.value = XML.GetTagText("OTHERSYSTEMCUSTOMERNUMBER");
	}
	
	XML.ActiveTag = null;
	XML.CreateTagList("CUSTOMERVERSION");

	if(XML.SelectTagListItem(0))
	{
		<% /* PSC 13/10/2005 MAR57 */ %>
		m_sCustomerCategory = XML.GetTagText("CUSTOMERCATEGORY");
		m_sCustomerCategory = ComboXML.GetComboDescriptionForValidation("CustomerCategory", m_sCustomerCategory);
		frmScreen.txtSurname.value = XML.GetTagText("SURNAME");
		frmScreen.txtFirstForename.value = XML.GetTagText("FIRSTFORENAME");		
		frmScreen.txtSecondForename.value = XML.GetTagText("SECONDFORENAME");
		
		<% /* EP225 pct 17/03/2006 */%>
		/*
		if (frmScreen.txtSecondForename.value != ""){
			frmScreen.txtSecondForename.className="msgReadOnly";
			frmScreen.txtSecondForename.readOnly=true;
			frmScreen.txtSecondForename.disabled=true;
		}
		*/
				
		frmScreen.txtOtherForenames.value = XML.GetTagText("OTHERFORENAMES");
		<% /* EP225 pct 17/03/2006 */%>
		/*
		if (frmScreen.txtOtherForenames.value != ""){
			frmScreen.txtOtherForenames.className="msgReadOnly";
			frmScreen.txtOtherForenames.readOnly=true;
			frmScreen.txtOtherForenames.disabled=true;
		}		
		*/
		
		frmScreen.cboTitle.value = XML.GetTagText("TITLE");
		frmScreen.txtOtherTitle.value = XML.GetTagText("TITLEOTHER");
		frmScreen.txtDateOfBirth.value = XML.GetTagText("DATEOFBIRTH");
		m_sDateOfBirth = frmScreen.txtDateOfBirth.value;
		
		<% /* PE - 25/01/2006 - MAR939 */ %>
		<% /* PSC 30/01/2006 MAR939 */ %>
		<% /* EP225 pct 17/03/2006 */%>
		/*
		if (m_sDateOfBirth != ""){
			frmScreen.txtDateOfBirth.className="msgReadOnly";
			frmScreen.txtDateOfBirth.readOnly=true;
			frmScreen.txtDateOfBirth.disabled=true;
		} 
		*/
		
		frmScreen.txtAge.value = XML.GetTagText("AGE");
		frmScreen.cboGender.value = XML.GetTagText("GENDER");				
		// EP2_610 - Default to Select.
		if( XML.GetTagText("SPECIALNEEDS")!= null && XML.GetTagText("SPECIALNEEDS")!= "") 		
			frmScreen.cboSpecialNeeds.value = XML.GetTagText("SPECIALNEEDS"); 
		else
			frmScreen.cboSpecialNeeds.selectedIndex = 0;	

		frmScreen.txtContactEMailAddress.value = XML.GetTagText("CONTACTEMAILADDRESS");
		frmScreen.optPREFERREDCONTACT5.checked = (XML.GetTagText("EMAILPREFERRED") == "1")? true:false;
		scScreenFunctions.SetRadioGroupValue(frmScreen, "MailshotInd", XML.GetTagText("MAILSHOTREQUIRED"));
	}
}

function GetAddressDetails(XML)
{
<%	/*	Retrieves the details from the <CUSTOMERADDRESSLIST> section of the returned XML
		As there is only 1 <CUSTOMERADDRESSLIST> and it only contains <CUSTOMERADDRESS> tags, we can access
		<CUSTOMERADDRESS> directly

		Loop through the <CUSTOMERADDRESS> tags
		ADDRESSTYPE 1 is Home Address
		ADDRESSTYPE 2 is Correspondence Address
	*/
%>	XML.ActiveTag = null;
	XML.CreateTagList("CUSTOMERADDRESS");

	for(var nLoop = 0;nLoop < XML.ActiveTagList.length;nLoop++)
	{
		XML.SelectTagListItem(nLoop);
		if(XML.GetTagText("ADDRESSTYPE") == "1")
		{
			/* CORE UPGRADE 702 m_sCurrentCustAddressSeqNo = XML.GetTagText("CUSTOMERADDRESSSEQUENCENUMBER");
			m_sCurrentAddressGUID = XML.GetTagText("ADDRESSGUID");*/
			
			//SG 29/05/02 SYS4767 START
			//SG 01/03/02 SYS4186
			//Placed if statement round existing code
			if (m_sMetaAction != "CreateNewCustomerForExistingApplication")
			{
				m_sCurrentCustAddressSeqNo = XML.GetTagText("CUSTOMERADDRESSSEQUENCENUMBER");
				m_sCurrentAddressGUID = XML.GetTagText("ADDRESSGUID");
			}
			//SG 01/03/02 SYS4186
			//SG 29/05/02 SYS4767 END
			
			frmScreen.txtCurrentAddressPostcode.value = XML.GetTagText("POSTCODE");
			frmScreen.txtCurrentAddressFlatNumber.value = XML.GetTagText("FLATNUMBER");
			frmScreen.txtCurrentAddressHouseName.value = XML.GetTagText("BUILDINGORHOUSENAME");
			frmScreen.txtCurrentAddressHouseNumber.value = XML.GetTagText("BUILDINGORHOUSENUMBER");
			frmScreen.txtCurrentAddressStreet.value = XML.GetTagText("STREET");
			frmScreen.txtCurrentAddressDistrict.value = XML.GetTagText("DISTRICT");
			frmScreen.txtCurrentAddressTown.value = XML.GetTagText("TOWN");
			frmScreen.txtCurrentAddressCounty.value = XML.GetTagText("COUNTY");
			frmScreen.cboCurrentAddressCountry.value = XML.GetTagText("COUNTRY");
			m_bCurrentAddressPAFIndicator = (XML.GetTagText("PAFINDICATOR") == "1");
		}

		if(XML.GetTagText("ADDRESSTYPE") == "2")
		{
			/* CORE UPGRADE 702 m_sCorrespondenceCustAddressSeqNo = XML.GetTagText("CUSTOMERADDRESSSEQUENCENUMBER");
			m_sCorrespondenceAddressGUID = XML.GetTagText("ADDRESSGUID"); */
			//SG 29/05/02 SYS4767 START
			//SG 01/03/02 SYS4186
			//Placed if statement round existing code
			if (m_sMetaAction != "CreateNewCustomerForExistingApplication")
			{
				m_sCorrespondenceCustAddressSeqNo = XML.GetTagText("CUSTOMERADDRESSSEQUENCENUMBER");
				m_sCorrespondenceAddressGUID = XML.GetTagText("ADDRESSGUID");
			}
			//SG 01/03/02 SYS4186
			//SG 29/05/02 SYS4767 END
			
			frmScreen.txtCorrespondenceAddressPostcode.value = XML.GetTagText("POSTCODE");
			frmScreen.txtCorrespondenceAddressFlatNumber.value = XML.GetTagText("FLATNUMBER");
			frmScreen.txtCorrespondenceAddressHouseName.value = XML.GetTagText("BUILDINGORHOUSENAME");
			frmScreen.txtCorrespondenceAddressHouseNumber.value = XML.GetTagText("BUILDINGORHOUSENUMBER");
			frmScreen.txtCorrespondenceAddressStreet.value = XML.GetTagText("STREET");
			frmScreen.txtCorrespondenceAddressDistrict.value = XML.GetTagText("DISTRICT");
			frmScreen.txtCorrespondenceAddressTown.value = XML.GetTagText("TOWN");
			frmScreen.txtCorrespondenceAddressCounty.value = XML.GetTagText("COUNTY");
			frmScreen.cboCorrespondenceAddressCountry.value = XML.GetTagText("COUNTRY");
			frmScreen.chkCorrespondenceAddressBFPO.checked = (XML.GetTagText("BFPO") == "1");
			m_bCorrespondenceAddressPAFIndicator = (XML.GetTagText("PAFINDICATOR") == "1");
			
			
			EnableDisableCorrespondenceAddressBFPO(); // EP8 pct
		}
	}
}

function GetTelephoneDetails(XML)
{
<%	// Retrieves the details from the <CUSTOMERTELEPHONENUMBERLIST> section of the returned XML
	// As there is only 1 <CUSTOMERTELEPHONENUMBERLIST> and it only contains <CUSTOMERTELEPHONENUMBER> tags,
	// we can access <CUSTOMERTELEPHONENUMBER> directly
	// Loops through the <CUSTOMERTELEPHONENUMBER> tags
	
	// Arrays of the fields to be populated
%>	var sType	= new Array("cboType1"      , "cboType2"      , "cboType3"      , "cboType4");
	var	sCountryCode = new Array("txtCountryCode1"     , "txtCountryCode2"     , "txtCountryCode3"     , "txtCountryCode4");
	var	sAreaCode = new Array("txtAreaCode1"     , "txtAreaCode2"     , "txtAreaCode3"     , "txtAreaCode4");
	var	sNumber	= new Array("txtTelNumber1"     , "txtTelNumber2"     , "txtTelNumber3"     , "txtTelNumber4");
	var	sExtension	= new Array("txtExtensionNumber1"     , "txtExtensionNumber2"     , "txtExtensionNumber3"     , "txtExtensionNumber4");
	var sTime	= new Array("txtTime1", "txtTime2", "txtTime3", "txtTime4");
	var sOption	= new Array("optPREFERREDCONTACT1","optPREFERREDCONTACT2", "optPREFERREDCONTACT3", "optPREFERREDCONTACT4");

	XML.ActiveTag = null;
	XML.CreateTagList("CUSTOMERTELEPHONENUMBER");

	for(var nLoop = 0;nLoop < XML.ActiveTagList.length && nLoop < 4;nLoop++)
	{
		/* CORE UPGRADE 702 XML.SelectTagListItem(nLoop);
		frmScreen(sType[nLoop]).value = XML.GetTagText("USAGE");
		frmScreen(sCountryCode[nLoop]).value = XML.GetTagText("COUNTRYCODE");
		frmScreen(sAreaCode[nLoop]).value = XML.GetTagText("AREACODE");
		frmScreen(sNumber[nLoop]).value = XML.GetTagText("TELEPHONENUMBER");
		frmScreen(sExtension[nLoop]).value = XML.GetTagText("EXTENSIONNUMBER");
		frmScreen(sTime[nLoop]).value = XML.GetTagText("CONTACTTIME");
		frmScreen(sOption[nLoop]).checked = (XML.GetTagText("PREFERREDMETHODOFCONTACT") == "1")? true:false; */
		XML.SelectTagListItem(nLoop);	
		
		//SG 29/05/02 SYS4767 START				
		//SG 28/02/02 SYS4186 Placed if statement around existing population code
		
		<% /* ASu BMIDS00690 - Start. Change to always get first applicants 'Home' telephone number only where 'Copy'function used. */ %>
		if ((m_sMetaAction == "CreateNewCustomerForExistingApplication" || "UpdateExistingCustomer") 
			&& XML.GetTagAttribute("USAGE","TEXT") != "Home" && m_sCopy == true)
			<% /* ASu - End */ %>	
		{
			//Do not populate		
		}
		else
		{				
			frmScreen(sType[nLoop]).value = XML.GetTagText("USAGE");
			frmScreen(sCountryCode[nLoop]).value = XML.GetTagText("COUNTRYCODE");
			frmScreen(sAreaCode[nLoop]).value = XML.GetTagText("AREACODE");
			frmScreen(sNumber[nLoop]).value = XML.GetTagText("TELEPHONENUMBER");	
			frmScreen(sExtension[nLoop]).value = XML.GetTagText("EXTENSIONNUMBER");
			frmScreen(sTime[nLoop]).value = XML.GetTagText("CONTACTTIME");
			frmScreen(sOption[nLoop]).checked = (XML.GetTagText("PREFERREDMETHODOFCONTACT") == "1")? true:false;
		}
		//SG 29/05/02 SYS4767 END
	}
}

function CommitScreen()
{
<%  /*XMLstructure:<REQUESTUSERID=?,USERTYPE=?,UNIT=?,ACTION="CREATE" or "UPDATE">
			Customer Details
			Address Details
			Areas Of Interest Details
		<REQUEST>

		APS 10/09/99 - UNIT TEST REF 102 Setting CustomerLockIndicator = true
	*/
%>	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	/* CORE UPGRADE 702 var XML = new scXMLFunctions.XMLObject(); */

	var TagREQUEST = XML.CreateRequestTag(window,null);
	var TagRequestType = null;
	var sASPFile;
	if(m_sMetaAction == "CreateNewCustomerForNewApplication" || m_sMetaAction == "CreateNewCustomerForExistingApplication")
	{
		<% /* PSC 11/10/2005 MAR57 */ %>
		XML.SetAttribute("MessageType", "CreateCustomer");
		TagRequestType = XML.CreateActiveTag("CREATE");
		scScreenFunctions.SetContextParameter(window,"idCustomerLockIndicator","1");
		sASPFile = "CreateCustomerDetails.asp";
	}
	//SD MAR258 Start
	//NO critical data check needed, if it is a new application
	//if (m_sMetaAction == "CreateExistingCustomerForNewApplication" || m_sMetaAction == "CreateExistingCustomerForExistingApplication" || m_sMetaAction == "UpdateExistingCustomer" ||	m_sContext == "CompletenessCheck")
	if (m_sMetaAction == "CreateExistingCustomerForNewApplication")
	{
		TagRequestType = XML.CreateActiveTag("UPDATE");
		sASPFile = "UpdateCustomerDetails.asp";
	}
	
	//SD MAR590 cause of error
	/*
	if(m_sMetaAction == "CreateNewCustomerForExistingApplication")
	{		
		XML.SetAttribute("OPERATION","CriticalDataCheck");
		TagRequestType = XML.CreateActiveTag("CREATE");
		sASPFile = "CreateCustomerDetails";  //method name for critical data context
	}*/
	
	if(m_sMetaAction == "CreateExistingCustomerForExistingApplication" || m_sMetaAction == "UpdateExistingCustomer" || m_sContext == "CompletenessCheck")
	{
		XML.SetAttribute("OPERATION","CriticalDataCheck");
		TagRequestType = XML.CreateActiveTag("UPDATE");
		sASPFile = "UpdateCustomerDetails";  //method name for critical data context
	}
	if(TagRequestType != null)
	{
		var TagCUSTOMERVERSION	= WriteCustomerDetails(XML,TagRequestType);
		WriteAddressDetails(XML,TagCUSTOMERVERSION);
		WriteTelephoneDetails(XML,TagCUSTOMERVERSION);

		if (m_sMetaAction == "CreateExistingCustomerForExistingApplication" || m_sMetaAction == "UpdateExistingCustomer" || m_sContext == "CompletenessCheck")
		{
			//SD 09/11/05	
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
			XML.SetAttribute("METHOD",sASPFile);	
			
			window.status = "Critical Data Check - please wait";
			switch (ScreenRules()) 
			{
				case 1: // Warning
				case 0: // OK
						XML.RunASP(document,"OmigaTMBO.asp");
						window.status = "";
						break;
				default: // Error
					XML.SetErrorResponse();
			}
			
			window.status = "";
		
		}
		else
		{
		//new application
			switch (ScreenRules()) 
				{
				case 1: // Warning
				case 0: // OK
						XML.RunASP(document,sASPFile);
					break;
				default: // Error
					XML.SetErrorResponse();
				}
		}		

		//SD MAR258 End
		
		if(XML.IsResponseOK())
		{
			XML.CreateTagList("CUSTOMERKEY");
			if(XML.SelectTagListItem(0))
			{
				m_sCustomerNumber = XML.GetTagText("CUSTOMERNUMBER");
				scScreenFunctions.SetContextParameter(window,"idCustomerNumber",m_sCustomerNumber);
				m_sCustomerVersionNumber = XML.GetTagText("CUSTOMERVERSIONNUMBER");
				scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber",m_sCustomerVersionNumber);
				<% /* PSC 11/10/2005 MAR57 - Start */ %>			
				m_sOtherSystemCustomerNumber = XML.GetTagText("OTHERSYSTEMCUSTOMERNUMBER");
				scScreenFunctions.SetContextParameter(window,"idOtherSystemCustomerNumber",m_sOtherSystemCustomerNumber);
				<% /* PSC 11/10/2005 MAR57 - End */ %>			

			}

			return true;
		}
	}

	return false;
}

function CommitToPackage()
{
<%	/*	XML structure
		<REQUEST>
			<CUSTOMERPACKAGE>
				fields
			</CUSTOMERPACKAGE>
		</REQUEST>

		AY/SR 10/03/00 - Processing to calculate the next customerorder number
	*/
%>	if(m_sMetaAction == "CreateNewCustomerForExistingApplication" || m_sMetaAction == "CreateExistingCustomerForExistingApplication")
	{
		var nCustomerOrder = 0;
		var nApplicantCount = 0 ;
		// PJO BM0006 for(var nLoop = 1; nLoop <= 5; nLoop++)
		for(var nLoop = 1; nLoop <= m_iMaxCustomers; nLoop++)
			if(scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop) == "1") 
			{
			    nCustomerOrder++;
			    nApplicantCount++;
			}

		nCustomerOrder++ ;

		/* CORE UPGRADE 702 var XML = new scXMLFunctions.XMLObject(); */
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		//XML.CreateRequestTag(window,"CREATE");
		var TagRequestType = XML.CreateRequestTag(window, null);
		XML.SetAttribute("OPERATION","CriticalDataCheck");
		XML.CreateActiveTag("CUSTOMERPACKAGE");
		XML.CreateTag("CUSTOMERNUMBER",m_sCustomerNumber);
		XML.CreateTag("CUSTOMERVERSIONNUMBER",m_sCustomerVersionNumber);
		// PJO BM0006
		if (nApplicantCount >= m_iMaxApplicants)
		{
		    XML.CreateTag("CUSTOMERROLETYPE","2");
		}
		else
		{
     		XML.CreateTag("CUSTOMERROLETYPE","1");
        }
		XML.CreateTag("CUSTOMERORDER",nCustomerOrder);
		XML.CreateTag("PACKAGENUMBER",m_sPackageNumber);
		XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		XML.CreateTag("TYPEOFAPPLICATION",m_sTypeOfApplicationValue);

		<% /* EP2_56 GHun */ %>
		if (m_bNewToEIndicator == 1)
		{		
			XML.CreateTag("CIFNUMBER", m_sOtherSystemCustomerNumber);
			XML.CreateTag("TITLE", frmScreen.cboTitle.value);
			XML.CreateTag("FIRSTFORENAME", frmScreen.txtFirstForename.value);
			XML.CreateTag("SECONDFORENAME", frmScreen.txtSecondForename.value);
			XML.CreateTag("OTHERFORENAMES", frmScreen.txtOtherForenames.value);
			XML.CreateTag("SURNAME", frmScreen.txtSurname.value);
		}
		<% /* EP2_56 End */ %>

		//SD 09/11/05	
		XML.SelectTag(null,"REQUEST");
		XML.CreateActiveTag("CRITICALDATACONTEXT");
		XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		XML.SetAttribute("SOURCEAPPLICATION","Omiga");
		XML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
		XML.SetAttribute("ACTIVITYINSTANCE","1");
		XML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
		XML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
		XML.SetAttribute("COMPONENT","omApp.ApplicationManagerBO");
		XML.SetAttribute("METHOD","AddCustomerToApplication");	
			
		window.status = "Critical Data Check - please wait";
			
		switch (ScreenRules()) 
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "OmigaTMBO.asp");	
					window.status = ""
				break;
			default: // Error
				XML.SetErrorResponse();
			}
		window.status = "";

		if(XML.IsResponseOK()) return true;
	}
	else return true;

	return false;
}

function WriteCustomerDetails(XML,TagParent)
{
<%	/*	XML structure:
		<CUSTOMER>
			fields
			<CUSTOMERVERSION>
				fields
			</CUSTOMERVERSION>
		</CUSTOMER>

		APS UNIT TEST REF 13 - Not storing Mothers MaidenName or email address
	*/
%>	XML.ActiveTag = TagParent;
	XML.CreateActiveTag("CUSTOMER");
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	
	<% /* PSC 13/10/2005 MAR57 - Start */ %>
	if(m_sMetaAction == "CreateNewCustomerForNewApplication" || m_sMetaAction == "CreateNewCustomerForExistingApplication")
	{
		XML.CreateTag("UPDATECRSCUSTOMER", "1");
	}
	<% /* PSC 13/10/2005 MAR57 - End */ %>

	var TagCUSTOMERVERSION = XML.CreateActiveTag("CUSTOMERVERSION");
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	XML.CreateTag("DATEOFBIRTH", frmScreen.txtDateOfBirth.value);
	XML.CreateTag("FIRSTFORENAME", frmScreen.txtFirstForename.value);
	XML.CreateTag("GENDER", frmScreen.cboGender.value);
	XML.CreateTag("MAILSHOTREQUIRED", scScreenFunctions.GetRadioGroupValue(frmScreen,"MailshotInd"));
	XML.CreateTag("OTHERFORENAMES", frmScreen.txtOtherForenames.value);
	XML.CreateTag("SECONDFORENAME", frmScreen.txtSecondForename.value);
	XML.CreateTag("SURNAME", frmScreen.txtSurname.value);
	XML.CreateTag("TITLE", frmScreen.cboTitle.value);
	XML.CreateTag("TITLEOTHER", frmScreen.txtOtherTitle.value);
	XML.CreateTag("SPECIALNEEDS", frmScreen.cboSpecialNeeds.value);
	XML.CreateTag("CONTACTEMAILADDRESS", frmScreen.txtContactEMailAddress.value);
	XML.CreateTag("EMAILPREFERRED", frmScreen.optPREFERREDCONTACT5.checked ? "1":"0");

	<% /* EP903 */ %>
	XML.CreateTag("APPLICATIONDATE", m_sApplicationDate);

	<% /* BMIDS836 If a customer is being added to an existing application, set the NewToECustomerInd */ %>
	if(m_sMetaAction == "CreateNewCustomerForExistingApplication" || m_sMetaAction == "CreateExistingCustomerForExistingApplication")
	{
		XML.CreateTag("NEWTOECUSTOMERIND", m_bNewToEIndicator);
	}
	
	return TagCUSTOMERVERSION;
}

function WriteAddressDetails(XML,TagParent)
{
<%	/*	XML structure:
		<CUSTOMERADDRESSLIST>
			<CUSTOMERADDRESS> (for current address)
				fields
				<ADDRESS>
					fields
				</ADDRESS>
			</CUSTOMERADDRESS>
			<CUSTOMERADDRESS> (for correspondence address)
				fields
				<ADDRESS>
					fields
				</ADDRESS>
			</CUSTOMERADDRESS>
		</CUSTOMERADDRESSLIST>
	*/
%>	XML.ActiveTag = TagParent;
	var TagCUSTOMERADDRESSLIST = XML.CreateActiveTag("CUSTOMERADDRESSLIST");
	XML.CreateActiveTag("CUSTOMERADDRESS");
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	if(m_sCurrentCustAddressSeqNo != null) XML.CreateTag("CUSTOMERADDRESSSEQUENCENUMBER",m_sCurrentCustAddressSeqNo);
	XML.CreateTag("ADDRESSTYPE", "1");

	XML.CreateActiveTag("ADDRESS");
	if(m_sCurrentAddressGUID != null) XML.CreateTag("ADDRESSGUID",m_sCurrentAddressGUID);
	XML.CreateTag("BUILDINGORHOUSENAME", frmScreen.txtCurrentAddressHouseName.value);
	XML.CreateTag("BUILDINGORHOUSENUMBER", frmScreen.txtCurrentAddressHouseNumber.value);
	XML.CreateTag("FLATNUMBER", frmScreen.txtCurrentAddressFlatNumber.value);
	XML.CreateTag("STREET", frmScreen.txtCurrentAddressStreet.value);
	XML.CreateTag("DISTRICT", frmScreen.txtCurrentAddressDistrict.value);
	XML.CreateTag("TOWN", frmScreen.txtCurrentAddressTown.value);
	XML.CreateTag("POSTCODE", frmScreen.txtCurrentAddressPostcode.value);
	XML.CreateTag("COUNTY", frmScreen.txtCurrentAddressCounty.value);
	XML.CreateTag("COUNTRY", frmScreen.cboCurrentAddressCountry.value);
	XML.CreateTag("PAFINDICATOR", m_bCurrentAddressPAFIndicator ? "1" : "0");

	XML.ActiveTag = TagCUSTOMERADDRESSLIST;
	XML.CreateActiveTag("CUSTOMERADDRESS");
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	if(m_sCorrespondenceCustAddressSeqNo != null) XML.CreateTag("CUSTOMERADDRESSSEQUENCENUMBER",m_sCorrespondenceCustAddressSeqNo);
	XML.CreateTag("ADDRESSTYPE", "2");

	XML.CreateActiveTag("ADDRESS");
	if(m_sCorrespondenceAddressGUID != null) XML.CreateTag("ADDRESSGUID",m_sCorrespondenceAddressGUID);
	XML.CreateTag("BUILDINGORHOUSENAME", frmScreen.txtCorrespondenceAddressHouseName.value);
	XML.CreateTag("BUILDINGORHOUSENUMBER", frmScreen.txtCorrespondenceAddressHouseNumber.value);
	XML.CreateTag("FLATNUMBER", frmScreen.txtCorrespondenceAddressFlatNumber.value);
	XML.CreateTag("STREET", frmScreen.txtCorrespondenceAddressStreet.value);
	XML.CreateTag("DISTRICT", frmScreen.txtCorrespondenceAddressDistrict.value);
	XML.CreateTag("TOWN", frmScreen.txtCorrespondenceAddressTown.value);
	XML.CreateTag("POSTCODE", frmScreen.txtCorrespondenceAddressPostcode.value);
	XML.CreateTag("COUNTY", frmScreen.txtCorrespondenceAddressCounty.value);
	XML.CreateTag("COUNTRY", frmScreen.cboCorrespondenceAddressCountry.value);
	XML.CreateTag("BFPO", frmScreen.chkCorrespondenceAddressBFPO.checked ? "1":"0");
	XML.CreateTag("PAFINDICATOR", m_bCorrespondenceAddressPAFIndicator ? "1" : "0");
}

function WriteTelephoneDetails(XML,TagParent)
{
<%	/*	XML structure:
		<CUSTOMERTELEPHONENUMBERLIST>
			<CUSTOMERTELEPHONENUMBER>
				fields
			</CUSTOMERTELEPHONENUMBER>
			<CUSTOMERTELEPHONENUMBER>
				fields
			</CUSTOMERTELEPHONENUMBER>
			<CUSTOMERTELEPHONENUMBER>
				fields
			</CUSTOMERTELEPHONENUMBER>
			<CUSTOMERTELEPHONENUMBER>
				fields
			</CUSTOMERTELEPHONENUMBER>
		</CUSTOMERTELEPHONENUMBERLIST>
	*/
%>	var sType	= new Array("cboType1"      , "cboType2"      , "cboType3"      , "cboType4");
	var	sCountryCode = new Array("txtCountryCode1"     , "txtCountryCode2"     , "txtCountryCode3"     , "txtCountryCode4");
	var	sAreaCode = new Array("txtAreaCode1"     , "txtAreaCode2"     , "txtAreaCode3"     , "txtAreaCode4");
	var	sNumber	= new Array("txtTelNumber1"     , "txtTelNumber2"     , "txtTelNumber3"     , "txtTelNumber4");
	var	sExtension	= new Array("txtExtensionNumber1"     , "txtExtensionNumber2"     , "txtExtensionNumber3"     , "txtExtensionNumber4");
	var sTime	= new Array("txtTime1", "txtTime2", "txtTime3", "txtTime4");
	var sOption	= new Array("optPREFERREDCONTACT1","optPREFERREDCONTACT2", "optPREFERREDCONTACT3", "optPREFERREDCONTACT4");

	XML.ActiveTag = TagParent;

	var TagCUSTOMERTELEPHONENUMBERLIST = XML.CreateActiveTag("CUSTOMERTELEPHONENUMBERLIST");
	for(var nLoop = 0;nLoop < 4;nLoop++)
	{
		XML.ActiveTag = TagCUSTOMERTELEPHONENUMBERLIST;
		XML.CreateActiveTag("CUSTOMERTELEPHONENUMBER");
		XML.CreateTag("CUSTOMERNUMBER"       ,m_sCustomerNumber);
		XML.CreateTag("CUSTOMERVERSIONNUMBER",m_sCustomerVersionNumber);
		XML.CreateTag("COUNTRYCODE"      ,frmScreen(sCountryCode[nLoop]).value);
		XML.CreateTag("AREACODE"      ,frmScreen(sAreaCode[nLoop]).value);
		XML.CreateTag("TELEPHONENUMBER"      ,frmScreen(sNumber[nLoop]).value);
		XML.CreateTag("EXTENSIONNUMBER"      ,frmScreen(sExtension[nLoop]).value);
		XML.CreateTag("USAGE"                ,frmScreen(sType[nLoop]).value);
		XML.CreateTag("CONTACTTIME"          ,frmScreen(sTime[nLoop]).value);
		XML.CreateTag("PREFERREDMETHODOFCONTACT"	 ,(frmScreen(sOption[nLoop]).checked)? "1":"0");
	}
}

<%
/* ASu Removed As per BMIDS00139 Start
function WriteAreasOfInterest(XML,TagParent)
{
	XML.ActiveTag = TagParent;

	if(m_sXMLAreasOfInterest == null) XML.CreateActiveTag("AREASOFINTERESTLIST");
	else
	{	
		var AreasXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		// CORE UPGRADE 702 var AreasXML = new scXMLFunctions.XMLObject(); //
		AreasXML.LoadXML(m_sXMLAreasOfInterest);
		XML.AddXMLBlock(AreasXML.XMLDocument);
		AreasXML = null;
	}
}
//ASu End */
%>
<%
	/* SR - 15/03/00 : SYS0192 - Validate the telephone details and, set the focus to appropriate
					   field if the validation fails.
	*/
%>

<% /* SYS1053 SA 24/5/01 New function to check 1st character of names */ %>
function CheckNames(sNameToCheck)
{
	var CheckStr = sNameToCheck.substr(0,1)
	if (_validation.isNum(CheckStr))
	{
		return false;
	}
	
	return true;	
}

function ValidateTelephoneDetails()
{
	var sType	= new Array("cboType1", "cboType2", "cboType3" , "cboType4");
	var	sNumber	= new Array("txtTelNumber1", "txtTelNumber2", "txtTelNumber3", "txtTelNumber4");
	var sTime	= new Array("txtTime1", "txtTime2", "txtTime3", "txtTime4");
	
	for(var nLoop = 0;nLoop < 4;nLoop++)
	{
		if (frmScreen(sType[nLoop]).value !== '')
		{
			if (frmScreen(sNumber[nLoop]).value === '')
			{
				alert('Please enter a Telephone Number for each Contact')
				frmScreen.all(sType[nLoop]).focus()
				return false ;	
			}
		}
		else
		{
			if (frmScreen(sNumber[nLoop]).value !== '' || frmScreen(sTime[nLoop]).value !== '')
			{
				alert('For each telephone number please enter the Type and Telephone Number')
				frmScreen.all(sType[nLoop]).focus()
				return false ;	
			}
		}
	}
	return true;
}
function ClearCurrentAddressPAFIndicator()
{
<%	// The XML needs to store 1 or 0 as opposed to true or false
%>	m_bCurrentAddressPAFIndicator = false;
}

function ClearCorrespondenceAddressPAFIndicator()
{
<%	// The XML needs to store 1 or 0 as opposed to true or false
%>	m_bCorrespondenceAddressPAFIndicator = false;
}

function RetrieveContextData()
{
<%	// APS UNIT TEST REF 8 - If we have an existing customer then we want to
 	// read the context else it will be null from initialisation
%>	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","CreateNewCustomerForNewApplication");
	m_sContext = scScreenFunctions.GetContextParameter(window,"idProcessContext",null);
	
	//SG 29/05/02 SYS4767 START
	//SG 20/03/02 MSMS0013
	m_sCustNo1 = scScreenFunctions.GetContextParameter(window,"idCustomerNumber1",null);
	m_sCustVersNo1 = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber1",null);
	m_sCustRoleType1 = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType1",null);
	//SG 20/03/02 MSMS0013
	//SG 29/05/02 SYS4767 END
	
	<% /* MF Modification of MAR19 */ %>
	if(m_sMetaAction == "CreateNewCustomerForNewApplication" || m_sMetaAction == "CreateNewCustomerForExistingApplication"){
		m_bNewCustomer=true;		
	}
	
	<% /* BMIDS00393 MDC 03/09/2002 - Still try to get context values when creating a new customer*/ %>
 	if(m_sMetaAction == "CreateExistingCustomerForNewApplication" || 
 		m_sMetaAction == "CreateExistingCustomerForExistingApplication" ||
 		m_sContext == "CompletenessCheck" ||
		m_sMetaAction == "UpdateExistingCustomer" || 
		m_sMetaAction == "CreateNewCustomerForNewApplication")
	{
		m_sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber","");
		m_sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber","");
	}

	<%//GD Added SYS1752 20.02.2001
	%>
	m_sOtherSystemCustomerNumber = scScreenFunctions.GetContextParameter(window,"idOtherSystemCustomerNumber",null);


	m_sPackageNumber = scScreenFunctions.GetContextParameter(window,"idPackageNumber",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
	m_sTypeOfApplicationValue = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sCustomerLockIndicator = scScreenFunctions.GetContextParameter(window,"idCustomerLockIndicator","0");
	
	<% /* BMIDS758  Set the Transfer Of Equity flag depending on Mortgage Type */ %>
	var ValidationList = new Array(1);
	ValidationList[0] = "CC";   // Transfer Of Equity    BMIDS836
	
	if ( m_sTypeOfApplicationValue != "" )
	{
		var combXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		if(combXML.IsInComboValidationList(document,"TypeOfMortgage", m_sTypeOfApplicationValue , ValidationList)) 
			m_bNewToEIndicator = 1;
	}	
	
	<%/* EP903 Read Application Date */%> 
	m_sApplicationDate = scScreenFunctions.GetContextParameter(window,"idApplicationDate","");  
				
}

function frmScreen.cboType1.onchange()
{
scScreenFunctions.SetFieldState(frmScreen, "txtExtensionNumber1", ((frmScreen.cboType1.value) == "2") ? "W":"D");
scScreenFunctions.SetFieldState(frmScreen, "optPREFERREDCONTACT1", ((frmScreen.cboType1.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtTelNumber1", ((frmScreen.cboType1.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtTime1", ((frmScreen.cboType1.value) == "") ? "D":"W");
//JR - Omiplus24
scScreenFunctions.SetFieldState(frmScreen, "txtCountryCode1", ((frmScreen.cboType1.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtAreaCode1", ((frmScreen.cboType1.value) == "") ? "D":"W");

}
function frmScreen.cboType2.onchange()
{
scScreenFunctions.SetFieldState(frmScreen, "txtExtensionNumber2", ((frmScreen.cboType2.value) == "2") ? "W":"D");
scScreenFunctions.SetFieldState(frmScreen, "optPREFERREDCONTACT2", ((frmScreen.cboType2.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtTelNumber2", ((frmScreen.cboType2.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtTime2", ((frmScreen.cboType2.value) == "") ? "D":"W");
//JR - Omiplus24
scScreenFunctions.SetFieldState(frmScreen, "txtCountryCode2", ((frmScreen.cboType2.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtAreaCode2", ((frmScreen.cboType2.value) == "") ? "D":"W");
}
function frmScreen.cboType3.onchange()
{
scScreenFunctions.SetFieldState(frmScreen, "txtExtensionNumber3", ((frmScreen.cboType3.value) == "2") ? "W":"D");
scScreenFunctions.SetFieldState(frmScreen, "optPREFERREDCONTACT3", ((frmScreen.cboType3.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtTelNumber3", ((frmScreen.cboType3.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtTime3", ((frmScreen.cboType3.value) == "") ? "D":"W");
//JR - Omiplus24
scScreenFunctions.SetFieldState(frmScreen, "txtCountryCode3", ((frmScreen.cboType3.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtAreaCode3", ((frmScreen.cboType3.value) == "") ? "D":"W");
}
function frmScreen.cboType4.onchange()
{
scScreenFunctions.SetFieldState(frmScreen, "txtExtensionNumber4", ((frmScreen.cboType4.value) == "2") ? "W":"D");
scScreenFunctions.SetFieldState(frmScreen, "optPREFERREDCONTACT4", ((frmScreen.cboType4.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtTelNumber4", ((frmScreen.cboType4.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtTime4", ((frmScreen.cboType4.value) == "") ? "D":"W");
//JR - Omiplus24
scScreenFunctions.SetFieldState(frmScreen, "txtCountryCode4", ((frmScreen.cboType4.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtAreaCode4", ((frmScreen.cboType4.value) == "") ? "D":"W");
}

function frmScreen.txtContactEMailAddress.onchange()
{
	scScreenFunctions.SetFieldState(frmScreen, "optPREFERREDCONTACT5", ((frmScreen.txtContactEMailAddress.value) == "") ? "D":"W");
}

// SYS3964 Changed Age calc to Jscript instead of VB as it is more robust
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
									
			if(nAppAge < sMinAge ) alert("Age is below accepted age");
			else if (nAppAge > sMaxAge) alert("Age is greater than accepted age");
			else if (nAge != -1)
				frmScreen.txtAge.value = nAge.toString();
		}
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
		nAge = top.frames[1].document.all.scMathFunctions.GetYearsBetweenDates(dteBirthdate, dteToday);
	}

	return(nAge);
}

<% /* EP903 New function */ %>
function CalculateCustomerAgeOnApplication()
{
	var nAppAge = -1;
	var dtApplicationDate;
	
	var dteBirthdate = scScreenFunctions.GetDateObject(frmScreen.txtDateOfBirth);
	if(dteBirthdate != null)
	{
		<% /* If there is no application date, use today's date */ %>
		if (m_sApplicationDate == "")
			dtApplicationDate = scScreenFunctions.GetAppServerDate(true);
		else
			dtApplicationDate = scScreenFunctions.GetDateObjectFromString(m_sApplicationDate);
		
		nAppAge = top.frames[1].document.all.scMathFunctions.GetYearsBetweenDates(dteBirthdate, dtApplicationDate);
	}

	return(nAppAge);
}


//SG 29/05/02 SYS4767 Function added
function frmScreen.btnCopy.onclick()	//SG 28/02/02 SYS4186
{
	CopyAddress();	
}

<% /* PJO BM0006 - Function added copied from CR030 */ %>
function GetMaximumApplicantsData()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	m_iMaxCustomers = XML.GetGlobalParameterAmount(document, "MaximumCustomers") ;
	m_iMaxApplicants = XML.GetGlobalParameterAmount(document, "MaximumApplicants") ;
	
	XML = null ;
}



//SG 29/05/02 SYS4767 Function added
function CopyAddress()	//SG 28/02/02 SYS4186
{
	//Copies the address from the first customer and populates details on screen
	//so they can be saved for the next customer being added.

	//Get the customer details
	<%/* SG 05/06/02 SYS4804 */%>	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	<%/* var XML = new scXMLFunctions.XMLObject(); */%>
	<%/* SG 05/06/02 SYS4804 */%>	

	XML.CreateRequestTag(window,"SEARCH");
	XML.CreateActiveTag("SEARCH");
	XML.CreateActiveTag("CUSTOMER");

	<% /* MV - 05/07/2002 - BMIDS00096 - Commented the Previos code and added the following lines to find
		customernumber and customernumber 
	//XML.CreateTag("CUSTOMERNUMBER",m_sCustNo1);
	//XML.CreateTag("CUSTOMERVERSIONNUMBER",m_sCustVersNo1); */ %>
	<% /* ASu BMIDS00690 - Start. Change to always get first applicants address details */ %>
	//XML.CreateTag("CUSTOMERNUMBER",scScreenFunctions.GetContextParameter(window,"idCustomerNumber",m_sCustomerNumber));
	//XML.CreateTag("CUSTOMERVERSIONNUMBER",scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber",m_sCustomerVersionNumber));
	XML.CreateTag("CUSTOMERNUMBER",scScreenFunctions.GetContextParameter(window,"idCustomerNumber1",m_sCustNo1));
	XML.CreateTag("CUSTOMERVERSIONNUMBER",scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber1",m_sCustVersNo1));
	<% /* ASu - End */ %>	
	// 	XML.RunASP(document,"GetCustomerDetails.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	
	<% /* MAR160 case is blocked because no need to apply ScreenRules() before copy address
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"GetCustomerDetails.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}  
	*/ %>
	XML.RunASP(document,"GetCustomerDetails.asp");

	//Populate the screen...
	if(XML.IsResponseOK())
	{
		GetAddressDetails(XML);
		m_sCopy = true;
		GetTelephoneDetails(XML);
		Initialise();
	}
	
	XML = null;	
	<% /* ASu BMIDS00690 - Start. Reset trigger */ %>
	m_sCopy = false;	
	<% /* ASu - End. */ %>
}
<% /* EP8 pct 13/03/2006 */ %>
function EnableDisableCorrespondenceAddressBFPO()
{
	if(frmScreen.chkCorrespondenceAddressBFPO.checked)
	{
		frmScreen.txtCorrespondenceAddressPostcode.value = "";
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtCorrespondenceAddressPostcode");
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "cboCorrespondenceAddressCountry");
		frmScreen.btnCorrespondenceAddressPAFSearch.disabled=true;
		frmScreen.txtCorrespondenceAddressPostcode.removeAttribute("required");
		frmScreen.txtCorrespondenceAddressPostcode.parentElement.parentElement.style.color = "";
		frmScreen.txtCorrespondenceAddressFlatNumber.focus();
	}
	else
	{
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtCorrespondenceAddressPostcode");
		scScreenFunctions.SetFieldToWritable(frmScreen, "cboCorrespondenceAddressCountry");
		frmScreen.btnCorrespondenceAddressPAFSearch.disabled=false;
		frmScreen.txtCorrespondenceAddressPostcode.setAttribute("required", "true");
		frmScreen.txtCorrespondenceAddressPostcode.parentElement.parentElement.style.color = "Red";
		frmScreen.txtCorrespondenceAddressPostcode.focus();
	}
}
<% /* EP8 pct 13/03/2006 END*/ %>
-->
		</script>
		<% /* CORE UPGRADE 702. BMIDS957  Unused VBScript code removed. */ %>
	</body>
</html>
