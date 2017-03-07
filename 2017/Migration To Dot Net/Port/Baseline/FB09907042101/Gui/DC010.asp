<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
	<% /* 
Workfile:      DC010.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Application Source Details and Intermediary details
			   Screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		03/11/99	Created.
JLD		07/12/99	bug fixes : DC/008 - made fields wider to fill screen width
						DC/010 - When change to direct, remove some field entries
							   - When group connection changed to no, check for existing group connections
						DC/012 - In readonly, disable search button.
JLD		10/12/1999	DC/017 - ApplicationPackageIndicator now set correctly
					DC/018 - spelling corrected in messages
					DC/019 - IntroAgentRef max length allowed corrected
					DC/020 - Readonly mode should now still route.
					DC/021 - Address format checked.
JLD		16/12/1999	SYS0075 - groupconnectionindicator set correctly when moving between screens
JLD		16/12/1999	DC/019 - Check length of additional details. Error if > 500.
AD		28/01/2000	Rework
AY		14/02/00	Change to msgButtons button types
IW		23/03/00	SYS0127 - Made Intermediary name Mandatory etc.
AD		23/03/2000	SYS0127 - Intermediary now being saved properly.
AY		29/03/00	New top menu/scScreenFunctions change
MH		15/05/00    SYS0738 Update on first time required.
IW		16/05/00	SYS0743 Move OK & Cancel up a bit (to avoid scroll bars)
BG		17/05/00	SYS0752 Removed Tooltips
CL		05/03/01	SYS1920 Read only functionality added
SA		04/06/01	SYS0855 Extra validation added to check marketing source.
DJP		06/08/01	SYS2734 Client cosmetic customisation SYS2564
JR		10/10/01	Omiplus24 Modified PopulateIntermediaryDetails to inlcude Country/Area Code for telephone number change
STB		11/04/02	SYS3644 Altered validation code to use combo validation types rather than valueIDs.
DPF		21/06/02	BMIDS00077 Amended code to bring in line with Core V7.0.2, changes are...
					SYS4727 Use cached versions of frame functions
					SYS4767 MSMS to Core integration
					SYS0081 intermediary status is mandatory.
TW      09/10/02    SYS5115 Modified to incorporate client validation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
MV		28/05/2002	BMIDS00013 - BM044 - amended "Has the BM Declaration Received Yes or No"
MV		18/07/2002	BMIDS00179 - Core Upgrade Rollback Ref SYS4786
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
ASu		02/09/2002	BMIDS00037 - Addition of Reset data freeze indicator.
MO		20/09/2002	BMIDS00502 - INWP1, BM061, Complete change to the Introducer details
MO		04/10/2002	BMIDS00582 - BM061, Bug when entering the screen the value of direct/indirect 
									wasnt being noted.
MO		17/10/2002	BMIDS00663 - CC006, Change to the Introducer details, to capture who to 
									correspond with.  Also made
									the change so that omAdmin was called rather than omInt
MO		24/10/2002	BMIDS00713	- Fixed bug where on entry to the screen the Add button was disabled.
GD		15/11/2002	BMIDS00376 Disable screen if readonly
BS		12/06/2003	BM0521 Enable Details button if Introducer in list box
HMA     17/09/2003  BM0063 Amend HTML text for radio buttons
MC		11/05/2004	BMIDS744	Is Application Opt Out? section added to the screen
KRW     25/05/2004  BMIDS762   - FSA Validation
KRW     25/06/2004  BMIDS774   Regulation changes
KRW     25/06/2004  BMIDS776   LandUsage & Additional Changes
INR		14/06/2004	BMIDS744	added Opt Out processing.
KRW     25/06/2004  BMIDS776   Amemdments to LandUsage & Additional Changes
KRW     25/06/2004  BMIDS776   Further ammendements to Population of Regulation Indicator 
INR		08/07/2004  BMIDS776   Check whether the Regulation indicator needs to be reset 
INR		12/07/2004	BMIDS776   txtApplicationDateReceived should be txtEstimatedCompletionDate for ContextParameter idXML
INR		12/07/2004	BMIDS776   Don't use global variable m_sCurrLandUsage get value as required.
KRW     13/07/2004  BMIDS766   Changed processing to check for a null value when returning from DC013
INR		14/07/2004	BMIDS776   Delete Family Let info if special scheme combo is changed from buy to let.
SR		19/07/2004  BMIDS802
HMA     26/07/2004  BMIDS806   Save TPDDeclaration answer.
HMA     11/08/2004  BMIDS828   Remove btnIntermedSearch reference on submit.
GHun	10/09/2004	BMIDS871   Fixed spelling of ApplicationReceivedDate
JD		21/09/2004  BMIDS887   default land usage as YES if no value. Set readonly flag correctly
GHun	22/09/2004	BMIDS887   Call DeletePropertyData on Submit, rather than when SpecialScheme changes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History :

Prog	Date		AQR			Description
GHun	22/07/2005	MAR10		Made DC013 window larger, changed client specific label, fixed HTML errors
MF		22/07/2005	IA_WP01		Change of process flow & default opt-out and advice combos
GHun	02/08/2005	MAR14		Prevent combos from being hidden when using the menu
MahaT	10/10/2005	MAR94		Disable Introducer Add button.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History :

Prog	Date		AQR			Description
pct		21/03/2005	EP8			Miscelaneous Changes
pct		24/03/2005	EP15		Changes to accomodated Packager/Broker Intermediaries 
IK		19/04/2006	EP371		add self-certified indicator
AW		04/05/2006	EP393		Extended Packager/Network, moved remaining fields down
AW		09/05/2006	EP485		Added extra fee types (Broker, Third party valuation)
IK		09/05/2006	EP518		retrieve INGESTIONAPPLICATIONNUMBER & make read-only
PB		19/05/2006	EP586		'Delete' but enabled when in readonly mode
SAB		22/05/2006	EP565		Checks that the date received is not a date in the future
PE		09/06/2006	EP695		Round down values (txtBrokerFee, txtBrokerFeeRefund, txtPackagerValFee, txtPackagerValFeeRebate)
PB		30/06/2006	EP847		Made 'KFI Received by all applicants' mandatory
SAB		06/07/2006  EP945		Display the individual name rather than the contact name when a user
								elects to display the Packager in DC011
AW		08/08/2006	EP1076		Reworked EP945
PB		21/09/2006	EP1117		CC59 - Additional packager contact details
INR		31/10/2006	EP2_12		Changes for New Intermediary structure
AShaw	10/11/2006	EP2_8		Change to Save/Cancel routing, and add new fields.
INR		15/11/2006	EP2_12		Changes for New Intermediary structure
GHun	20/11/2006	EP2_123		Fixed misnamed variable InitialiseScreenValues
MCh		28/11/2006	EP2_127		Disable txtNumberOfDependants if DCDependantsAtApplicationLevel = False
PSC		04/12/2006	EP2_249		Add AssociationFee and PackagingFee and populate these and ProcurationFee from
                                MortgageIntroducerFee
MAH		11/12/2006	EP2_319		Altered LegalRep Radio Group
MAH		12/12/2006	EP2_101		Removed SelfCert Option
INR		09/01/2007	EP2_701		Need to save Credit Scheme
PE		19/01/2007	EP2_847		Incorrect screen navigation on cancel button
INR		22/01/2007	EP2_677		E2ECR35 review changes
INR		22/01/2007	EP2_645		Added Self Cert Reason
SR		29/01/2007	EP2_115		modified btnDeail.onClick.
SR		29/01/2007	EP2_115		modified btnDeail.onClick. use getRowSelectedIndex() to find the row selected.
PSC		01/02/2007	EP2_1112	Disable nature of loan and credit scheme if returning customer
DS		05/02/2007	EP2_966		Added code to persist value of 'Legal Rep to be used' radio button
IK		06/02/2007	EP2_1141	Disable income status if returning customer
PSC		07/02/2007	EP2_373		Don't validate KFI received if in read only mode
PB		16/02/2007	EP2_1475	Added fixed widths to SPAN objects to stop text moving when font is modified
AShaw	19/02/2007	EP2_705		New method to save data Without Critical Data check for QQ cases.
AShaw	20/02/2007	EP2_1202	Remove unnecessary calls to DB.
PB		21/02/2007	EP2_1475	Added fixed widths to more SPAN objects to stop text moving when font is modified
SR		01/03/2007	EP2_1272/EP2_1196	
SR		07/03/2007	EP2_1335	Add combo ChannelId
SR		07/03/2007	EP2_1335	change the position of combo ChannelId
SR		12/03/2007	EP2_1335	use global parameters in logic - to decide enable/disable channelId combo
SR		14/0/2007	EP2_1335	Enable the combo ChannelId when the value selected is Web rather than 'Paper'
SR		14/0/2007	EP2_1335	Enable the combo when no value was selected - this is not followed for any case in general.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
	<head>
		<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4" />
		<link href="stylesheet.css" rel="stylesheet" type="text/css" />
		<title></title>
	</head>
	<body>
		<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
		<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden"
			tabIndex="-1" type="text/x-scriptlet" width="1" viewastext>
		</object>
		<script src="validation.js" language="JScript"></script>
		<% /* removed as per Core V7.0.2  - DPF 21/06/02 - BMIDS00077
<object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
*/ %>
		<%/* FORMS */%>
		<form id="frmToDC015" method="post" action="DC015.asp" STYLE="DISPLAY: none">
		</form>
		<form id="frmToMN060" method="post" action="MN060.asp" STYLE="DISPLAY: none">
		</form>
		<form id="frmToDC020" method="post" action="DC020.asp" STYLE="DISPLAY: none">
		</form>
		<form id="frmToDC030" method="post" action="DC030.asp" STYLE="DISPLAY: none">
		</form>
		<form id="frmToDC012" method="post" action="DC012.asp" STYLE="DISPLAY: none">
		</form>
		<form id="frmToGN300" method="post" action="gn300.asp" STYLE="DISPLAY: none">
		</form>
		<form id="frmToMQ010" method="post" action="mq010.asp" STYLE="DISPLAY: none">
		</form>
		<form id="frmToDC016" method="post" action="DC016.asp" STYLE="DISPLAY: none">
		</form>
		<span id="spnToLastField" tabindex="0"></span>
		<span id="spnListScroll">
			<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
				<object data="scTableListScroll.asp" id="scScrollTable" style="HEIGHT: 0px; LEFT: 0px; TOP: 0px; WIDTH: 0px"
					tabindex="-1" type="text/x-scriptlet" viewastext>
				</object>
			</span>
		</span>
		<%/* SCREEN */%>
		<form id="frmScreen" mark validate="onchange" year4>
			<div id="divCreditOpt" style="HEIGHT: 34px; LEFT: 9px; POSITION: absolute; TOP: 60px; WIDTH: 607px"
				class="msgGroup">
				<span id="lblCreditCheckOptOutIndicator" style="LEFT: 11px; POSITION: absolute; TOP: 10px"
					class="msgLabel">
					<label id="idOptOutLabel"></label>
				</span>
				<span style="LEFT: 295px; POSITION: absolute; TOP: 7px" class="msgLabel">
					<input type="radio" value="0" id="idCreditCheckOptOutYes" disabled readonly name="CreditCheckOptOutIndicator">
					<label for="idCreditCheckOptOutYes" class="msgLabel">Yes</label>
					<input type="radio" value="1" id="idCreditCheckOptOutNo" disabled readonly name="CreditCheckOptOutIndicator">
					<label for="idCreditCheckOptOutNo" class="msgLabel">No</label>
				</span>
				<span style="LEFT: 450px; POSITION: absolute; TOP: 6px" class="msgLabel">
					<input id="btnTPDDeclaration" type="button" value="TPD Declaration" class="msgButton" style="LEFT: 0px; TOP: -1px">
				</span>
			</div>
			<div id="divApplicationSource" style="HEIGHT: 180px; LEFT: 9px; POSITION: absolute; TOP: 99px; WIDTH: 607px"
				class="msgGroup">
				<span style="LEFT: 10px; POSITION: absolute; TOP: 4px" class="msgLabel">
					Is at least 40% of the land/property<br>area to be used for residential<br>purposes? (BTL is considered<br>Residential)
					<span style="LEFT: 180px; POSITION: absolute; TOP: 2px; WIDTH: 100px;">
						<input id="optLandUsageYes" name="LandUsage" type="radio" value="1" style="LEFT: -1px; TOP: 0px">
						<label for="optLandUsageYes" class="msgLabel">Yes</label>
					</span> 
					<span style="LEFT: 226px; POSITION: absolute; TOP: 2px; WIDTH: 100px;">
						<input id="optLandUsageNo" name="LandUsage" type="radio" value="0" style="LEFT: 0px; TOP: 1px">
						<label for="optLandUsageNo" class="msgLabel">No</label>
					</span> 
				</span>
				<span style="LEFT: 301px; POSITION: absolute; TOP: 9px; color:red;" class="msgLabel">
					KFI Received by all applicants?
					<span style="LEFT: 160px; POSITION: absolute; TOP: -3px; WIDTH: 100px;">
						<input id="optKFIReceivedYes" name="KFIReceived" type="radio" value="1" style="LEFT: 1px; TOP: 0px">
						<label for="optKFIReceivedYes" class="msgLabel">Yes</label>
					</span> 
					<span style="LEFT: 206px; POSITION: absolute; TOP: -3px; WIDTH: 100px;">
						<input id="optKFIReceivedNo" name="KFIReceived" type="radio" value="0" style="LEFT: -1px; TOP: 0px">
						<label for="optKFIReceivedNo" class="msgLabel">No</label>
					</span> 
				</span>
				<span style="LEFT: 301px; POSITION: absolute; TOP: 37px" class="msgLabel" id="spnChannelID">
					Channel ID
					<span style="LEFT: 94px; POSITION: absolute; TOP: -3px">
						<select id="cboChannel" style="LEFT: 0px; TOP: -1px; WIDTH: 180px" class="msgCombo" menusafe="true" NAME="cboChannel">
						</select>
					</span>
				</span>					
				<span style="LEFT: 10px; POSITION: absolute; TOP: 68px" class="msgLabel">
					Level of Advice
					<span style="LEFT: 94px; POSITION: absolute; TOP: -3px">
						<select id="cboLevelOfAdvice" style="WIDTH: 180px" class="msgCombo" menusafe="true">
						</select>
					</span>
				</span>
				<span style="LEFT: 10px; POSITION: absolute; TOP: 96px" class="msgLabel">
					Direct or Indirect
					<span style="LEFT: 94px; POSITION: absolute; TOP: -3px">
						<select id="cboDirectIndirect" style="LEFT: 0px; TOP: 1px; WIDTH: 180px" class="msgCombo"
							menusafe="true">
						</select>
					</span>
				</span>
				<span style="LEFT: 10px; POSITION: absolute; TOP: 124px" class="msgLabel">
					Marketing Source
					<span style="LEFT: 94px; POSITION: absolute; TOP: -3px">
						<select id="cboMarketingSource" style="WIDTH: 180px" class="msgCombo" menusafe="true" NAME="cboMarketingSource">
						</select>
					</span>
				</span>
				<span style="LEFT: 10px; POSITION: absolute; TOP: 152px" class="msgLabel">
					Income Status
					<span style="LEFT: 94px; POSITION: absolute; TOP: -3px">
						<select id="cboAppIncomeStatus" style="WIDTH: 180px" class="msgCombo" menusafe="true" NAME="cboAppIncomeStatus">
						</select>
					</span>
				</span>
				<span style="LEFT: 10px; POSITION: absolute; TOP: 180px" class="msgLabel">
					Self-cert Reason
					<span style="LEFT: 94px; POSITION: absolute; TOP: -3px">
						<select id="cboSelfCertReason" style="WIDTH: 180px" class="msgCombo" menusafe="true" NAME="cboSelfCertReason">
						</select>
					</span>
				</span>
				<span style="LEFT: 301px; POSITION: absolute; TOP: 68px" class="msgLabel">
					Regulation Indicator
					<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
						<select id="cboRegulationIndicator" style="LEFT: 1px; TOP: 0px; WIDTH: 180px" class="msgCombo"
							menusafe="true">
						</select>
					</span>
				</span>
				<span style="LEFT: 301px; POSITION: absolute; TOP: 96px" class="msgLabel">
					Special Scheme
					<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
						<select id="cboSpecialScheme" style="LEFT: 1px; TOP: 0px; WIDTH: 180px" class="msgCombo"
							menusafe="true">
						</select>
					</span>
				</span>
				<span style="LEFT: 496px; POSITION: absolute; TOP: 97px" class="msgLabel">
					<span style="LEFT: 85px; POSITION: absolute; TOP: -8px">
						<input id="btnSpecialScheme" type="button" style="HEIGHT: 26px; LEFT: -1px; TOP: 0px; WIDTH: 26px"
							class="msgDDButton">
					</span>
				</span>
				<span style="LEFT: 301px; POSITION: absolute; TOP: 124px" class="msgLabel">
					Nature of Loan
					<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
						<select id="cboNatureOfLoan" style="LEFT: 0px; TOP: -1px; WIDTH: 180px" class="msgCombo"
							menusafe="true" NAME="cboNatureOfLoan">
						</select>
					</span>
				</span>
				<span style="LEFT: 301px; POSITION: absolute; TOP: 152px" class="msgLabel">
					Credit Scheme
					<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
						<select id="cboCreditScheme" style="LEFT: 0px; TOP: -1px; WIDTH: 180px" class="msgCombo"
							menusafe="true" NAME="cboCreditScheme">
						</select>
					</span>
				</span>
				<span style="LEFT: 301px; POSITION: absolute; TOP: 180px" class="msgLabel">
					Details
					<span style="LEFT: 60px; POSITION: absolute; TOP: -3px">
						<textarea class="msgTxt" id="txtSelfCertDetails" rows="2" style="LEFT: 0px; POSITION: absolute; TOP: 2px; WIDTH: 220px"
							NAME="txtSelfCertDetails"></textarea>
					</span>
				</span>
				<span style="LEFT: 10px; POSITION: absolute; TOP: 220px" class="msgLabel">
					Estimated<br>Completion Date
					<span style="LEFT: 94px; POSITION: absolute; TOP: 2px">
						<input id="txtEstimatedCompletionDate" type="text" maxlength="10" class="msgTxt" style="HEIGHT: 20px; WIDTH: 100px">
					</span>
				</span>
				<span style="LEFT: 405px; POSITION: absolute; TOP: 220px" class="msgLabel">
					Application<br>Date Received
					<span style="LEFT: 75px; POSITION: absolute; TOP: 2px">
						<input id="txtApplicationDateReceived" type="text" maxlength="10" class="msgTxt" style="HEIGHT: 20px; LEFT: 4px; TOP: 1px; WIDTH: 100px">
					</span>
				</span>
				<% /* EP8 pct 21/03/2006 */ %>
				<span style="LEFT: 220px; POSITION: absolute; TOP: 220" class="msgLabel">
					Application<br>Date Signed
					<span style="LEFT: 65px; POSITION: absolute; TOP: 2px">
						<input id="txtApplicationDateSigned" type="text" maxlength="10" class="msgTxt" style="HEIGHT: 20px; LEFT: 4px; TOP: 1px; WIDTH: 100px">
					</span>
				</span>
				<% /* EP2_645 */ %>
				<span style="LEFT: 10px; POSITION: absolute; TOP: 263px" class="msgLabel">
					Is Legal Rep to be used
					<span style="LEFT: 170px; POSITION: absolute; TOP: -3px;width:55px;">
						<input id="optLegalRepYes" name="LegalRep" type="radio" value="1" style="LEFT: -1px; TOP: 0px">
						<label for="optLegalRepYes" class="msgLabel">Yes</label>
					</span> 
					<span style="LEFT: 213px; POSITION: absolute; TOP: -3px;width:55px;">
						<input id="optLegalRepNo" name="LegalRep" type="radio" value="0" style="LEFT: 0px; TOP: 1px">
						<label for="optLegalRepNo" class="msgLabel">No</label>
					</span> 
				</span>
				<span style="LEFT: 301px; POSITION: absolute; TOP: 263px;" class="msgLabel">
					Total Number of Dependants
					<span style="LEFT: 210px; POSITION: absolute; TOP: -3px">
						<input id="txtNoOfDependents" type="text" maxlength="2" class="msgTxt" style="HEIGHT: 20px; WIDTH: 70px">
					</span> 					
				</span>
			</div>
			<div id="divIntermediaryDetails" style="HEIGHT: 283px; LEFT: 9px; POSITION: absolute; TOP: 380px; WIDTH: 607px"
				class="msgGroup">
				<% /* BMIDS00502 - BM061 - Introducer details - Start */ %>
				<span style="LEFT: 10px; POSITION: absolute; TOP: 9px" class="msgLabel">
					<strong>Introducer Details</strong>
				</span>
				<span style="LEFT: 10px; POSITION: absolute; TOP: 30px" class="msgLabel">
					Packager<br>Application Ref
					<span style="LEFT: 94px; POSITION: absolute; TOP: 2px">
						<input id="txtPackagerApplicationRef" type="text" maxlength="10" class="msgReadOnly" readonly="true"
							style="HEIGHT: 20px; LEFT: 4px; TOP: 1px; WIDTH: 100px">
					</span>
				</span>
			
				<% /* EP15 pct END */ %>
				<div id="divIntroducerDetailsTable" style="HEIGHT: 37px; LEFT: 7px; POSITION: absolute; TOP: 66px; WIDTH: 595px">
					<table id="tblIntroducerDetails" width="594" border="0" cellspacing="0" cellpadding="0"
						class="msgTable" style="LEFT: 0px; TOP: 1px">
						<tr id="rowTitles">
							<td width="16%" class="TableHead">Introducer Id</td>
							<td width="13%" class="TableHead">FSA Ref #</td>
							<td width="12%" class="TableHead">Name</td>
							<td width="27%" class="TableHead">Address</td>
							<td width="8%" class="TableHead">Type</td>
							<td width="16%" class="TableHead">Listing Status</td>
							<td width="8%" class="TableHead">Status</td>
						</tr>
						<tr id="row01">
							<td width="16%" class="TableTopLeft">&nbsp;</td>
							<td width="13%" class="TableTopCenter">&nbsp;</td>
							<td width="12%" class="TableTopCenter">&nbsp;</td>
							<td width="27%" class="TableTopCenter">&nbsp;</td>
							<td width="8%" class="TableTopCenter">&nbsp;</td>
							<td width="16%" class="TableTopCenter">&nbsp;</td>
							<td width="8%" class="TableTopRight">&nbsp;</td>
						</tr>
						<tr id="row02">
							<td width="16%" class="TableLeft">&nbsp;</td>
							<td width="13%" class="TableCenter">&nbsp;</td>
							<td width="12%" class="TableCenter">&nbsp;</td>
							<td width="27%" class="TableCenter">&nbsp;</td>
							<td width="8%" class="TableCenter">&nbsp;</td>
							<td width="16%" class="TableCenter">&nbsp;</td>
							<td width="8%" class="TableRight">&nbsp;</td>
						</tr>
						<tr id="row03">
							<td width="16%" class="TableLeft">&nbsp;</td>
							<td width="13%" class="TableCenter">&nbsp;</td>
							<td width="12%" class="TableCenter">&nbsp;</td>
							<td width="27%" class="TableCenter">&nbsp;</td>
							<td width="8%" class="TableCenter">&nbsp;</td>
							<td width="16%" class="TableCenter">&nbsp;</td>
							<td width="8%" class="TableRight">&nbsp;</td>
						</tr>
						<tr id="row04">
							<td width="16%" class="TableLeft">&nbsp;</td>
							<td width="13%" class="TableCenter">&nbsp;</td>
							<td width="12%" class="TableCenter">&nbsp;</td>
							<td width="27%" class="TableCenter">&nbsp;</td>
							<td width="8%" class="TableCenter">&nbsp;</td>
							<td width="16%" class="TableCenter">&nbsp;</td>
							<td width="8%" class="TableRight">&nbsp;</td>
						</tr>
						<tr id="row05">
							<td width="16%" class="TableLeft">&nbsp;</td>
							<td width="13%" class="TableCenter">&nbsp;</td>
							<td width="12%" class="TableCenter">&nbsp;</td>
							<td width="27%" class="TableCenter">&nbsp;</td>
							<td width="8%" class="TableCenter">&nbsp;</td>
							<td width="16%" class="TableCenter">&nbsp;</td>
							<td width="8%" class="TableRight">&nbsp;</td>
						</tr>
						<tr id="row06">
							<td width="16%" class="TableBottomLeft">&nbsp;</td>
							<td width="13%" class="TableBottomCenter">&nbsp;</td>
							<td width="12%" class="TableBottomCenter">&nbsp;</td>
							<td width="27%" class="TableBottomCenter">&nbsp;</td>
							<td width="8%" class="TableBottomCenter">&nbsp;</td>
							<td width="16%" class="TableBottomCenter">&nbsp;</td>
							<td width="8%" class="TableBottomRight">&nbsp;</td>
						</tr>
					</table>
				</div>
				<span style="LEFT: 430px; POSITION: absolute; TOP: 170px">
					<input id="btnAdd" value="Add" type="button" style="LEFT: 0px; TOP: 1px; WIDTH: 50px" class="msgButton">
				</span>
				<span style="LEFT: 490px; POSITION: absolute; TOP: 170px">
					<input id="btnDetail" value="Detail" type="button" style="LEFT: 0px; TOP: 1px; WIDTH: 50px"
						class="msgButton">
				</span>
				<span style="LEFT: 550px; POSITION: absolute; TOP: 170px">
					<input id="btnDelete" value="Delete" type="button" style="LEFT: 0px; TOP: 1px; WIDTH: 50px"
						class="msgButton">
				</span>
				<% /* MV - 24/05/2002 - BMIDS00013 - BM044 - Amend When Application  */%>
				<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
				<% /* MAR10 GHun changed text to more generic rather than client specific */ %>
				<span style="LEFT: 10px; POSITION: absolute; TOP: 195px" class="msgLabel">
					Has the declaration been received?
					<span style="LEFT: 200px; POSITION: absolute; TOP: -3px">
						<input id="optDeclarationYes" name="Declaration" type="radio" value="1"><label for="optDeclarationYes" class="msgLabel">Yes</label>
					</span> 
					<span style="LEFT: 245px; POSITION: absolute; TOP: -3px">
						<input id="optDeclarationNo" name="Declaration" type="radio" value="0"><label for="optDeclarationNo" class="msgLabel">No</label>
					</span> 
				</span>
				<% /* PB EP2_1475 16/02/2007 Begin */ %>
				<span id="spnPackagedFully" style="LEFT: 320px; POSITION: absolute; TOP: 195px; WIDTH: 250px;"
					class="msgLabel">
					Is the Application Packaged Fully?
					<span style="LEFT: 170px; POSITION: absolute; TOP: -3px; WIDTH: 40px;">
						<input id="optApplPackagedFullyYes" name="ApplPackagedFullyGroup" type="radio" value="1"
							style="LEFT: 0px; TOP: 1px"><label for="optApplPackagedFullyYes" class="msgLabel">Yes</label>
					</span> 
					<span style="LEFT: 215px; POSITION: absolute; TOP: -3px; WIDTH: 40px;">
						<input id="optApplPackagedFullyNo" name="ApplPackagedFullyGroup" type="radio" value="0">
						<label for="optApplPackagedFullyNo" class="msgLabel">No</label>
					</span> 
				</span>
				<% /* PB EP2_1475 End */ %>
				<span style="LEFT: 10px; POSITION: absolute; TOP: 231px" class="msgLabel">
					Additional Details
					<span style="LEFT: 105px; POSITION: absolute; TOP: -15px">
						<textarea class="msgTxt" id="txtAdditionalDetails" rows="4" style="LEFT: 0px; POSITION: absolute; TOP: 2px; WIDTH: 486px"></textarea>
					</span>
				</span>
				<% /* EP15 pct
				<span style="LEFT: 280px; POSITION: absolute; TOP: 214px" class="msgLabel">
					Additional Broker <BR>Fee Description
					<span style="LEFT: 89px; POSITION: absolute; TOP: 3px">
						<input id="txtAdditionalBrokerFeeDesc" type="text" maxlength="100"  class="msgTxt" style="HEIGHT: 20px; WIDTH: 232px">
					</span>
				</span>
				*/ %>
				<span style="LEFT: 9px; POSITION: absolute; TOP: 290px" class="msgLabel">
					Procuration Fee
					<span style="LEFT: 105px; POSITION: absolute; TOP: -3px">
						<input id="txtProcurationFee" type="text" maxlength="9" class="msgReadOnly" readonly="true"
							style="LEFT: 14px; TOP: 1px">
					</span>
				</span>
				<% /* EP15 END pct */ %>
				<span style="LEFT: 9px; POSITION: absolute; TOP: 320px" class="msgLabel">
					Association Fee
					<span style="LEFT: 105px; POSITION: absolute; TOP: -3px">
						<input id="txtAssociationFee" type="text" maxlength="9" class="msgReadOnly" readonly="true"
							style="LEFT: 14px; TOP: 1px">
					</span>
				</span>
				<span style="LEFT: 9px; POSITION: absolute; TOP: 350px" class="msgLabel">
					Packaging Fee
					<span style="LEFT: 105px; POSITION: absolute; TOP: -3px">
						<input id="txtPackagingFee" type="text" maxlength="9" class="msgReadOnly" readonly="true"
							style="LEFT: 14px; TOP: 1px">
					</span>
				</span>
				<% /* SR 21/02/2007 : EP2_1196 */ %>
				<span style="LEFT: 290px; POSITION: absolute; TOP: 290px" class="msgLabel">
					Proc Fee Passed on to Customer?
					<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
						<select id="cboProcFeePassedToCustomer" style="LEFT: 0px; TOP: -1px; WIDTH: 140px" class="msgCombo"
							menusafe="true">
						</select>
					</span>
				</span>
				<span style="LEFT: 290px; POSITION: absolute; TOP: 320px" class="msgLabel" id="spnProcFeeAmtPassOn">
					Amount of Proc Fee Passed on
					<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
						<input id="txtAmtProcFeePassedOn" type="text" maxlength="9" class="msgTxt" style="LEFT: 14px; TOP: 1px">
					</span>
				</span>
				<% /* SR 21/02/2007 : EP2_1196  - End*/ %>
				<% /* SR 21/02/2007 : EP2_1272 */ %>
				<span style="LEFT: 290px; POSITION: absolute; TOP: 350px" class="msgLabel">
					<input id="btnIntroducerFees" value="Fees" type="button" style="LEFT: 0px; TOP: 1px; WIDTH: 50px"
						class="msgButton">
				</span>
				<% /* SR 21/02/2007 : EP2_1272 - End */ %>
				<% /*  SR EP2_1272	
				<span style="LEFT: 9px; POSITION: absolute; TOP: 350px" class="msgLabel">
					Broker Fee
					<span style="LEFT: 105px; POSITION: absolute; TOP: -3px">
						<input id="txtBrokerFee" type="text" maxlength="6" class="msgTxt" style="LEFT: 14px; TOP: 1px">
					</span>
				</span>
				<span style="LEFT: 301px; POSITION: absolute; TOP: 350px" class="msgLabel">
					Broker Proc Fee Refund
					<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
						<input id="txtBrokerFeeRefund" type="text" maxlength="6" class="msgTxt" style="HEIGHT: 20px; WIDTH: 120px">
					</span>
				</span>
				<span style="LEFT: 9px; POSITION: absolute; TOP: 380px" class="msgLabel">
					Packager Valuation <BR>Fee
				<span style="LEFT: 105px; POSITION: absolute; TOP: 2px">
					<input id="txtPackagerValFee" type="text" maxlength="6" class="msgTxt" style="LEFT: 14px; TOP: 1px">
				</span>
				</span>
				<span style="LEFT: 301px; POSITION: absolute; TOP: 380px" class="msgLabel">
					Rebate
					<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
						<input id="txtPackagerValFeeRebate" type="text" maxlength="6" class="msgTxt" style="HEIGHT: 20px; WIDTH: 120px">
					</span>
				</span> */ %>
				<% /* BMIDS00502 - BM061 - Introducer details - End */ %>
			</div>
		</form>
		<div id="msgButtons" style="HEIGHT: 19px; LEFT: 12px; POSITION: absolute; TOP: 770px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" -->
		</div>
		<% /* Span to keep tabbing within this screen */ %>
		<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->
		<%/* File containing field attributes (remove if not required) */%>
		<!--  #include FILE="attribs/DC010attribs.asp" --> <!--  #include FILE="customise/DC010Customise.asp" -->
		<%/* CODE */%>
		<script language="JScript" type="text/javascript">
<!--
var AppXML = null;
var m_sStoredDataXML = "";
var scScreenFunctions;
var IntroducerXML = null;
var IntermedXML = null;

var m_sMetaAction = null;
var m_sContext = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sReadOnly = null;

var m_bIntermediaryDetailsFetched = false;
var sButtonList = new Array("Submit","Cancel");
var m_blnReadOnly = false;

var strIntroducerFSANumberLevel1= ""; 
var strIntroducerFSANumberLevel2= "";
var strIntroducerFSANumberLevel3= "";

var strIntroducerCorrespondIndLevel1 = "";
var strIntroducerCorrespondIndLevel2 = "";
var strIntroducerCorrespondIndLevel3 = "";

var strDirectIndirectBusiness = "";

var bIntroducersChanged = false;

<% /*BMIDS776 var m_sCurrLandUsage = ""; */ %>
var strTypeOfApplication = "";
var m_sRegulatedId = "";
var m_sNonRegulatedId = "";
var m_sCCARegulatedId = "";
var strEstimatedCompletionDate = "";
var sDate = "";
var m_sFamilyLet = "" ;
var m_sFamilyMember = "" ;
var m_sFamilyMemberValType = "";
var m_sSecuredPersonalLoan = "";

<% /* BMIDS744 */ %>
var m_sTPDDeclaration = "" ;
var bUpdateRequired = false;
var CreditCheckXML = null;
<% /* EP2_12 */ %>
var m_selIndex = "";
var XMLIntroducerSearchType = null;
var XMLListingStatus = null;
var XMLFirmStatus = null;
var m_iTableLength = 6;
<% /* BMIDS00502 - BM061 - Introducer details */ %>
var m_sQuickQuote = ""; //EP2_8

var lsIsLegalReptoBeUsed = false; // EP2_8
var iTotalNoOfDependants //EP2_127

<% /* PSC 04/12/2006 EP2_249 - Start */ %>
var m_sProcFee = "";
var m_sAssociationFee = "";
var m_sPackagingFee = "";
<% /* PSC 04/12/2006 EP2_249 - End */ %>

<% /* PSC 01/02/2007 EP2_1112 */ %>
var m_bReturningCustomer = false;
var m_bScreenDirty = false ;

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;

function window.onload()
{
	<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	<% /* next line replaced by line below - DPF 21/06/02 - BMIDS00077 */ %>
	<% /* scScreenFunctions = new scScrFunctions.ScreenFunctionsObject(); */ %>
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	FW030SetTitles("Application Details","DC010",scScreenFunctions);
	RetrieveContextData();
	
	<% /* This function is contained in the field attributes file (remove if not required) */ %>
	SetMasks();

	<% /* DJP SYS2564 for client customisation */ %>
	Customise();

	Validation_Init();
	

	<% /* BMIDS744 Has a Credit check been performed for this application*/%>
	bStatus = GetCreditCheckStatus();
	if(bStatus == true)
	{
		<% /* Don't let anyone change OptOutIndicator */%>
		scScreenFunctions.SetRadioGroupState(frmScreen, "CreditCheckOptOutIndicator", "R");
	}
	
	PopulateScreen();
	ShowMainButtons(sButtonList);
	
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
	else if (frmScreen.idCreditCheckOptOutYes.disabled == true)
		frmScreen.optLandUsageYes.focus();
	else
		<% /* BMIDS744 change focus */ %>
		frmScreen.idCreditCheckOptOutYes.focus();
		
	<% /* BMIDS744 change focus scScreenFunctions.SetFocusToFirstField(frmScreen); */ %>
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC010");
	<% /* GD BMIDS000376 START */ %>
	if (m_blnReadOnly == true)
	{
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnDelete.disabled = true;
		m_sReadOnly = "1"; <% /* JD BMIDS887 */ %>
		//frmScreen.btnDetail.disabled = true;
	}
	<% /* GD BMIDS000376 END */ %>
	<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
	ClientPopulateScreen();
}

<% /* BMIDS00502 - BM061 - Introducer details */ %>
function ResetIntroducerList()	{
	
	scScrollTable.clear();
	
}


function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function RetrieveContextData()
{
	/*DEBUG
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","01160028");
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber","1");
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber","02013118");
	scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber","1");
	//End Debug*/
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sContext = scScreenFunctions.GetContextParameter(window,"idProcessContext",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","APP0001");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","2");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	if(m_sMetaAction == "LoadIntermediaryFromDC015" || m_sMetaAction == "FromDC016") //SR EP2_1272
	{
		<% /* On return from DC015, use the stored data values to set the combos and the Introducer details */ %>
		m_sStoredDataXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
	}
	m_sQuickQuote = scScreenFunctions.GetContextParameter(window,"idApplicationMode",null); //EP2_8
	
	if(m_sMetaAction == "FromDC016")
	{
		if(scScreenFunctions.GetContextParameter(window,"idXML2",null) == "<SCREENDIRTY>1</SCREENDIRTY>")
			bUpdateRequired = true ;
	}
}

<% /* BMIDS00502 - BM061 - Introducer details - Start */ %>
function PopulateScreen()
{
	PopulateCombos();
	GetApplicationData();
	setStateOnProcFeeAmtPassOnToCust();
	setCboChannelStatus();  <% /* SR EP2_1335 */ %>
}

function PopulateCombos()
{
	var XMLDirectIndirect = null;
	var XMLMarketingSource = null;
	var XMLSpecialScheme = null;
	var XMLNatureOfLoan = null;
	var XMLLevelOfAdvice = null;
	var XMLRegulationIndicator = null;
	var XMLTypeOfMortgage = null;
	
	var xmlNonRegulatedCombo = null ;
	var XMLRegulationIndicator = null;

	//EP2_8 New combos
	var XMLApplicationIncomeStatus = null;
	var XMLSpecialGroup = null;
	<% /* EP2_645 */ %>
	var XMLSelfCert = null;

								
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var sGroupList = new Array("LevelOfAdvice","RegulationIndicator", "Direct/Indirect", "MarketingSource","SpecialSchemes", "NatureOfLoan","TypeOfMortgage","IntroducerSearchType","ListingStatus","ApplicationIncomeStatus",
							   "SpecialGroup","FirmStatus","SelfCertReason", "ProcFeeToCust");
	
	if(XML.GetComboLists(document,sGroupList))
	{
		XMLDirectIndirect = XML.GetComboListXML("Direct/Indirect");
		XMLMarketingSource = XML.GetComboListXML("MarketingSource");
		XMLSpecialSchemes = XML.GetComboListXML("SpecialSchemes");
		XMLNatureOfLoan = XML.GetComboListXML("NatureOfLoan");
		XMLLevelOfAdvice = XML.GetComboListXML("LevelOfAdvice");
		XMLRegulationIndicator = XML.GetComboListXML("RegulationIndicator");
		XMLTypeOfMortgage = XML.GetComboListXML("TypeOfMortgage");
		<% /* EP2_12 IntroducerSearchType Only used for validation */ %>
		XMLIntroducerSearchType = XML.GetComboListXML("IntroducerSearchType");
		XMLListingStatus = XML.GetComboListXML("ListingStatus");
		XMLFirmStatus = XML.GetComboListXML("FirmStatus");
		XMLApplicationIncomeStatus = XML.GetComboListXML("ApplicationIncomeStatus");
		XMLSpecialGroup = XML.GetComboListXML("SpecialGroup");
		<% /* EP2_645 */ %>
		XMLSelfCert = XML.GetComboListXML("SelfCertReason");
		
		var blnSuccess = true;
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboLevelOfAdvice,XMLLevelOfAdvice,true);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboRegulationIndicator,XMLRegulationIndicator,false);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboDirectIndirect,XMLDirectIndirect,false);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboMarketingSource,XMLMarketingSource,true);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboSpecialScheme,XMLSpecialSchemes,true);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboNatureOfLoan,XMLNatureOfLoan,false);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboAppIncomeStatus,XMLApplicationIncomeStatus,false);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboCreditScheme,XMLSpecialGroup,true);
		<% /* EP2_645 */ %>
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboSelfCertReason,XMLSelfCert,true);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboProcFeePassedToCustomer,
															XML.GetComboListXML("ProcFeeToCust"),true);
		if(	blnSuccess) blnSuccess & populateChannelCombo();
		if(blnSuccess == false)
		{
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			DisableMainButton("Submit");
		}
		else
		{
			<% /* default combo DirectIndirect to 'Direct' */ %>
			strDirectIndirectBusiness = "1";
			frmScreen.cboDirectIndirect.value = 1;
			<% /* default combo NatureOfLoan to 'Residential' */ %>
			frmScreen.cboNatureOfLoan.value = 1;
			frmScreen.cboRegulationIndicator.value = 1;
		}
	}

	<% /* KRW BMIDS774 */ %>	
	var saTemp = new Array();
	saTemp = XML.GetComboIdsForValidation("RegulationIndicator", "R", XMLRegulationIndicator);
	m_sRegulatedId = saTemp[0];
		
	saTemp = XML.GetComboIdsForValidation("RegulationIndicator", "N", XMLRegulationIndicator);
	m_sNonRegulatedId = saTemp[0];
	
	saTemp = XML.GetComboIdsForValidation("RegulationIndicator", "CCA", XMLRegulationIndicator);
	m_sCCARegulatedId = saTemp[0];

	XMLTypeOfMortgage = XML.GetComboListXML("TypeOfMortgage");
	
	saTemp = XML.GetComboIdsForValidation("TypeOfMortgage", "CCA", XMLTypeOfMortgage);
	m_sSecuredPersonalLoan = saTemp[0];
	
	<% /* Disable Regulation field KRW BMIDS774 */ %>
	//SetRegulationIndicator();
	frmScreen.cboRegulationIndicator.disabled = true;
	
	<% /* KRW BMIDS774 End */ %>	
			
	XML = null;
}

function populateChannelCombo()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window, null);
	XML.SetAttribute("CRUD_OP", "READ");
	XML.SetAttribute("ENTITY", "DISTRIBUTIONCHANNEL");
	XML.CreateActiveTag("DISTRIBUTIONCHANNEL");
	XML.RunASP(document, "omCRUDIf.asp");
	
	if(XML.IsResponseOK())
	{
		XML.PopulateComboFromXML(document, frmScreen.cboChannel, XML.XMLDocument, true, true, 
								 "DISTRIBUTIONCHANNEL", "CHANNELID", "CHANNELNAME", false);		
	}
	
	XML = null;
}

function setCboChannelStatus()  <% /* SR EP2_1335 - new function  */ %>
{
	var XML ;
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var iWebChannelId = parseInt(XML.GetGlobalParameterAmount(document, 'WebChannelId'));
	var iChannelSwitchFirstStageSeq = parseInt(XML.GetGlobalParameterAmount(document, 'ChannelSwitchFirstStageSeq'));
	var iChannelSwitchLastStageSeq = parseInt(XML.GetGlobalParameterAmount(document, 'ChannelSwitchLastStageSeq'));
	var iCurrentStageSeqNo = scScreenFunctions.GetContextParameter(window,"idCurrentStageOrigSeqNo",null);
	var XML = null
	if((frmScreen.cboChannel.value == iWebChannelId && 
		(iCurrentStageSeqNo >= iChannelSwitchFirstStageSeq && iCurrentStageSeqNo <= iChannelSwitchLastStageSeq)
	    ) || frmScreen.cboChannel.value == '')
		 scScreenFunctions.SetCollectionState(spnChannelID, "W");
	else scScreenFunctions.SetCollectionState(spnChannelID, "R");
}

function GetApplicationData()
{	
	<% /* PSC 04/12/2006 EP2_249 - Start */ %>
	var sMortgageSubQuoteNumber = "";
	var xmlQuoteData = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	xmlQuoteData.CreateRequestTag(window, null);
	xmlQuoteData.CreateActiveTag("APPLICATION");
	xmlQuoteData.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	xmlQuoteData.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	xmlQuoteData.RunASP(document,"GetAcceptedOrActiveQuoteData.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = xmlQuoteData.CheckResponse(ErrorTypes);
	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		var xmlMortgageSubQuote = xmlQuoteData.SelectTag(null,"MORTGAGESUBQUOTE");
		
		if (xmlMortgageSubQuote != null )
		{
			sMortgageSubQuoteNumber = xmlQuoteData.GetTagText("MORTGAGESUBQUOTENUMBER");
		}
	}
	
	if (sMortgageSubQuoteNumber != "")
	{
		var xmlIntroducerFees = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		xmlIntroducerFees.CreateRequestTag(window);
		var xmlRequest = xmlIntroducerFees.XMLDocument.documentElement;
		xmlRequest.setAttribute("CRUD_OP","READ");
		xmlRequest.setAttribute("ENTITY_REF","MORTGAGEINTRODUCERFEE");
		var xmlRoot = xmlIntroducerFees.XMLDocument.createElement("MORTGAGEINTRODUCERFEE");
		xmlRoot.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		xmlRoot.setAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		xmlRoot.setAttribute("MORTGAGESUBQUOTENUMBER", sMortgageSubQuoteNumber);
		xmlRequest.appendChild(xmlRoot);
		xmlIntroducerFees.RunASP(document, "omCRUDIf.asp");
	
		if(xmlIntroducerFees.IsResponseOK())
		{
		
			var xmlFees = xmlIntroducerFees.SelectNodes("MORTGAGEINTRODUCERFEE");
			var nIndex = 0;
			var xmlValidation = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
			while (nIndex < xmlFees.length)
			{
				var sFeeType = xmlFees[nIndex].attributes.getNamedItem("INTRODUCERFEETYPE").value
				
				if( xmlValidation.IsInComboValidationList(document,"IntroducerFeeType", sFeeType, Array("PR")))
				{
					m_sProcFee = xmlFees[nIndex].attributes.getNamedItem("FEEAMOUNT").value
				}
				else if ( xmlValidation.IsInComboValidationList(document,"IntroducerFeeType", sFeeType, Array("A")))
				{
					m_sAssociationFee = xmlFees[nIndex].attributes.getNamedItem("FEEAMOUNT").value
				}
				else if ( xmlValidation.IsInComboValidationList(document,"IntroducerFeeType", sFeeType, Array("PA")))
				{
					m_sPackagingFee = xmlFees[nIndex].attributes.getNamedItem("FEEAMOUNT").value
				}
				
				nIndex++;
			}		
		}
	}
	<% /* PSC 04/12/2006 EP2_249 - End */ %>		
	
	<% /* next line replaced by line below - DPF 21/06/02 - BMIDS00077 */ %>
	//AppXML = new scXMLFunctions.XMLObject();
	AppXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AppXML.CreateRequestTag(window, null)
				
	AppXML.CreateActiveTag("APPLICATION");
	AppXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);

	AppXML.RunASP(document,"GetApplicationData.asp");
	
	ErrorTypes = new Array("RECORDNOTFOUND");
	ErrorReturn = AppXML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
		InitialiseScreenValues();
	
	
	<% /* KRW 09/06/2004 BMIDS774 */ %>
	NewPropertyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	NewPropertyXML.CreateRequestTag(window, null); 
	NewPropertyXML.CreateActiveTag("NEWPROPERTY");
	NewPropertyXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	NewPropertyXML.CreateTag("APPLICATIONFACTFINDNUMBER", "1");
	
	NewPropertyXML.RunASP(document,"GetNewPropertyGeneral.asp")
	var ErrorReturn = NewPropertyXML.CheckResponse(ErrorTypes);
	if ( ErrorReturn[0]  || (ErrorReturn[1] == ErrorTypes[0]))
	{
		if(ErrorReturn[1] == null ||(ErrorReturn[1] != ErrorTypes[0]))
		{
			m_sNewPropertyExists = "1"
			NewPropertyXML.SelectTag(null,"NEWPROPERTY");
			<% /* EP2_677 */ %>
			m_sFamilyLet = NewPropertyXML.GetTagText("OCCUPYPROPERTY");
			m_sFamilyMember = NewPropertyXML.GetTagText("FAMILYMEMBER");		
		}
		else m_sNewPropertyExists = "0" ;
	}
	<% /* KRW 09/06/2004 :  End */ %>
	
	ErrorTypes = null;
	ErrorReturn = null;
	
}

<% /* BMIDS00502 - BM061 - Introducer details */ %>
function InitialiseScreenValues()
{
	//EP2_8 - Set new defaults.
	var sLegalRep = true;<%/* MAH 11/12/2006 EP2_319 Variable name changed */%>

	<% /* EP2_645 */ %>
	var appIncomeStatus = "";
	
	<% /* BMIDS744 */ %>
	var bStatus = false;
	<% /* Populate the combos from the stored data if available */ %>
	if( m_sStoredDataXML != "")
	{
		<% /* next line replaced by line below - DPF 21/06/02 - BMIDS00077
		var storedDataXML = new scXMLFunctions.XMLObject(); */ %>
		var storedDataXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		storedDataXML.LoadXML(m_sStoredDataXML);
		storedDataXML.SelectTag(null,"DC010SETTINGS");		
		<% /* MF MAR19 WP01 default advice to non-advised */ %>
		if(storedDataXML.GetTagText("LEVELOFADVICE")!=""){
			frmScreen.cboLevelOfAdvice.value = storedDataXML.GetTagText("LEVELOFADVICE");
		}else{			
			scScreenFunctions.SetComboOnValidationType(frmScreen, "cboLevelOfAdvice", "NADV");
		}

		frmScreen.cboRegulationIndicator.value = storedDataXML.GetTagText("REGULATIONINDICATOR"); 
		frmScreen.cboDirectIndirect.value = storedDataXML.GetTagText("DIRECTINDIRECTBUSINESS");
		strDirectIndirectBusiness = storedDataXML.GetTagText("DIRECTINDIRECTBUSINESS");
		frmScreen.cboMarketingSource.value = storedDataXML.GetTagText("MARKETINGSOURCE");
		frmScreen.cboSpecialScheme.value = storedDataXML.GetTagText("SPECIALSCHEME");
		frmScreen.cboNatureOfLoan.value = storedDataXML.GetTagText("NATUREOFLOAN");
		//strTypeOfApplication = storedDataXML("TYPEOFAPPLICATION"); 
		frmScreen.txtEstimatedCompletionDate.value= storedDataXML.GetTagText("ESTIMATEDCOMPLETIONDATE");
		frmScreen.txtApplicationDateReceived.value= storedDataXML.GetTagText("APPLICATIONRECEIVEDDATE"); 
		frmScreen.txtApplicationDateSigned.value= storedDataXML.GetTagText("APPLICATIONSIGNEDDATE"); // EP8 pct 21/03/2006
		//frmScreen.txtAdditionalBrokerFee.value= storedDataXML.GetTagText("ADDITIONALBROKERFEE"); // EP15 pct
		//frmScreen.txtAdditionalBrokerFeeDesc.value= storedDataXML.GetTagText("ADDITIONALBROKERFEEDESC"); // EP15 pct
		//frmScreen.txtPackagerNetwork.value= storedDataXML.GetTagText("PACKAGERNETWORKNAME");  // EP15 pct
		frmScreen.txtPackagerApplicationRef.value= storedDataXML.GetTagText("INGESTIONAPPLICATIONNUMBER");  // EP15 pct
		<% /* SR EP2_1272
		frmScreen.txtBrokerFee.value= storedDataXML.GetTagText("ADDITIONALBROKERFEE");  // EP485  
		frmScreen.txtBrokerFeeRefund.value= storedDataXML.GetTagText("BROKERPROCFEEREFUND");  // EP485  
		frmScreen.txtPackagerValFee.value= storedDataXML.GetTagText("PACKAGEDVALUATIONFEE");  // EP485 
		frmScreen.txtPackagerValFeeRebate.value= storedDataXML.GetTagText("PACKAGEDVALUATIONFEEREBATE");  // EP485 
		*/ %>
		frmScreen.txtNoOfDependents.value = storedDataXML.GetTagText("NUMBEROFDEPENDANTS"); //EP2_127   
		
		//EP2_8 Set values
		if (storedDataXML.GetTagText("APPLICATIONINCOMESTATUS") != "")
			frmScreen.cboAppIncomeStatus.value = storedDataXML.GetTagText("APPLICATIONINCOMESTATUS");
		else
			frmScreen.cboAppIncomeStatus.value = "20";
			
		frmScreen.cboCreditScheme.value = storedDataXML.GetTagText("PRODUCTSCHEME");
		sLegalRep = storedDataXML.GetTagBoolean("ISLEGALREPTOBEUSED");<%/* MAH 11/12/2006 EP2_319 Variable name changed */%>

			
		if(storedDataXML.GetTagText("APPLICATIONPACKAGEINDICATOR") != "")
		{
			if(storedDataXML.GetTagText("APPLICATIONPACKAGEINDICATOR") == "1")
				frmScreen.optApplPackagedFullyYes.checked = true;
			else
				frmScreen.optApplPackagedFullyNo.checked = true;
		}

		<% /* EP2_645 SelfCert Reason enabled, if Income Status is "SC" */ %>
		if((storedDataXML.GetTagText("APPLICATIONINCOMESTATUS")!="") && (scScreenFunctions.IsValidationType(frmScreen.cboAppIncomeStatus, "SC")))
		{
			frmScreen.cboSelfCertReason.disabled = false;
			frmScreen.cboSelfCertReason.value = storedDataXML.GetTagText("SELFCERTREASON");
			if((frmScreen.cboSelfCertReason.value != "") && (scScreenFunctions.IsValidationType(frmScreen.cboSelfCertReason, "O")))
			{
				scScreenFunctions.SetFieldState(frmScreen, "txtSelfCertDetails", "W");
				frmScreen.txtSelfCertDetails.setAttribute("required", "true");
				frmScreen.txtSelfCertDetails.value = storedDataXML.GetTagText("SELFCERTREASONDETAILS");
			}
			else
			{
				scScreenFunctions.SetFieldState(frmScreen, "txtSelfCertDetails", "R");
			}
		}
		else
		{
			frmScreen.cboSelfCertReason.disabled = true;
			scScreenFunctions.SetFieldState(frmScreen, "txtSelfCertDetails", "R");
		}	

		<% /* JD BMIDS887 if land usage is not set default as Yes
		// KRW LandUsage Radio BMIDS776
		if(storedDataXML.GetTagText("LANDUSAGE") != "")
		{ */ %>
			if(storedDataXML.GetTagText("LANDUSAGE") == "0")
				frmScreen.optLandUsageNo.checked = true;
			else
				frmScreen.optLandUsageYes.checked = true;
		<% //} %>	
	
		<% /* KRW KFI Radio KRW LandUsage Radio BMIDS774 */ %>
		if(storedDataXML.GetTagText("KFIRECEIVEDBYALLAPPS") != "")
		{
			if(storedDataXML.GetTagText("KFIRECEIVEDBYALLAPPS") == "1")
				frmScreen.optKFIReceivedYes.checked = true;
			else
				frmScreen.optKFIReceivedNo.checked = true;
		}

		/*DS - EP2_966 */
		if(storedDataXML.GetTagText("ISLEGALREPTOBEUSED") != "")
		{
			if(storedDataXML.GetTagText("ISLEGALREPTOBEUSED") == "1")
				frmScreen.optLegalRepYes.checked = true;
			else
				frmScreen.optLegalRepYes.checked = true;
		}	
			
		/*DS - EP2_966 */
	
		<% /* MV - 25/05/2002 - BMO44 - Amend BM Declaration */%> 
		if(storedDataXML.GetTagText("BMDECLARATIONIND") != "")
		{
			if(storedDataXML.GetTagText("BMDECLARATIONIND") == "1")
				frmScreen.optDeclarationYes.checked = true;
			else
				frmScreen.optDeclarationNo.checked = true;
		}
		
		frmScreen.txtAdditionalDetails.value = storedDataXML.GetTagText("ADDITIONALINTERMEDIARYDETAILS");
		
		strIntroducerCorrespondIndLevel1 = storedDataXML.GetTagText("INTRODUCERCORRESPONDINDLEVEL1");
		strIntroducerCorrespondIndLevel2 = storedDataXML.GetTagText("INTRODUCERCORRESPONDINDLEVEL2");
		strIntroducerCorrespondIndLevel3 = storedDataXML.GetTagText("INTRODUCERCORRESPONDINDLEVEL3");
		strTypeOfApplication = storedDataXML.GetTagText("TYPEOFAPPLICATION"); 

		<% /* BMIDS744 */%> 
		if(storedDataXML.GetTagText("OPTOUTINDICATOR") != "")
		{
			if(storedDataXML.GetTagText("OPTOUTINDICATOR") == "1")
				frmScreen.idCreditCheckOptOutYes.checked = true;
			else
				frmScreen.idCreditCheckOptOutNo.checked = true;
		}
		else
		{
			<% /* Should never get here, default it to Yes */%> 
			<% /* MF MAR19 default to no */ %>
			frmScreen.idCreditCheckOptOutNo.checked = true;
		}
		
		m_sTPDDeclaration  = storedDataXML.GetTagText("EXTERNALSYSTEMTPDDECLARATION");
		
		<% /* PSC 01/02/2007 EP2_1112 - Start */ %>
		if (storedDataXML.GetTagText("RETURNINGCUSTOMER") == "1")
			m_bReturningCustomer = true;
		else
			m_bReturningCustomer = false;
		<% /* PSC 01/02/2007 EP2_1112 - End */ %>
		<% /* SR 01/03/2007 : EP2_1196 */ %>
		frmScreen.cboProcFeePassedToCustomer.value = storedDataXML.GetTagText("PROCFEETOCUST");
		frmScreen.txtAmtProcFeePassedOn.value = storedDataXML.GetTagText("PROCFEETOCUSTAMOUNT");
		<% /* SR 01/03/2007 : EP2_1196 - End*/ %>
		frmScreen.cboChannel.value = storedDataXML.GetTagText("CHANNELID"); <% /* SR 01/03/2007 : EP2_1335 */ %>
		storedDataXML = null;
	}
	else
	{
		<% /* PSC 01/02/2007 EP2_1112 - Start */ %>
		AppXML.SelectTag(null, "APPLICATION");
		
		if (AppXML.GetTagText("ACCOUNTGUID")!="")
			m_bReturningCustomer = true;
		else
			m_bReturningCustomer = false;
		<% /* PSC 01/02/2007 EP2_1112 - End */ %>
		
		// EP2_1202 - Extract from AppXML
		iTotalNoOfDependants = AppXML.GetTagText("NUMBEROFDEPENDANTS");

		/*DS - EP2_966 */
		if(AppXML.GetTagText("ISLEGALREPTOBEUSED") != "")
		{
			// EP2_1202 - Extract ISLEGALREPTOBEUSED from AppXML
			lsIsLegalReptoBeUsed = AppXML.GetTagText("ISLEGALREPTOBEUSED");

			if(AppXML.GetTagText("ISLEGALREPTOBEUSED") == "1")
				frmScreen.optLegalRepYes.checked = true;
			else
				frmScreen.optLegalRepYes.checked = true;
		}	

		<% /* SR 01/03/2007 : EP2_1196 */ %>
		frmScreen.cboProcFeePassedToCustomer.value = AppXML.GetTagText("PROCFEETOCUST");
		frmScreen.txtAmtProcFeePassedOn.value = AppXML.GetTagText("PROCFEETOCUSTAMOUNT");
		<% /* SR 01/03/2007 : EP2_1196 - End */ %>
		frmScreen.cboChannel.value = AppXML.GetTagText("CHANNELID"); <% /* SR 01/03/2007 : EP2_1335 */ %>
		
		//	EP518
		if(AppXML.XMLDocument.selectSingleNode("RESPONSE/APPLICATIONLATESTDETAILS/APPLICATION/INGESTIONAPPLICATIONNUMBER"))
			frmScreen.txtPackagerApplicationRef.value= AppXML.XMLDocument.selectSingleNode("RESPONSE/APPLICATIONLATESTDETAILS/APPLICATION/INGESTIONAPPLICATIONNUMBER").text;
				
		AppXML.SelectTag(null, "APPLICATIONFACTFIND");
		
		//EP2_1202 - Code moved here
		if (AppXML.GetTagText("APPLICATIONINCOMESTATUS") != "")
		{
			appIncomeStatus = AppXML.GetTagText("APPLICATIONINCOMESTATUS");
			frmScreen.cboAppIncomeStatus.value = appIncomeStatus;
		}
		else
			frmScreen.cboAppIncomeStatus.value = "20";
			
		frmScreen.cboCreditScheme.value = AppXML.GetTagText("PRODUCTSCHEME");
		
		<% /* MF MAR19 WP01 default advice to non-advised */ %>		
		if (AppXML.GetTagText("LEVELOFADVICE")!=""){ 
			frmScreen.cboLevelOfAdvice.value = AppXML.GetTagText("LEVELOFADVICE"); 
		}else{
			scScreenFunctions.SetComboOnValidationType(frmScreen, "cboLevelOfAdvice", "NADV");
		}
		frmScreen.cboRegulationIndicator.value = AppXML.GetTagText("REGULATIONINDICATOR");
			
		strDirectIndirectBusiness = AppXML.GetTagText("DIRECTINDIRECTBUSINESS");
		if(strDirectIndirectBusiness != "") {
			frmScreen.cboDirectIndirect.value = strDirectIndirectBusiness;
		} else {
			<% /* MO - 24/10/2002 BMIDS00713 - Set directindirect default on load */%>	
			strDirectIndirectBusiness = "1";
			frmScreen.cboDirectIndirect.value = "1";
		}
		frmScreen.cboMarketingSource.value = AppXML.GetTagText("MARKETINGSOURCE");
		frmScreen.cboSpecialScheme.value = AppXML.GetTagText("SPECIALSCHEME");
		var NatureOfLoan = AppXML.GetTagText("NATUREOFLOAN");
		if(NatureOfLoan != "")
			frmScreen.cboNatureOfLoan.value = NatureOfLoan;
		
		if(AppXML.GetTagText("APPLICATIONPACKAGEINDICATOR") != "")
		{
			if(AppXML.GetTagText("APPLICATIONPACKAGEINDICATOR") == "1")
				frmScreen.optApplPackagedFullyYes.checked = true;
			else
				frmScreen.optApplPackagedFullyNo.checked = true;
		}

		<% /* JD BMIDS887 if land usage is not set default as Yes
			// KRW LandUsage Radio BMIDS774
		if(AppXML.GetTagText("LANDUSAGE") != "")
		{ */ %>
			if(AppXML.GetTagText("LANDUSAGE") == "0")
				frmScreen.optLandUsageNo.checked = true;
			else
				frmScreen.optLandUsageYes.checked = true;
		<% //}	%>
			
			<% /* KRW KFI Radio KRW BMIDS776 */ %>
		if(AppXML.GetTagText("KFIRECEIVEDBYALLAPPS") != "")
		{
			if(AppXML.GetTagText("KFIRECEIVEDBYALLAPPS") == "1")
				frmScreen.optKFIReceivedYes.checked = true;
			else
				frmScreen.optKFIReceivedNo.checked = true;
		}

			
		/*DS - EP2_966 */
	
		<% /* MV - 25/05/2002 - BMO44 - Amend BM Declaration */%> 
		if(AppXML.GetTagText("BMDECLARATIONIND") != "")
		{
			if(AppXML.GetTagText("BMDECLARATIONIND") == "1")
				frmScreen.optDeclarationYes.checked = true;
			else
				frmScreen.optDeclarationNo.checked = true;
		}
		
		frmScreen.txtAdditionalDetails.value = AppXML.GetTagText("ADDITIONALINTERMEDIARYDETAILS");
		
		strIntroducerCorrespondIndLevel1 = AppXML.GetTagText("INTRODUCERCORRESPONDINDLEVEL1");
		strIntroducerCorrespondIndLevel2 = AppXML.GetTagText("INTRODUCERCORRESPONDINDLEVEL2");
		strIntroducerCorrespondIndLevel3 = AppXML.GetTagText("INTRODUCERCORRESPONDINDLEVEL3");
		
		strTypeOfApplication = AppXML.GetTagText("TYPEOFAPPLICATION"); 
		frmScreen.txtEstimatedCompletionDate.value= AppXML.GetTagText("ESTIMATEDCOMPLETIONDATE"); 
		frmScreen.txtApplicationDateReceived.value= AppXML.GetTagText("APPLICATIONRECEIVEDDATE");
		frmScreen.txtApplicationDateSigned.value= AppXML.GetTagText("APPLICATIONSIGNEDDATE");
		//frmScreen.txtAdditionalBrokerFee.value= AppXML.GetTagText("ADDITIONALBROKERFEE"); // EP15 pct
		//frmScreen.txtAdditionalBrokerFeeDesc.value= AppXML.GetTagText("ADDITIONALBROKERFEEDESC");  // EP15 pct
		//frmScreen.txtPackagerNetwork.value= AppXML.GetTagText("PACKAGERNETWORKNAME");  // EP15 pct
		//	EP518
		//	frmScreen.txtPackagerApplicationRef.value= AppXML.GetTagText("INGESTIONAPPLICATIONNUMBER");  // EP15 pct
		<% /*  SR EP2_1272
		frmScreen.txtBrokerFee.value = AppXML.GetTagText("ADDITIONALBROKERFEE");  // EP485 
		frmScreen.txtBrokerFeeRefund.value = AppXML.GetTagText("BROKERPROCFEEREFUND");  // EP485  
		frmScreen.txtPackagerValFee.value = AppXML.GetTagText("PACKAGEDVALUATIONFEE");  // EP485  
		frmScreen.txtPackagerValFeeRebate.value = AppXML.GetTagText("PACKAGEDVALUATIONFEEREBATE");  // EP485  
		*/ %>
		frmScreen.txtNoOfDependents.value = iTotalNoOfDependants;  // EP2_127
		
		sLegalRep = lsIsLegalReptoBeUsed;<%/* MAH 11/12/2006 EP2_319 Variable name changed */%>
		
		<% /* BMIDS744 */%> 
		if(AppXML.GetTagText("OPTOUTINDICATOR") != "")
		{
			if(AppXML.GetTagText("OPTOUTINDICATOR") == "1")
				frmScreen.idCreditCheckOptOutYes.checked = true;
			else
				frmScreen.idCreditCheckOptOutNo.checked = true;
		}
		else
		{
			<% /* First time in, default it to Yes */%>
			<% /* MF Default to no */ %> 
			frmScreen.idCreditCheckOptOutNo.checked = true;
		}
		
		<% /* EP2_645 SelfCert Reason enabled, if Income Status is "SC" */ %>
		if((appIncomeStatus != "") && (scScreenFunctions.IsValidationType(frmScreen.cboAppIncomeStatus, "SC")))
		{
			frmScreen.cboSelfCertReason.disabled = false;
			frmScreen.cboSelfCertReason.value = AppXML.GetTagText("SELFCERTREASON");
			if((frmScreen.cboSelfCertReason.value != "") && (scScreenFunctions.IsValidationType(frmScreen.cboSelfCertReason, "O")))
			{
				scScreenFunctions.SetFieldState(frmScreen, "txtSelfCertDetails", "W");
				frmScreen.txtSelfCertDetails.setAttribute("required", "true");
				frmScreen.txtSelfCertDetails.value = AppXML.GetTagText("SELFCERTREASONDETAILS");
			}
			else
			{
				scScreenFunctions.SetFieldState(frmScreen, "txtSelfCertDetails", "R");
			}
		}
		else
		{
			frmScreen.cboSelfCertReason.disabled = true;
			scScreenFunctions.SetFieldState(frmScreen, "txtSelfCertDetails", "R");
		}		
		m_sTPDDeclaration = AppXML.GetTagText("EXTERNALSYSTEMTPDDECLARATION");
		
	}


	if(scScreenFunctions.IsValidationType(frmScreen.cboSpecialScheme, "BTL")) <% /* SR 19/07/2004 : BMIDS802 */ %>
	{
	   //frmScreen.btnSpecialScheme.disabled = false;
	   scScreenFunctions.EnableDrillDown(frmScreen.btnSpecialScheme);	
	}
	else
	{
		//frmScreen.btnSpecialScheme.disabled = true ;
		scScreenFunctions.DisableDrillDown(frmScreen.btnSpecialScheme);
	}
	
	<% /* EP8 - pct - 21/03/2006 */ %>
	/*
	if(frmScreen.txtApplicationDateReceived.value== "")
	{	
		sDate = scScreenFunctions.DateToString(scScreenFunctions.GetAppServerDate());
		frmScreen.txtApplicationDateReceived.value= sDate	;
	}
	*/
	<% /* EP8 - pct - 21/03/2006 */ %>
	
	<% /* EP2_127 */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();		
	frmScreen.txtNoOfDependents.disabled = (XML.GetGlobalParameterBoolean(document, "DCDependantsAtApplicationLevel") == true) ? false : true;
	XML = null; //Kill
	
	GetIntroducers();
	
	if(frmScreen.cboRegulationIndicator.value == "")
		SetRegulationIndicator();

	// EP2_8 - Enable if Additional Borrowing.
	var sidMortgageApplicationValue = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","");	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	scScreenFunctions.SetRadioGroupValue(frmScreen, "LegalRep", sLegalRep);<%/* MAH 11/12/2006 EP2_319 */%>
	
	if( XML.IsInComboValidationList(document,"TypeOfMortgage", sidMortgageApplicationValue, Array("F")) && XML.IsInComboValidationList(document,"TypeOfMortgage", sidMortgageApplicationValue, Array("M")))
	{
	<%/* MAH 11/12/2006 EP2_319
		// Set value 
		if (LegalRep == true)
			frmScreen.optLegalRepYes.checked = true;
		else
			frmScreen.optLegalRepNo.checked = true; */%>
			
		frmScreen.optLegalRepYes.disabled = false;
		frmScreen.optLegalRepNo.disabled = false;

	}
	else		
	{
	<%/* MAH 11/12/2006 EP2_319
		if (LegalRep == true)
			frmScreen.optLegalRepYes.checked = true;
		else
			frmScreen.optLegalRepNo.checked = true; 
			
		frmScreen.optLegalRepYes.checked = true;*/%>
		frmScreen.optLegalRepYes.disabled = true;
		frmScreen.optLegalRepNo.disabled = true; 
	}
	
	<% /* PSC 04/12/2006 EP2_249 - Start */ %>
	frmScreen.txtProcurationFee.value = m_sProcFee;
	frmScreen.txtAssociationFee.value = m_sAssociationFee;
	frmScreen.txtPackagingFee.value = m_sPackagingFee;
	<% /* PSC 04/12/2006 EP2_249 - End */ %>
	
	<% /* PSC 01/02/2007 EP2_1112 - Start */ %>
	if (m_bReturningCustomer)
	{
		if (frmScreen.cboCreditScheme.value != "")
		{
			scScreenFunctions.SetFieldToReadOnly(frmScreen, "cboCreditScheme");
		}
			
		if (frmScreen.cboNatureOfLoan.value != "")
		{
			scScreenFunctions.SetFieldToReadOnly(frmScreen, "cboNatureOfLoan");
		}
		<% /* IK 06/02/2007 EP2_1141 */ %>
		if (frmScreen.cboAppIncomeStatus.value != "")
		{
			scScreenFunctions.SetFieldToReadOnly(frmScreen, "cboAppIncomeStatus");
		}
	}
	<% /* PSC 01/02/2007 EP2_1112 - End */ %>
	
}

function GetIntroducers()
{
	frmScreen.btnAdd.disabled = true;
	frmScreen.btnDelete.disabled = true;
	frmScreen.btnDetail.disabled = true;
	
	<%/*Request is done as an operation to allow us to add SCREENREF for the postproc xslt*/%>
	IntroducerXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	IntroducerXML.CreateRequestTag(window, null);

	var xn = IntroducerXML.XMLDocument.documentElement;
	xn.setAttribute("CRUD_OP","READ");
	xn.setAttribute("SCHEMA_NAME","epsomCRUD");
	xn.setAttribute("ENTITY_REF","APPLICATIONINTRODUCERS");
	xn.setAttribute("RETURNREQUEST","Y");
	xn.setAttribute("postProcRef","FindIntroducersResponseGUI.xslt" );

	var xe = IntroducerXML.XMLDocument.createElement("OPERATION");
	xe.setAttribute("CRUD_OP","READ");
	xe.setAttribute("SCHEMA_NAME","epsomCRUD");
	xe.setAttribute("ENTITY_REF","APPLICATIONINTRODUCERS");
	xe.setAttribute("COMBOLOOKUP","Y");

	var xe2 = IntroducerXML.XMLDocument.createElement("APPLICATIONINTRODUCERS");
	xe2.setAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	xe2.setAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	xe.appendChild(xe2);

	xn.appendChild(xe);
	
	xs = IntroducerXML.XMLDocument.createElement("SEARCHCRITERIA");
	xs.setAttribute("SCREENREF", "DC010");
	xn.appendChild(xs);

	IntroducerXML.RunASP(document, "omCRUDIf.asp");

	if(!IntroducerXML.IsResponseOK()) return;

	PopulateIntroducerDetails();

	if(!IntroducerXML.XMLDocument.documentElement.selectSingleNode("APPLICATIONINTERMEDARIES/APPLICATION/BROKER"))
		frmScreen.btnAdd.disabled = false;
}

function PopulateIntroducerDetails()
{
	ResetIntroducerList();
	IntroducerXML.ActiveTag = null;
	IntroducerXML.CreateTagList("INTERMEDIARY");
	var iNumberOfIntermediaries = IntroducerXML.ActiveTagList.length;

	if(iNumberOfIntermediaries > 0)
	{
		scScrollTable.initialiseTable(tblIntroducerDetails, 0, "", ShowList, m_iTableLength, iNumberOfIntermediaries);
		ShowList();
	} 

}

function ShowList()
{

	var iCount;
	scScrollTable.clear();
	for (iCount = 0; iCount < IntroducerXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		IntroducerXML.ActiveTag = null;
		IntroducerXML.CreateTagList("INTERMEDIARY");  
		IntroducerXML.SelectTagListItem(iCount);

		var sType;
		var listingStatus;
		var introducerId;
		var nodValueID;
		var sStatus;
		var sAppIntroducerSeqNo = IntroducerXML.GetAttribute("APPLICATIONINTRODUCERSEQNO");
		var sPrincipalFirmID = IntroducerXML.GetAttribute("PRINCIPALFIRMID");
		var sARFirmID = IntroducerXML.GetAttribute("ARFIRMID");
		var sClubAssocID = IntroducerXML.GetAttribute("CLUBNETWORKASSOCIATIONID");
		var sIntroducerID = IntroducerXML.GetAttribute("INTRODUCERID");
		var fsaRef =  IntroducerXML.GetAttribute("FSAREFNUMBER");
		var sName =  IntroducerXML.GetAttribute("NAME");
		var sValidateType =  IntroducerXML.GetAttribute("LISTSTATUSVALIDATION");
		nodValueID=XMLListingStatus.selectSingleNode(".//LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE='"+ sValidateType +"']");
		if (nodValueID)
		{
			listingStatus =  nodValueID.selectSingleNode(".//VALUENAME").text;
		}
		
		var sStatusValue =  IntroducerXML.GetAttribute("STATUS");
		nodValueID =  XMLFirmStatus.selectSingleNode(".//LISTENTRY[VALUEID='"+ sStatusValue +"']");
		if (nodValueID)
		{
			sStatus = nodValueID.selectSingleNode(".//VALUENAME").text;
		}

		sValidateType = IntroducerXML.GetAttribute("TYPEVALIDATION");
		nodValueID=XMLIntroducerSearchType.selectSingleNode(".//LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE='"+ sValidateType +"']");
		if (nodValueID)
		{
			sType = nodValueID.selectSingleNode(".//VALUENAME").text;
		}
		var sPostcode = IntroducerXML.GetAttribute("POSTCODE");
		var sFlatAndOrHouseNumber = IntroducerXML.GetAttribute("ADDRESSLINE1");
		var sHouseName = IntroducerXML.GetAttribute("ADDRESSLINE2");
		var sStreet = IntroducerXML.GetAttribute("ADDRESSLINE3");
		var sDistrict = IntroducerXML.GetAttribute("ADDRESSLINE4");
		var sTown = IntroducerXML.GetAttribute("ADDRESSLINE5");
		var sCounty = IntroducerXML.GetAttribute("ADDRESSLINE6");
		
		<% //create an intelligent address line %>
		var sAddress = "";
		if(sFlatAndOrHouseNumber != "") sAddress += sFlatAndOrHouseNumber + ",";
		if(sHouseName != "") sAddress += sHouseName + ",";
		if(sStreet != "") sAddress += sStreet + ",";
		if(sDistrict != "") sAddress += sDistrict + ",";
		if(sTown != "") sAddress += sTown + ",";
		if(sCounty != "") sAddress += sCounty + ",";
		if(sPostcode != "") sAddress += sPostcode;

		if(sPrincipalFirmID&&(sPrincipalFirmID.length > 0))
		{
			tblIntroducerDetails.rows(iCount+1).setAttribute("PrincipalFirmID", sPrincipalFirmID);
			introducerId = sPrincipalFirmID;
		}
			
		if(sARFirmID&&(sARFirmID.length > 0))
		{
			tblIntroducerDetails.rows(iCount+1).setAttribute("ARFirmID", sARFirmID);
			introducerId = sARFirmID;
		}
		if(sClubAssocID&&(sClubAssocID.length > 0))
		{
			tblIntroducerDetails.rows(iCount+1).setAttribute("ClubAssocID", sClubAssocID);
			introducerId = sClubAssocID;
		}
		if(sIntroducerID&&(sIntroducerID.length > 0))
		{	
			tblIntroducerDetails.rows(iCount+1).setAttribute("IntroducerID", sIntroducerID);
			introducerId = sIntroducerID;
		}
		
		<% /* Need this in case we delete and if details is selected */%>
		if(sAppIntroducerSeqNo && (sAppIntroducerSeqNo.length > 0))
		{
			tblIntroducerDetails.rows(iCount+1).setAttribute("AppIntroducerSeqNo", sAppIntroducerSeqNo);
		}


		<% /* Add to the search table */%>
		scScreenFunctions.SizeTextToField(tblIntroducerDetails.rows(iCount+1).cells(0),introducerId);
		scScreenFunctions.SizeTextToField(tblIntroducerDetails.rows(iCount+1).cells(1),fsaRef);
		scScreenFunctions.SizeTextToField(tblIntroducerDetails.rows(iCount+1).cells(2),sName);
		scScreenFunctions.SizeTextToField(tblIntroducerDetails.rows(iCount+1).cells(3),sAddress);
		scScreenFunctions.SizeTextToField(tblIntroducerDetails.rows(iCount+1).cells(4),sType);
		scScreenFunctions.SizeTextToField(tblIntroducerDetails.rows(iCount+1).cells(5),listingStatus);
		scScreenFunctions.SizeTextToField(tblIntroducerDetails.rows(iCount+1).cells(6),sStatus);

		listingStatus = "";
		sStatus = "";
		sType = "";

	}
}
<% /* BMIDS00502 - BM061 - Introducer details */ %>
function OldPopulateIntroducerDetails () 
{
	
	<% /* EP945 - 06/07/2006 - Add Packager Indicator */%>

	<% /* Level 1 */ %>
	var xn = IntroducerXML.XMLDocument.documentElement.selectSingleNode("APPLICATIONINTERMEDARIES/APPLICATION/PACKAGER");
	if(xn)
	{
		scScreenFunctions.SizeTextToField(tblIntroducerDetails.rows(1).cells(0),"Packager");
		scScreenFunctions.SizeTextToField(tblIntroducerDetails.rows(1).cells(1),xn.getAttribute("INTERMEDIARYPANELID"));
		scScreenFunctions.SizeTextToField(tblIntroducerDetails.rows(1).cells(2),getDetails(xn));
		tblIntroducerDetails.rows(1).setAttribute("IntermediaryGuid", xn.getAttribute("INTERMEDIARYGUID"));
		tblIntroducerDetails.rows(1).setAttribute("IsPackager", 1);
	}

	<% /* Level 2 */ %>
	var xn = IntroducerXML.XMLDocument.documentElement.selectSingleNode("APPLICATIONINTERMEDARIES/APPLICATION/BROKER");
	if(xn)
	{
		scScreenFunctions.SizeTextToField(tblIntroducerDetails.rows(2).cells(0),"Broker");
		scScreenFunctions.SizeTextToField(tblIntroducerDetails.rows(2).cells(1),xn.getAttribute("INTERMEDIARYPANELID"));
		scScreenFunctions.SizeTextToField(tblIntroducerDetails.rows(2).cells(2),getDetails(xn));
		tblIntroducerDetails.rows(2).setAttribute("IntermediaryGuid", xn.getAttribute("INTERMEDIARYGUID"));
		tblIntroducerDetails.rows(2).setAttribute("IsPackager", 0);
	}
	function getDetails(xn)
	{
		var s = "";
		if(xn.selectSingleNode("INTERMEDIARYORGANISATION[@NAME]"))
			s += xn.selectSingleNode("INTERMEDIARYORGANISATION/@NAME").text + ",";
		if(xn.selectSingleNode("ADDRESS[@BUILDINGORHOUSENAME]"))
			s += xn.selectSingleNode("ADDRESS/@BUILDINGORHOUSENAME").text + ",";
		if(xn.selectSingleNode("ADDRESS[@BUILDINGORHOUSENUMBER]"))
			s += xn.selectSingleNode("ADDRESS/@BUILDINGORHOUSENUMBER").text + " ";
		if(xn.selectSingleNode("ADDRESS[@STREET]"))
			s += xn.selectSingleNode("ADDRESS/@STREET").text + ",";
		if(xn.selectSingleNode("ADDRESS[@TOWN]"))
			s += xn.selectSingleNode("ADDRESS/@TOWN").text + ",";
		if(xn.selectSingleNode("ADDRESS[@COUNTY]"))
			s += xn.selectSingleNode("ADDRESS/@COUNTY").text + ",";
		if(xn.selectSingleNode("ADDRESS[@POSTCODE]"))
			s += xn.selectSingleNode("ADDRESS/@POSTCODE").text + ",";
		return(s.substr(0,s.length -1));
	}
}

<% /* MO - 18/10/2002 - CH006 - Changes the correspondind for the appropriate yes / no on double click */ %>
function divIntroducerDetailsTable.ondblclick () 
{
	<% /* GD BMIDS00376 */ %>
	if (m_blnReadOnly == false)
	{
		<% /* EP2_12 */ %>
		m_selIndex = scScrollTable.getRowSelected();

	}
}

function ConvertBoolToYesNo(strValue) {
	if (strValue == "1") {
		return "Yes"
	} else {
		return "No"
	}
}

function ConvertYesNoToBool(strValue) {
	if (strValue.toUpperCase == "YES") {
		return "1"
	} else {
		return "0"
	}
}

function WriteApplicationSourceDetails()
{
	var bSuccess = true;
	
	<% /* BMIDS887 GHun */ %>
	<% /* INR BMIDS776 Clear out Family Let info if cboSpecialScheme moves from BTL  */ %>
	if(!(scScreenFunctions.IsValidationType(frmScreen.cboSpecialScheme, "BTL")))
	{
		if((m_sFamilyMember != "") || (m_sFamilyLet != ""))
		{
			bSuccess = DeletePropertyData();
			if (bSuccess == false)
			{
				alert("Failed to delete Family Let information.");
			}
		}
	}
		
	if (bSuccess)
	{
	<% /* BMIDS887 End */ %>
	
		<% /* next line replaced with line below - DPF 21/06/02 - BMIDS00077 */ %>
		//var XML = new scXMLFunctions.XMLObject();
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window, null)
		
		//EP558: OPERATION LH_17/05/2001
		XML.SetAttribute("OPERATION","CriticalDataCheck");
	
		AppXML.SelectTag(null, "APPLICATION");
		var tagAPPLICATION = XML.CreateActiveTag("APPLICATION");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("PACKAGENUMBER", AppXML.GetTagText("PACKAGENUMBER"));
		XML.CreateTag("NUMBEROFDEPENDANTS", frmScreen.txtNoOfDependents.value) // EP2_127
		/*DS - EP2_966 */
		if(frmScreen.optLegalRepYes.checked)
			XML.CreateTag("ISLEGALREPTOBEUSED", "1");
		else if(frmScreen.optLegalRepNo.checked)
			XML.CreateTag("ISLEGALREPTOBEUSED", "0");
		else
			XML.CreateTag("ISLEGALREPTOBEUSED", "");
		
		XML.CreateTag("CHANNELID", frmScreen.cboChannel.value); <%/*  SR EP2_1335 */%>
			
		/*DS - EP2_966 */
		<% /* SR 01/03/2007 : EP2_1196 */ %>
		XML.CreateTag("PROCFEETOCUST", frmScreen.cboProcFeePassedToCustomer.value);
		XML.CreateTag("PROCFEETOCUSTAMOUNT", frmScreen.txtAmtProcFeePassedOn.value);
		<% /* SR 01/03/2007 : EP2_1196 - End */ %>
		
		var tagAPPLICATIONFACTFIND = XML.CreateActiveTag("APPLICATIONFACTFIND");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("LEVELOFADVICE", frmScreen.cboLevelOfAdvice.value); 
		XML.CreateTag("REGULATIONINDICATOR", frmScreen.cboRegulationIndicator.value);
		XML.CreateTag("DIRECTINDIRECTBUSINESS", frmScreen.cboDirectIndirect.value);
	
		if(frmScreen.optApplPackagedFullyYes.checked)
			XML.CreateTag("APPLICATIONPACKAGEINDICATOR", "1");
		else if(frmScreen.optApplPackagedFullyNo.checked)
			XML.CreateTag("APPLICATIONPACKAGEINDICATOR", "0");
		else
			XML.CreateTag("APPLICATIONPACKAGEINDICATOR", "");

		<% /* IK EP371 self-certified indicator */ %>
		<%/*MAH EP2_101 Removed SelfCert
		if(frmScreen.optSelfCertYes.checked)
			XML.CreateTag("SELFCERTIND", "1");
		else if(frmScreen.optSelfCertNo.checked)
			XML.CreateTag("SELFCERTIND", "0");
		else
			XML.CreateTag("SELFCERTIND", "");
		*/%>
		<% /* KRW BMIDS776 - Land Usage */%>	
		if(frmScreen.optLandUsageYes.checked)
			XML.CreateTag("LANDUSAGE", "1");
		else if(frmScreen.optLandUsageNo.checked)
			XML.CreateTag("LANDUSAGE", "0");
		else
			XML.CreateTag("LANDUSAGE", "");

		<% /* KRW BMIDS776 - KFI Received by all Apps */%>	
		if(frmScreen.optKFIReceivedYes.checked)
			XML.CreateTag("KFIRECEIVEDBYALLAPPS", "1");
		else if(frmScreen.optKFIReceivedNo.checked)
			XML.CreateTag("KFIRECEIVEDBYALLAPPS", "0");
		else
			XML.CreateTag("KFIRECEIVEDBYALLAPPS", "");
	

		<% /* MV - 25/05/2002 - BMO44 - Amend BM Declaration */%> 
		if(frmScreen.optDeclarationYes.checked)
			XML.CreateTag("BMDECLARATIONIND", "1");
		else if(frmScreen.optDeclarationNo.checked)
			XML.CreateTag("BMDECLARATIONIND", "0");
		else
			XML.CreateTag("BMDECLARATIONIND", "");
			
		XML.CreateTag("ADDITIONALINTERMEDIARYDETAILS", frmScreen.txtAdditionalDetails.value);
		XML.CreateTag("MARKETINGSOURCE", frmScreen.cboMarketingSource.value);
		XML.CreateTag("SPECIALSCHEME", frmScreen.cboSpecialScheme.value);
		XML.CreateTag("NATUREOFLOAN", frmScreen.cboNatureOfLoan.value);
	
		XML.CreateTag("FIRMFSANUMBERLEVEL1", strIntroducerFSANumberLevel1); // KRW
		XML.CreateTag("FIRMFSANUMBERLEVEL2", strIntroducerFSANumberLevel2);
		XML.CreateTag("FIRMFSANUMBERLEVEL3", strIntroducerFSANumberLevel3);
	
		XML.CreateTag("INTRODUCERCORRESPONDINDLEVEL1", strIntroducerCorrespondIndLevel1);
		XML.CreateTag("INTRODUCERCORRESPONDINDLEVEL2", strIntroducerCorrespondIndLevel2);
		XML.CreateTag("INTRODUCERCORRESPONDINDLEVEL3", strIntroducerCorrespondIndLevel3);
		XML.CreateTag("ESTIMATEDCOMPLETIONDATE", frmScreen.txtEstimatedCompletionDate.value); // KRW
		XML.CreateTag("APPLICATIONRECEIVEDDATE", frmScreen.txtApplicationDateReceived.value); // KRW
		XML.CreateTag("APPLICATIONSIGNEDDATE", frmScreen.txtApplicationDateSigned.value); // EP8 pct 21/03/2006
		//XML.CreateTag("ADDITIONALBROKERFEE", frmScreen.txtAdditionalBrokerFee.value); // KRW // EP15 pct
		//XML.CreateTag("ADDITIONALBROKERFEEDESC", frmScreen.txtAdditionalBrokerFeeDesc.value); // KRW  // EP15 pct
		XML.CreateTag("TOTALPROCURATIONFEE", frmScreen.txtProcurationFee.value); // EP15 pct
		//XML.CreateTag("PACKAGERNETWORKNAME", frmScreen.txtPackagerNetwork.value); // EP15 pct
		//XML.CreateTag("INGESTIONAPPLICATIONNUMBER", frmScreen.txtPackagerApplicationRef.value); // EP15 pct
		<% /*  SR EP2_1272
		//XML.CreateTag("ADDITIONALBROKERFEE", frmScreen.txtBrokerFee.value);  // EP485 
		//XML.CreateTag("BROKERPROCFEEREFUND", frmScreen.txtBrokerFeeRefund.value);  // EP485 
		//XML.CreateTag("PACKAGEDVALUATIONFEE", frmScreen.txtPackagerValFee.value);  // EP485 
		//XML.CreateTag("PACKAGEDVALUATIONFEEREBATE", frmScreen.txtPackagerValFeeRebate.value);  // EP485 
		*/ %>
		<% /* EP2_701 */%> 
		XML.CreateTag("PRODUCTSCHEME", frmScreen.cboCreditScheme.value);
		
		<% /* BMIDS744 */%> 
		if(frmScreen.idCreditCheckOptOutYes.checked)
			XML.CreateTag("OPTOUTINDICATOR", "1");
		else if(frmScreen.idCreditCheckOptOutNo.checked)
			XML.CreateTag("OPTOUTINDICATOR", "0");
		else
			<% /* Should never get here, but default it to yes */%> 
			<% /* MF MARS default to no */%>
			XML.CreateTag("OPTOUTINDICATOR", "0");

		XML.CreateTag("EXTERNALSYSTEMTPDDECLARATION", m_sTPDDeclaration);
		<% /* EP2_645 */%> 
		XML.CreateTag("APPLICATIONINCOMESTATUS", frmScreen.cboAppIncomeStatus.value);
		XML.CreateTag("SELFCERTREASON", frmScreen.cboSelfCertReason.value);
		XML.CreateTag("SELFCERTREASONDETAILS", frmScreen.txtSelfCertDetails.value);
	
		// EP558: add CRITICALDATACONTEXT LH 17/05/2006
		XML.SelectTag(null,"REQUEST");
		XML.CreateActiveTag("CRITICALDATACONTEXT");
		XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		XML.SetAttribute("SOURCEAPPLICATION","Omiga");
		XML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
		XML.SetAttribute("ACTIVITYINSTANCE","1");
		XML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
		XML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
		XML.SetAttribute("COMPONENT","omAPP.ApplicationBO");
		XML.SetAttribute("METHOD","UpdateApplication");
				
		window.status = "Critical Data Check - please wait";
				
		// 	XML.RunASP(document,"WriteApplicationData.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
				XML.RunASP(document,"OmigaTMBO.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}
		
		window.status = "";
		
		bSuccess = XML.IsResponseOK();
		XML = null;
	}

	return(bSuccess);
}

// EP2_705 - No CD Check if Quick Quote.
// Copy of WriteApplicationSourceDetails method minus Critical Data Checks.
function WriteApplicationSourceDetailsNoCDCheck()
{
	var bSuccess = true;
	
	if(!(scScreenFunctions.IsValidationType(frmScreen.cboSpecialScheme, "BTL")))
	{
		if((m_sFamilyMember != "") || (m_sFamilyLet != ""))
		{
			bSuccess = DeletePropertyData();
			if (bSuccess == false)
			{
				alert("Failed to delete Family Let information.");
			}
		}
	}
		
	if (bSuccess)
	{
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window, null)
		
		AppXML.SelectTag(null, "APPLICATION");
		var tagAPPLICATION = XML.CreateActiveTag("APPLICATION");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("PACKAGENUMBER", AppXML.GetTagText("PACKAGENUMBER"));
		XML.CreateTag("NUMBEROFDEPENDANTS", frmScreen.txtNoOfDependents.value) 

		if(frmScreen.optLegalRepYes.checked)
			XML.CreateTag("ISLEGALREPTOBEUSED", "1");
		else if(frmScreen.optLegalRepNo.checked)
			XML.CreateTag("ISLEGALREPTOBEUSED", "0");
		else
			XML.CreateTag("ISLEGALREPTOBEUSED", "");
		XML.CreateTag("CHANNELID", frmScreen.cboChannel.value); <%/*  SR EP2_1335 */%>
		<% /* SR 01/03/2007 : EP2_1196 */ %>
		XML.CreateTag("PROCFEETOCUST", frmScreen.cboProcFeePassedToCustomer.value);
		XML.CreateTag("PROCFEETOCUSTAMOUNT", frmScreen.txtAmtProcFeePassedOn.value);
		<% /* SR 01/03/2007 : EP2_1196 - End */ %>
			
		var tagAPPLICATIONFACTFIND = XML.CreateActiveTag("APPLICATIONFACTFIND");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("LEVELOFADVICE", frmScreen.cboLevelOfAdvice.value); 
		XML.CreateTag("REGULATIONINDICATOR", frmScreen.cboRegulationIndicator.value);
		XML.CreateTag("DIRECTINDIRECTBUSINESS", frmScreen.cboDirectIndirect.value);
	
		if(frmScreen.optApplPackagedFullyYes.checked)
			XML.CreateTag("APPLICATIONPACKAGEINDICATOR", "1");
		else if(frmScreen.optApplPackagedFullyNo.checked)
			XML.CreateTag("APPLICATIONPACKAGEINDICATOR", "0");
		else
			XML.CreateTag("APPLICATIONPACKAGEINDICATOR", "");

		if(frmScreen.optLandUsageYes.checked)
			XML.CreateTag("LANDUSAGE", "1");
		else if(frmScreen.optLandUsageNo.checked)
			XML.CreateTag("LANDUSAGE", "0");
		else
			XML.CreateTag("LANDUSAGE", "");

		if(frmScreen.optKFIReceivedYes.checked)
			XML.CreateTag("KFIRECEIVEDBYALLAPPS", "1");
		else if(frmScreen.optKFIReceivedNo.checked)
			XML.CreateTag("KFIRECEIVEDBYALLAPPS", "0");
		else
			XML.CreateTag("KFIRECEIVEDBYALLAPPS", "");
	

		if(frmScreen.optDeclarationYes.checked)
			XML.CreateTag("BMDECLARATIONIND", "1");
		else if(frmScreen.optDeclarationNo.checked)
			XML.CreateTag("BMDECLARATIONIND", "0");
		else
			XML.CreateTag("BMDECLARATIONIND", "");
			
		XML.CreateTag("ADDITIONALINTERMEDIARYDETAILS", frmScreen.txtAdditionalDetails.value);
		XML.CreateTag("MARKETINGSOURCE", frmScreen.cboMarketingSource.value);
		XML.CreateTag("SPECIALSCHEME", frmScreen.cboSpecialScheme.value);
		XML.CreateTag("NATUREOFLOAN", frmScreen.cboNatureOfLoan.value);
	
		XML.CreateTag("FIRMFSANUMBERLEVEL1", strIntroducerFSANumberLevel1); 
		XML.CreateTag("FIRMFSANUMBERLEVEL2", strIntroducerFSANumberLevel2);
		XML.CreateTag("FIRMFSANUMBERLEVEL3", strIntroducerFSANumberLevel3);
	
		XML.CreateTag("INTRODUCERCORRESPONDINDLEVEL1", strIntroducerCorrespondIndLevel1);
		XML.CreateTag("INTRODUCERCORRESPONDINDLEVEL2", strIntroducerCorrespondIndLevel2);
		XML.CreateTag("INTRODUCERCORRESPONDINDLEVEL3", strIntroducerCorrespondIndLevel3);
		XML.CreateTag("ESTIMATEDCOMPLETIONDATE", frmScreen.txtEstimatedCompletionDate.value); 
		XML.CreateTag("APPLICATIONRECEIVEDDATE", frmScreen.txtApplicationDateReceived.value); 
		XML.CreateTag("APPLICATIONSIGNEDDATE", frmScreen.txtApplicationDateSigned.value); 
		XML.CreateTag("TOTALPROCURATIONFEE", frmScreen.txtProcurationFee.value); 
		<% /*   SR EP2_1272
		XML.CreateTag("ADDITIONALBROKERFEE", frmScreen.txtBrokerFee.value);  
		XML.CreateTag("BROKERPROCFEEREFUND", frmScreen.txtBrokerFeeRefund.value);
		XML.CreateTag("PACKAGEDVALUATIONFEE", frmScreen.txtPackagerValFee.value);
		XML.CreateTag("PACKAGEDVALUATIONFEEREBATE", frmScreen.txtPackagerValFeeRebate.value);  
		*/ %>
		XML.CreateTag("PRODUCTSCHEME", frmScreen.cboCreditScheme.value); 
		
		if(frmScreen.idCreditCheckOptOutYes.checked)
			XML.CreateTag("OPTOUTINDICATOR", "1");
		else if(frmScreen.idCreditCheckOptOutNo.checked)
			XML.CreateTag("OPTOUTINDICATOR", "0");
		else
			XML.CreateTag("OPTOUTINDICATOR", "0");

		XML.CreateTag("EXTERNALSYSTEMTPDDECLARATION", m_sTPDDeclaration);
		XML.CreateTag("APPLICATIONINCOMESTATUS", frmScreen.cboAppIncomeStatus.value);
		XML.CreateTag("SELFCERTREASON", frmScreen.cboSelfCertReason.value);
		XML.CreateTag("SELFCERTREASONDETAILS", frmScreen.txtSelfCertDetails.value);
	
		switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"UpdateApplication.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

		bSuccess = XML.IsResponseOK();
		XML = null;
	}

	return(bSuccess);
}

<% /* BMIDS00502 - BM061 - Introducer details */ %>
function frmScreen.cboDirectIndirect.onchange()
{
	//EP2_12  Not Used ?????
	<% /* Only display if a directindirect has already been selected */ %>
	//if (strDirectIndirectBusiness > "") {
		
		<% /* Are there any Introducers? */ %>
	//	if (strIntroducerIdLevel1.length > 0) {
		
	//		var bResult = confirm("If the Direct/Indirect indicator is changed, all Introducer details will have to be re-searched");
		
	//		if (bResult == true) {
	//			strDirectIndirectBusiness = frmScreen.cboDirectIndirect.value;
	//			ResetIntroducerList();
	//		} else {
				<% /* Set it back to the original value */ %>
	//			frmScreen.cboDirectIndirect.value = strDirectIndirectBusiness;
	//		}
	//	}
	//}
}

<% /* KRW - Special Scheme */ %>
function frmScreen.cboSpecialScheme.onchange()
{
	var bSuccess = false;
	
	if(scScreenFunctions.IsValidationType(frmScreen.cboSpecialScheme, "BTL")) <% /* SR 19/07/2004 : BMIDS802 */ %>
	{
	  // frmScreen.btnSpecialScheme.disabled = false;
	  scScreenFunctions.EnableDrillDown(frmScreen.btnSpecialScheme);	
	}
	else
	{
		//frmScreen.btnSpecialScheme.disabled = true;
		scScreenFunctions.DisableDrillDown(frmScreen.btnSpecialScheme);
	}

	<% /* INR - BMIDS776 Check whether the Regulation indicator needs to be reset */ %>
	SetRegulationIndicator();		

}

<% /* EP15 - Introducer details */ %>
function divIntroducerDetailsTable.onclick()
{
	if (scScrollTable.getRowSelected() != -1) frmScreen.btnDetail.disabled = false;
	else frmScreen.btnDetail.disabled = true;
	var blnDeleteDisabled = true;
	<% /* EP2_12 - allow delete on any row */ %>
	if((scScrollTable.getRowSelected() != -1)&(m_blnReadOnly == false))
		blnDeleteDisabled = false;
	frmScreen.btnDelete.disabled = blnDeleteDisabled;
	<% /* EP586 End */ %>
}

<% /* KRW 28/05/2004 : BBG82 */ %>
function frmScreen.optLandUsageYes.onclick()
{
	<% /* BMIDS776 m_sCurrLandUsage = 1;  */ %>
	SetRegulationIndicator(); 
	
//	frmScreen.cboRegulationIndicator.value = m_sNonRegulatedId ;
}

function frmScreen.optLandUsageNo.onclick()
{
	<% /* BMIDS776 m_sCurrLandUsage = 0;  */ %>
	SetRegulationIndicator(); 
	
//	frmScreen.cboRegulationIndicator.value = m_sNonRegulatedId ;
}

function SetRegulationIndicator()
{

	if(strTypeOfApplication == m_sSecuredPersonalLoan)
	{
		frmScreen.cboRegulationIndicator.value = m_sCCARegulatedId;
	}
	else
	{
		<% /* BMIDS776 INR
	   if(m_sCurrLandUsage == "1")	 */ %>
	   if (frmScreen.optLandUsageYes.checked)
	   {	
			 if(scScreenFunctions.IsValidationType(frmScreen.cboSpecialScheme, "BTL"))	<% /* SR 19/07/2004 : BMIDS802 */ %>	  	
		     {
	 	  		 <% /* BMIDS776 INR Only go in here if selected, DC013 value may be NULL */ %>
	 	  		 if(m_sFamilyLet == 1 )
				 {	  
				     if(m_sFamilyMemberValType != "O")
				     {
					    	frmScreen.cboRegulationIndicator.value = m_sRegulatedId;
 				     }
 				     else
				     {
				            frmScreen.cboRegulationIndicator.value = m_sNonRegulatedId;
				     }
			     }
			     else
			     {
			    		frmScreen.cboRegulationIndicator.value = m_sNonRegulatedId;
			     }
		     }
		     else
		     {
			     frmScreen.cboRegulationIndicator.value = m_sRegulatedId;
		     }
		}     
	    else
        { // LandUsage = No
			<% /* BMIDS776 INR
			if(m_sCurrLandUsage == "0")  */ %>
			if (frmScreen.optLandUsageNo.checked)
			{
				frmScreen.cboRegulationIndicator.value = m_sNonRegulatedId;
			}
			else
			{
				frmScreen.cboRegulationIndicator.value = "";
			}
	    }
	  
	}
}

<% /* BMIDS00502 - BM061 - Introducer details */ %>
function frmScreen.btnAdd.onclick()
{
	<% /* SR 22/02/2007 : EP2_1272 : saving the current data in to context - moved to diff function saveCurrentDataInContext() */ %>
	saveCurrentDataInContext();
	frmToDC015.submit();
}

function frmScreen.btnIntroducerFees.onclick()
{
	saveCurrentDataInContext();
	if(IsChanged()) 
	{
		scScreenFunctions.SetContextParameter(window,"idXML2", "<SCREENDIRTY>1</SCREENDIRTY>");
	}
	frmToDC016.submit();
}

<% /* SR 22/02/2007 : EP2_1272 : new function saveCurrentDataInContext() */ %>
function saveCurrentDataInContext()
{
	<% /* save the combo settings in idXML context to use when we return
	<% /* next line replaced by line below - DPF 21/06/02 - BMIDS00077 */ %>
	//var XML = new scXMLFunctions.XMLObject(); */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject(); 
	var strAppPackagedInd, strDeclaraionInd, strOptOutIndicator, strOptOutIndicator,strOptLandUsageIndicator, strOptKFIIndicator;
	/*DS - EP2_966*/
	var strOptLegalRepUsed;
	/*DS - EP2_966*/
	
	XML.CreateActiveTag("DC010SETTINGS");
	XML.CreateTag("DIRECTINDIRECTBUSINESS", frmScreen.cboDirectIndirect.value);
	XML.CreateTag("LEVELOFADVICE", frmScreen.cboLevelOfAdvice.value);
	XML.CreateTag("REGULATIONINDICATOR", frmScreen.cboRegulationIndicator.value);
	
	XML.CreateTag("DIRECTINDIRECTBUSINESSVALTYPE", scScreenFunctions.GetComboValidationType(frmScreen, frmScreen.cboDirectIndirect.id));
	
	XML.CreateTag("MARKETINGSOURCE", frmScreen.cboMarketingSource.value);
	XML.CreateTag("SPECIALSCHEME", frmScreen.cboSpecialScheme.value );
	XML.CreateTag("NATUREOFLOAN", frmScreen.cboNatureOfLoan.value);
	if (frmScreen.optApplPackagedFullyYes.checked == true) {
		strAppPackagedInd = "1"
	} else {
		if (frmScreen.optApplPackagedFullyYes.checked == true) {
			strAppPackagedInd = "0"
		} else {
			strAppPackagedInd = ""
		}
	}
	XML.CreateTag("APPLICATIONPACKAGEINDICATOR", strAppPackagedInd);

	<% /* IK EP371 self-certified indicator */ %>	
	<%/*MAH EP2_101 Removed SelfCert
	var strSelfCert = "";
	if (frmScreen.optSelfCertYes.checked == true) strSelfCert = "1";
	else if (frmScreen.optSelfCertNo.checked == true) strSelfCert = "0";
	XML.CreateTag("SELFCERTIND", strSelfCert);
	*/%>
	if (frmScreen.optDeclarationYes.checked == true) {
		strDeclaraionInd = "1"
	} else {
		if (frmScreen.optDeclarationNo.checked == true) {
			strDeclaraionInd = "0"
		} else {
			strDeclaraionInd = ""
		}
	}
	XML.CreateTag("BMDECLARATIONIND", strDeclaraionInd);
	
	XML.CreateTag("ADDITIONALINTERMEDIARYDETAILS", frmScreen.txtAdditionalDetails.value);
	
	XML.CreateTag("INTRODUCERCORRESPONDINDLEVEL1", strIntroducerCorrespondIndLevel1);
	XML.CreateTag("INTRODUCERCORRESPONDINDLEVEL2", strIntroducerCorrespondIndLevel2);
	XML.CreateTag("INTRODUCERCORRESPONDINDLEVEL3", strIntroducerCorrespondIndLevel3);

	<% /* BMIDS744 Save state of OptOutIndicator if going into DC015*/ %>
	if (frmScreen.idCreditCheckOptOutYes.checked == true) {
		strOptOutIndicator = "1"
	} else {
		if (frmScreen.idCreditCheckOptOutNo.checked == true) {
			strOptOutIndicator = "0"
		} else {
			<% /* Default OptOutIndicator is Yes*/ %>
			strOptOutIndicator = "1"
		}
	}
			
	if (frmScreen.optLandUsageYes.checked == true) {
		strOptLandUsageIndicator = "1"
	} else {
		if (frmScreen.optLandUsageNo.checked == true) {
			strOptLandUsageIndicator = "0"
		} else {
			<% /* Default OptOutIndicator is Yes*/ %>
			strOptLandUsageIndicator = ""
		}
	}
		
	if (frmScreen.optKFIReceivedYes.checked == true) {
		strOptKFIIndicator = "1"
	} else {
		if (frmScreen.optKFIReceivedNo.checked == true) {
			strOptKFIIndicator = "0"
		} else {
			<% /* Default OptOutIndicator is Yes*/ %>
			strOptKFIIndicator = ""
		}
	}

	/*DS - EP2_966 */
	
	if (frmScreen.optLegalRepYes.checked == true) {
		strOptLegalRepUsed = "1"
	} else {
		if (frmScreen.optLegalRepNo.checked == true) {
			strOptLegalRepUsed = "0"
		} else {
			<% /* Default OptOutIndicator is Yes*/ %>
			strOptLegalRepUsed = ""
		}
	}
	
	XML.CreateTag("ISLEGALREPTOBEUSED", strOptLegalRepUsed);
	/*DS - EP2_966 */

	XML.CreateTag("OPTOUTINDICATOR", strOptOutIndicator);	
	XML.CreateTag("LANDUSAGE", strOptLandUsageIndicator);
	XML.CreateTag("KFIRECEIVEDBYALLAPPS", strOptKFIIndicator);
	<% /* INR BMIDS776 txtApplicationDateReceived should be txtEstimatedCompletionDate */ %>
	XML.CreateTag("ESTIMATEDCOMPLETIONDATE", frmScreen.txtEstimatedCompletionDate.value);
	XML.CreateTag("APPLICATIONRECEIVEDDATE", frmScreen.txtApplicationDateReceived.value);
	XML.CreateTag("APPLICATIONSIGNEDDATE", frmScreen.txtApplicationDateSigned.value); //EP8 pct 21/03/2006
	//XML.CreateTag("ADDITIONALBROKERFEE", frmScreen.txtAdditionalBrokerFee.value); // EP15 pct
	//XML.CreateTag("ADDITIONALBROKERFEEDESC",frmScreen.txtAdditionalBrokerFeeDesc.value); // EP15 pct
	XML.CreateTag("TOTALPROCURATIONFEE", frmScreen.txtProcurationFee.value); // EP15 pct
	
	<% /* PSC 04/12/2006 EP2_249 - Start */ %>
	XML.CreateTag("ASSOCIATIONFEE", frmScreen.txtAssociationFee.value);
	XML.CreateTag("PACKAGINGFEE", frmScreen.txtPackagingFee.value);
	<% /* PSC 04/12/2006 EP2_249 - End */ %>
	
	<% /* SR 01/13/2007 EP2_1196 - Start */ %>
	XML.CreateTag("PROCFEETOCUST", frmScreen.cboProcFeePassedToCustomer.value);
	XML.CreateTag("PROCFEETOCUSTAMOUNT", frmScreen.txtAmtProcFeePassedOn.value);
	<% /* SR 01/13/2007 EP2_1196 - End */ %>
	XML.CreateTag("CHANNELID", frmScreen.cboChannel.value); <%/*  SR EP2_1335 */%>
	
	//XML.CreateTag("PACKAGERNETWORKNAME", frmScreen.txtPackagerNetwork.value); // EP15 pct
	XML.CreateTag("INGESTIONAPPLICATIONNUMBER", frmScreen.txtPackagerApplicationRef.value); // EP15 pct
	<% /*  SR EP2_1272
	XML.CreateTag("ADDITIONALBROKERFEE", frmScreen.txtBrokerFee.value);  // EP485
	XML.CreateTag("BROKERPROCFEEREFUND", frmScreen.txtBrokerFeeRefund.value);  // EP485 
	XML.CreateTag("PACKAGEDVALUATIONFEE", frmScreen.txtPackagerValFee.value);  // EP485
	XML.CreateTag("PACKAGEDVALUATIONFEEREBATE", frmScreen.txtPackagerValFeeRebate.value);  // EP485
	*/ %>	
	XML.CreateTag("EXTERNALSYSTEMTPDDECLARATION",m_sTPDDeclaration);
	XML.CreateTag("NUMBEROFDEPENDANTS", frmScreen.txtNoOfDependents.value);
	<% /* EP2_645 */ %>
	XML.CreateTag("APPLICATIONINCOMESTATUS", frmScreen.cboAppIncomeStatus.value);
	XML.CreateTag("SELFCERTREASON", frmScreen.cboSelfCertReason.value);
	XML.CreateTag("SELFCERTREASONDETAILS", frmScreen.txtSelfCertDetails.value);
	
	<% /* PSC 01/02/2007 EP2_1112 - Start */ %>
	if (m_bReturningCustomer)
		XML.CreateTag("RETURNINGCUSTOMER", "1");
	else
		XML.CreateTag("RETURNINGCUSTOMER", "0");
		
	XML.CreateTag("PRODUCTSCHEME", frmScreen.cboCreditScheme.value);
	<% /* PSC 01/02/2007 EP2_1112 - End */ %>
	scScreenFunctions.SetContextParameter(window,"idXML",XML.ActiveTag.xml);
}

<% /* KRW 09/06/2004 BMIDS776*/ %>
function frmScreen.btnSpecialScheme.onclick()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var saReturn = null;
	var sRegulatedIndicator = "";
	var ArrayArguments = new Array();
	
	ArrayArguments[0] = m_sApplicationNumber;
	ArrayArguments[1] = m_sNewPropertyExists;
	ArrayArguments[2] = m_sFamilyLet;
	ArrayArguments[3] = m_sFamilyMember;
	ArrayArguments[4] = (frmScreen.cboRegulationIndicator.value == m_sRegulatedId) ? "1" : "0" ;
	ArrayArguments[5] = scScreenFunctions.GetRadioGroupValue(frmScreen,"LandUsage");
	ArrayArguments[6] = m_sReadOnly;
	<% /* BMIDS887 GHun */ %>
	ArrayArguments[7] = frmScreen.cboRegulationIndicator.value;
	ArrayArguments[8] = frmScreen.cboSpecialScheme.value;
	ArrayArguments[9] = m_sApplicationFactFindNumber;
	ArrayArguments[10] = frmScreen.optLandUsageNo.checked ? "0" : "1";
	ArrayArguments[11] = XML.CreateRequestAttributeArray(window);
	<% /* BMIDS887 End */ %>
	
	<% /* MAR10 GHun Made popup window larger so scrollbars are not necessary*/ %>
	<% /* EP2_677 make dialog larger */ %>
	saReturn = scScreenFunctions.DisplayPopup(window, document, "DC013.asp", ArrayArguments, 428, 210);
	
	if (saReturn != null)  <% /* BMIDS766 KW 13/07/2004 */ %>
	<% /* Get the values saved in DC013 and populate them here */ %>
	{
		m_sFamilyLet			= saReturn[1];
		m_sFamilyMember			= saReturn[2];
		frmScreen.cboRegulationIndicator.value =  saReturn[3];
		m_sFamilyMemberValType = saReturn[4];
		m_sNewPropertyExists	= "1";
	}  <% /* SR 05/04/2004 : BBG121 - End */ %>
		
	XML = null;
	
	SetRegulationIndicator();

}

function frmScreen.btnDetail.onclick() 
{
	<% /* EP2_12  Pass through the selected info to DC011*/%>
	var ArrayArguments = new Array(2); 
	//var iRowSelected = scScrollTable.getRowSelected();  //-- SR EP2_115
	var iRowSelected = scScrollTable.getRowSelectedIndex()  //SR EP2_115
	//var sAppIntroducerSeqNo = tblIntroducerDetails.rows(iRowSelected).getAttribute("AppIntroducerSeqNo"); -- SR EP2_115
	var IntroducerDetailXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	IntroducerDetailXML.CreateRequestTag(window, null);
	//SR EP2_115 - choose the node from xml based on number of the row selected in the list box. 
	//APPLICATIONINTRODUCERSEQNO can be same for two rows.
	// var xe = IntroducerXML.XMLDocument.selectSingleNode("//INTERMEDIARY[@APPLICATIONINTRODUCERSEQNO='"+ sAppIntroducerSeqNo +"']");
	var xe = IntroducerXML.XMLDocument.selectNodes("//INTERMEDIARY")[iRowSelected-1];
	IntroducerDetailXML.XMLDocument.documentElement.appendChild(xe.cloneNode(true));

	ArrayArguments[0] = IntroducerDetailXML.XMLDocument.xml;

	scScreenFunctions.DisplayPopup(window, document, "dc011.asp", ArrayArguments, 420, 570);

	return;
}

<% /* EP15 - Introducer details */ %>
function frmScreen.btnDelete.onclick() 
{
	var nRowSelected =  scScrollTable.getRowSelected();
		
	if (nRowSelected != -1) 
	{

		<% //Delete selected APPLICATIONINTRODUCER %>
		var applicationIntroducerSeqNo = tblIntroducerDetails.rows(nRowSelected).getAttribute("AppIntroducerSeqNo");
	
		IntermedXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		IntermedXML.CreateRequestTag(window);
		var xn = IntermedXML.XMLDocument.documentElement;
		xn.setAttribute("CRUD_OP","DELETE");
		var xe = IntermedXML.XMLDocument.createElement("APPLICATIONINTRODUCER");
		xe.setAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
		xe.setAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		xe.setAttribute("APPLICATIONINTRODUCERSEQNO",applicationIntroducerSeqNo);
		xe.setAttribute('CRUD_OP', 'DELETE');
		xn.appendChild(xe);

		IntermedXML.RunASP(document, "omCRUDIf.asp");

		if(!IntermedXML.IsResponseOK()) return;

		scScrollTable.clear();

		GetIntroducers();
	}
}

function btnCancel.onclick()
{
	<% /* clear down any Context's */ %>
	if(m_sContext == "CompletenessCheck")
		frmToGN300.submit();
	else
	{
		scScreenFunctions.SetContextParameter(window,"idXML","");
		scScreenFunctions.SetContextParameter(window,"idXML2","");
		scScreenFunctions.SetContextParameter(window,"idMetaAction","");
		<% // SAB 03/04/2006 EP8 - Change to process flow %>
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		if (XML.GetGlobalParameterBoolean(document, "HideApplicationVerification"))
			frmToMN060.submit();
		else
		// EP2_8 - Check for QQ and route appropriately.
			if (m_sQuickQuote == "Quick Quote")
					frmToMQ010.submit();
			else
				frmToDC012.submit()

			
	}
}

function btnSubmit.onclick()
{
	var bSuccess = frmScreen.onsubmit();
	
	<% /* PSC 07/02/2007 EP2_373 */ %>
	if (m_sReadOnly != "1" && frmScreen.optKFIReceivedNo.checked==false && frmScreen.optKFIReceivedYes.checked==false)
	{
		alert('Please select either Yes or No for the field "KFI Received by all applicants?"');
		frmScreen.optKFIReceivedYes.focus();
		return false;
	}
	<% /* BMIDS828  Remove reference to btnIntermedSearch 
	if (!bSuccess) frmScreen.btnIntermedSearch.focus()  */ %>
	
	<% /* SYS0738 If either tag is null this screen has never been used before.
	      Therefore force an update */ %>
	      
	if (AppXML!=null)
		<% /* BMIDS744 Only make this call if bUpdateRequired is false, we may
		be already updating because of the TPDDecalartion screen and we don't
		want to change it to false*/ %>
	
		if (bUpdateRequired == false)
			bUpdateRequired = ( AppXML.GetTagText("DIRECTINDIRECTBUSINESS")=="" ||
						    AppXML.GetTagText("NATUREOFLOAN")=="");
	else
		bUpdateRequired=true;

	if (m_sReadOnly != "1" && bSuccess)
	{		 		
		<% /* SYS0855 Extra validation */ %>
		<% /* SYS3644 - Use Validation types rather than ValueIDs. */ %>
		if (scScreenFunctions.GetComboValidationType(frmScreen, frmScreen.cboDirectIndirect.id) == "D" &&
			(scScreenFunctions.GetComboValidationType(frmScreen, frmScreen.cboMarketingSource.id) == "F" ||
			 scScreenFunctions.GetComboValidationType(frmScreen, frmScreen.cboMarketingSource.id) == "I" ||
			 scScreenFunctions.GetComboValidationType(frmScreen, frmScreen.cboMarketingSource.id) == "O"))
		<% /* SYS3644 - End. */ %>
		{	
			alert("A Direct marketing source cannot be Referral, Intermediary or Other");
			frmScreen.cboMarketingSource.focus();
			return false;
		}
		else
		{
			if(frmScreen.txtAdditionalDetails.value.length > 500)
			{
				bSuccess = false;
				alert("The notes in Additional Details are too long. Please re-phrase.");
			}

			<% /* EP565 */ %>
			if (bSuccess && scScreenFunctions.CompareDateFieldToSystemDate(frmScreen.txtApplicationDateReceived,">"))
			{
				bSuccess = false;
				alert("The Application Received Date cannot a date in the future.");
				frmScreen.txtApplicationDateReceived.focus();
			}
			<% /* EP565 */ %>

			if (bSuccess && (IsChanged() || bUpdateRequired || m_sMetaAction == "LoadIntermediaryFromDC015" || m_sContext == "CompletenessCheck" || bIntroducersChanged == true))
			{
				// EP2_705 - No Critical Data Checks for QQ cases.
				if (m_sQuickQuote == "Quick Quote")
					bSuccess = WriteApplicationSourceDetailsNoCDCheck();
				
				else
					bSuccess = WriteApplicationSourceDetails();
			}
		}
	}

	if(bSuccess)
	{
		if(m_sContext == "CompletenessCheck")
			frmToGN300.submit();
		else
		{		
			scScreenFunctions.SetContextParameter(window,"idXML","");
			scScreenFunctions.SetContextParameter(window,"idXML2","");
			scScreenFunctions.SetContextParameter(window,"idMetaAction","");
			// EP2_8 - Check for QQ and route appropriately.
			if (m_sQuickQuote == "Quick Quote")
				frmToMQ010.submit();
			else
				frmToDC030.submit();
		}
	}
}

<% /* BMIDS744 New call to DC014 */ %>
function frmScreen.btnTPDDeclaration.onclick()
{
	var sReturn = null;
	var ArrayArguments = new Array();
	ArrayArguments[0] = m_sTPDDeclaration;
	ArrayArguments[1] = m_sReadOnly;

	sReturn = scScreenFunctions.DisplayPopup(window, document, "dc014.asp", ArrayArguments, 550, 460);
	if (sReturn != null)
	{
		<% /* Only change this if it is false */ %>
		if (bUpdateRequired == false)
			bUpdateRequired = sReturn[0];
		m_sTPDDeclaration = sReturn[1];
	}
}

<% /* INR BMIDS744 New Function  */ %>

function GetCreditCheckStatus()
{
	var bSuccess = false;
	var nNumCreditCheck;

	CreditCheckXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	with (CreditCheckXML)
	{
		<% /* Retrieve Credit Check status */ %>
		CreateRequestTag(window, null)

		<% /* Create request block */ %>
		CreateActiveTag("SEARCH");
		CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);

		<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					RunASP(document, "GetCreditCheckStatus.asp");
				break;
			default: // Error
				SetErrorResponse();
			}

		if (!IsResponseOK())
		{
			alert("ERROR : Call to ApplicationCreditCheck failed.");
		}
		else
		{
			SelectSingleNode("//APPLICATIONCREDITCHECK");
			if (ActiveTag != null) 
			{
				nNumCreditCheck = GetAttribute("SEQUENCENUMBER")
				if (nNumCreditCheck > 0) 
					bSuccess = true;
			}
		}
	}

	return bSuccess;
}

<% /* INR BMIDS776 New Function  */ %>
function DeletePropertyData()
{
	var bSuccess = true;
	m_sFamilyMember = "";
	m_sFamilyMemberValType = "";
	m_sFamilyLet = "";
	
	var DelXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	DelXML.CreateRequestTag(window, null)

	DelXML.CreateActiveTag("NEWPROPERTY");
	DelXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	DelXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	<% /* EP2_677 */ %>
	DelXML.CreateTag("OCCUPYPROPERTY", m_sFamilyLet);
	DelXML.CreateTag("FAMILYMEMBER", m_sFamilyMember);
		
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			DelXML.RunASP(document,"UpdateNewPropertyGeneral.asp");
					
			break;
		default: // Error
			DelXML.SetErrorResponse();
	}
	
	bSuccess = DelXML.IsResponseOK();
	DelXML = null;
	
	return(bSuccess);
}

<% /* EP695 - Peter Edney - 09/06/2006 */ %>
<% /* Round down values */ %>
<% /* //SR Ep2_1272 function frmScreen.txtBrokerFee.onblur()
{
	frmScreen.txtBrokerFee.value =  Math.floor(frmScreen.txtBrokerFee.value);
}
function frmScreen.txtBrokerFeeRefund.onblur()
{
	frmScreen.txtBrokerFeeRefund.value =  Math.floor(frmScreen.txtBrokerFeeRefund.value);
}
function frmScreen.txtPackagerValFee.onblur()
{
	frmScreen.txtPackagerValFee.value =  Math.floor(frmScreen.txtPackagerValFee.value);
}
function frmScreen.txtPackagerValFeeRebate.onblur()
{
	frmScreen.txtPackagerValFeeRebate.value =  Math.floor(frmScreen.txtPackagerValFeeRebate.value);
}
/* %>
<% /* EP2_645 */ %>
function frmScreen.cboAppIncomeStatus.onchange()
{
	
	if(scScreenFunctions.IsValidationType(frmScreen.cboAppIncomeStatus, "SC")) 
	{
		<% /* Change to Self Cert, enable Self-cert Reason */ %>
		frmScreen.cboSelfCertReason.disabled = false;
	}
	else
	{
		<% /* Change from Self Cert, disable Self-cert Reason  & details*/ %>
		frmScreen.cboSelfCertReason.disabled = true;
		frmScreen.txtSelfCertDetails.setAttribute("required", "false");
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtSelfCertDetails");
		frmScreen.txtSelfCertDetails.value = "";
		frmScreen.cboSelfCertReason.value = "";
	}

}
<% /* EP2_645 */ %>
function frmScreen.cboSelfCertReason.onchange()
{
	
	if(scScreenFunctions.IsValidationType(frmScreen.cboSelfCertReason, "O"))
	{
		<% /* Change to Reason of Other */ %>
		scScreenFunctions.SetFieldState(frmScreen, "txtSelfCertDetails", "W");
		frmScreen.txtSelfCertDetails.setAttribute("required", "true");
	}
	else
	{
		<% /* Change to Reason other than Other */ %>
		frmScreen.txtSelfCertDetails.setAttribute("required", "false");
		frmScreen.txtSelfCertDetails.value = "";
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtSelfCertDetails");
	}

}

<% /* SR 01/03/2007 : EP2_1196 */ %>
function frmScreen.cboProcFeePassedToCustomer.onchange()
{
	setStateOnProcFeeAmtPassOnToCust();
}
function setStateOnProcFeeAmtPassOnToCust()
{
	if(scScreenFunctions.IsValidationType(frmScreen.cboProcFeePassedToCustomer, 'S'))
	{
		scScreenFunctions.SetCollectionToWritable(spnProcFeeAmtPassOn);
	}
	else 
	{
		scScreenFunctions.SetCollectionToReadOnly(spnProcFeeAmtPassOn);
		frmScreen.txtAmtProcFeePassedOn.value = "" ;
	}
}
<% /* SR 01/03/2007 : EP2_1196 - End */ %>
-->
		</script>
	</body>
</html>
