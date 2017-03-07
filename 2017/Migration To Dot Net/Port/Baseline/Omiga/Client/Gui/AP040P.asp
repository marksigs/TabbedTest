<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP040P.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Application Summary Popup - Application
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
LH		22/09/06	Created
LH		28/09/06	Modified - EP1180: Address' now display the street name
LH		11/10/06	Modified - EP1209: Address table now populating all addresses,
					made a necessary change also to the employment scroll table,
					added table row length constants.
LH		11/10/06	Modified - EP1210: 'Special Scheme' label changed to 'Product Category'.
PSC		16/03/07	EP2_1726 - Amend display of introducer data
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>AP040P Application Summary <!-- #include file="includes/TitleWhiteSpace.asp" -->       </title>
</head>

<body>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<span style="LEFT: 300px; POSITION: absolute; TOP: 126px">
<OBJECT id=scScrollTableApplicants 
style="LEFT: 0px; WIDTH: 304px; TOP: 0px; HEIGHT: 24px" tabIndex=-1 
type=text/x-scriptlet data=scTableListScroll.asp 
VIEWASTEXT></OBJECT>
</span>

<span style="LEFT: 300px; POSITION: absolute; TOP: 232px">
<OBJECT id=scScrollTableEmployment 
style="LEFT: 0px; WIDTH: 304px; TOP: 0px; HEIGHT: 24px" tabIndex=-1 
type=text/x-scriptlet data=scTableListScroll.asp 
VIEWASTEXT></OBJECT>
</span>

<span style="LEFT: 300px; POSITION: absolute; TOP: 338px">
<OBJECT id=scScrollTableOccupancy 
style="LEFT: 0px; WIDTH: 304px; TOP: 0px; HEIGHT: 24px" tabIndex=-1 
type=text/x-scriptlet data=scTableListScroll.asp VIEWASTEXT></OBJECT>
</span><!-- Hide this scroll bar for now -->
<span style="LEFT: 300px;  POSITION: absolute; TOP: 507px">
<OBJECT id=scScrollTableIntroducer 
style="LEFT: 0px; WIDTH: 304px; TOP: 0px; HEIGHT: 24px" tabIndex=-1 
type=text/x-scriptlet data=scTableListScroll.asp VIEWASTEXT></OBJECT>
</span><!-- Hide this scroll bar for now -->
<span style="LEFT: 300px; VISIBILITY: hidden; POSITION: absolute; TOP: 140px">
<OBJECT id=scScrollTableLoan 
style="LEFT: 0px; WIDTH: 304px; TOP: 0px; HEIGHT: 24px" tabIndex=-1 
type=text/x-scriptlet data=scTableListScroll.asp 
VIEWASTEXT></OBJECT>
</span>

<span style="LEFT: 300px; POSITION: absolute; TOP: 1194px">
<OBJECT id=scScrollTableThirdParty 
style="LEFT: 0px; WIDTH: 304px; TOP: 0px; HEIGHT: 24px" tabIndex=-1 
type=text/x-scriptlet data=scTableListScroll.asp VIEWASTEXT></OBJECT>
</span>


<% /* Specify Forms Here */ %>
<form id="frmToAP010" method="post" action="AP010.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP011" method="post" action="AP011P.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4>
	<div id="divApplicants" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 10px; HEIGHT: 235px" class="msgGroup">
		<span style="LEFT: 6px; POSITION: absolute; TOP: 10px" class="msgLabelHead">
			Applicants (select to view details):
		</span>
		
		<span id="spnTblApplicants" style="LEFT: 4px; POSITION: absolute; TOP: 25px">
			<table id="tblApplicants" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">
				<td width="15%" class="TableHead">Surname</td>
				<td width="17%" class="TableHead">Forenames</td>
				<td width="5%" class="TableHead">Title</td>
				<td width="17%" class="TableHead">Marital Status</td>
				<td width="13%" class="TableHead">DOB</td>
				<td width="10%" class="TableHead">Age</td>
				<td width="10%" class="TableHead">Gender</td>
				<td width="13%" class="TableHead">Role</td>
			</tr>
			<tr id="row01">
				<td class="TableTopLeft">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row02">
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row03">
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row04">
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row04">
				<td class="TableBottomLeft">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomRight">&nbsp;</td>
			</tr></table>
		</span>

		<span style="LEFT: 5px; POSITION: absolute; TOP: 123px" class="msgLabel">
			Contact Details
			<span style="LEFT: 90px; POSITION: absolute; TOP: -8px">
				<input id="btnContact" type ="button" style="LEFT: 0px; WIDTH: 26px; TOP: -1px; HEIGHT: 26px" class="msgDDButton">
			</span>
		</span>
			
		<span id="spnTblEmployment" style="LEFT: 4px; POSITION: absolute; TOP: 159px">
			<table id="tblEmployment" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">
				<td width="21%" class="TableHead">Employment Status</td>
				<td width="12%" class="TableHead">Occupation</td>
				<td width="12%" class="TableHead">Date Start</td>
				<td width="13%" class="TableHead">Date Ended</td>
				<td width="18%" class="TableHead">Time Employed</td>
				<td width="16%" class="TableHead">Gross Income</td>
				<td width="8%" class="TableHead">Main</td>
			</tr>
			<tr id="row01">
				<td class="TableTopLeft">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row02">
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row03">
				<td class="TableBottomLeft">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomRight">&nbsp;</td>
			</tr></table>
		</span>

		<span id="spnTblOccupancy" style="LEFT: 4px; POSITION: absolute; TOP: 265px">
			<table id="tblOccupancy" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">
				<td width="33%" class="TableHead">Address</td>
				<td width="15%" class="TableHead">Occupancy</td>
				<td width="16%" class="TableHead">Date Moved In</td>
				<td width="17%" class="TableHead">Date Moved Out</td>
				<td width="19%" class="TableHead">Time At Address</td>
			</tr>
			<tr id="row01">
				<td class="TableTopLeft">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row02">
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row03">
				<td class="TableBottomLeft">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomRight">&nbsp;</td>
			</tr></table>
		</span>

		<span style="LEFT: 7px; POSITION: absolute; TOP: 365px" class="msgLabelHead">
			Mortgage Property Address
		</span>

		<span style="LEFT: 6px; POSITION: absolute; TOP: 380px">
			<input id="txtMortgagePropertyAddress" maxlength="100" style="LEFT: 0px; WIDTH: 594px; TOP: 0px; HEIGHT: 20px" class="msgTxt" size=59>
		</span>


		<span style="LEFT: 6px; POSITION: absolute; TOP: 420px" class="msgLabelHead">
			Introducer Details
		</span>
		<span id="spnTblIntroducer" style="LEFT: 4px; POSITION: absolute; TOP: 435px">
			<table id="tblIntroducer" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
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
				<td width="16%" class="TableBottomLeft">&nbsp;</td>
				<td width="13%" class="TableBottomCenter">&nbsp;</td>
				<td width="12%" class="TableBottomCenter">&nbsp;</td>
				<td width="27%" class="TableBottomCenter">&nbsp;</td>
				<td width="8%" class="TableBottomCenter">&nbsp;</td>
				<td width="16%" class="TableBottomCenter">&nbsp;</td>
				<td width="8%" class="TableBottomRight">&nbsp;</td>
			</tr>
			</table>
		</span>
	</div>
	
	<div id="divAppDetails" style="LEFT: 220px; WIDTH: 297px; POSITION: absolute; TOP: 534px; HEIGHT: 310px" class="msgGroup">
		<span style="LEFT: 22px; POSITION: absolute; TOP: 10px" class="msgLabel" id=SPAN1>
			Packager Application Ref
			<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
				<input id="txtPackagerApplicationRef" maxlength="20" style="WIDTH: 132px; HEIGHT: 20px" class="msgTxt" size=31>
			</span>
		</span>
		<span style="LEFT: 330px; POSITION: absolute; TOP: 6px">
			<input id="btnDetail" value="Detail" type="button" style="LEFT: 0px; WIDTH: 60px; TOP: -1px" class="msgButton">
		</span>
	</div>

	<div id="divMortgageDetails" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 599px; HEIGHT: 235px" class="msgGroup">
		<span style="LEFT: 4px; POSITION: absolute; TOP: -13px" class="msgLabelHead" id=SPAN1>
			Mortgage Details
		</span>

		<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
			Total Loan Amount
			<span style="LEFT: 95px; POSITION: absolute; TOP: 0px">
				<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgTxt"></label>
				<input id="txtTotalLoanAmount" maxlength="10" style="LEFT: 1px; WIDTH: 67px; POSITION: absolute; TOP: -3px; HEIGHT: 20px" class="msgTxt" size=18>
			</span>
		</span>

		<span style="LEFT: 184px; POSITION: absolute; TOP: 10px" class="msgLabel">
			Amount Requested
			<span style="LEFT: 95px; POSITION: absolute; TOP: 0px">
				<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgTxt"></label>
				<input id="txtAmountRequested" maxlength="10" style="LEFT: 4px; WIDTH: 67px; POSITION: absolute; TOP: -3px; HEIGHT: 20px" class="msgTxt" size=18>
			</span>
		</span>

		<span style="LEFT: 368px; POSITION: absolute; TOP: 10px" class="msgLabel">
			Purchase Price
			<span style="LEFT: 95px; POSITION: absolute; TOP: 0px">
				<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgTxt"></label>
				<input id="txtPurchasePrice" maxlength="10" style="LEFT: -16px; WIDTH: 67px; POSITION: absolute; TOP: -3px; HEIGHT: 20px" class="msgTxt" size=18>
			</span>
		</span>

		<span style="LEFT: 539px; POSITION: absolute; TOP: 10px" class="msgLabel">
			LTV
			<span style="LEFT: 95px; POSITION: absolute; TOP: 0px">
				<label style="LEFT: -29px; POSITION: absolute; TOP: 0px" class="msgLabel" id=LABEL1>%</label>
				<input id="txtLTV" maxlength="3" style="LEFT: -69px; WIDTH: 37px; POSITION: absolute; TOP: -3px; HEIGHT: 20px" class="msgTxt" size=18>
			</span>
		</span>

	</div>

	<div id="divLoanDetails" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 644px; HEIGHT: 235px" class="msgGroup">
		<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabelHead">
			Loan Details
		</span>
		<span id="spnTblLoan" style="LEFT: 4px; POSITION: absolute; TOP: 30px">
			<table id="tblLoan" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">
				<td width="8%" class="TableHead">Prod. No</td>
				<td width="8%" class="TableHead">Prod. Type</td>
				<td width="12%" class="TableHead">Product Description</td>
				<td width="8%" class="TableHead">Rate</td>
				<td width="8%" class="TableHead">Manual Adj %</td>
				<td width="10%" class="TableHead">Loan Amount</td>
				<td width="8%" class="TableHead">Term</td>
				<td width="8%" class="TableHead">Repay Type</td>
				<td width="8%" class="TableHead">APR</td>
				<td width="10%" class="TableHead">Monthly Cost</td>
				<td width="12%" class="TableHead">Final Rate Cost</td>
			</tr>
			<tr id="row01">
				<td class="TableTopLeft">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row02">
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row03">
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row04">
				<td class="TableBottomLeft">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomRight">&nbsp;</td>
			</tr></table>
		</span>
		
		<!--<span style="LEFT: 399px; WIDTH: 141px; POSITION: absolute; TOP: 134px; HEIGHT: 14px" class="msgLabel">
			Total Cost
			<span style="LEFT: 140px; POSITION: absolute; TOP: 0px">
				<label style="LEFT: 11px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
				<input id="txtTotalCost" maxlength="10" style="LEFT: -55px; WIDTH: 100px; POSITION: absolute; TOP: -3px" class="msgTxt">
			</span>
		</span>-->
		
		<span style="LEFT: 400px; WIDTH: 141px; POSITION: absolute; TOP: 128px; HEIGHT: 14px" class="msgLabel">
			Total Cost
			<span style="LEFT: 60px; POSITION: absolute; TOP: 0px">
				<label style="LEFT: 11px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
				<input id="txtTotalCost" maxlength="10" style="LEFT: 20px; WIDTH: 100px; POSITION: absolute; TOP: -3px" class="msgTxt">
			</span>
		</span>
			
	</div>

	<div id="divAppDetails1" style="LEFT: 10px; WIDTH: 297px; POSITION: absolute; TOP: 783px; HEIGHT: 310px" class="msgGroup">
		<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabelHead">
			Application Details
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 35px" class="msgLabel">
			Product Category
			<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<input id="txtProductCategory" maxlength="8" style="WIDTH: 100px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 60px" class="msgLabel">
			Type of Buyer
			<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<input id="txtTypeOfBuyer" maxlength="50" style="WIDTH: 100px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 85px" class="msgLabel">
			Declared Income Multiple 
			<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<input id="txtDeclaredIncomeMultiple" maxlength="40" style="WIDTH: 100px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 110px" class="msgLabel">
			Confirmed Income Multiple
			<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<input id="txtConfirmedIncomeMultiple" maxlength="40" style="LEFT: 50px; WIDTH: 100px; TOP: 1px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 135px" class="msgLabel">
			Income Status
			<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<input id="txtIncomeStatus" maxlength="50" style="WIDTH: 100px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 160px" class="msgLabel">
			DIP Date
			<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<input id="txtDIPDate" maxlength="10" style="WIDTH: 60px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 185px" class="msgLabel">
			Application Date Received
			<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<input id="txtApplicationDate" maxlength="10" style="WIDTH: 60px" class="msgTxt">
			</span>
		</span>
	
		<span style="LEFT: 4px; VISIBILITY: hidden; POSITION: absolute; TOP: 210px" class="msgLabel">
			Level of Advice
			<span style="LEFT: 140px; VISIBILITY: hidden; POSITION: absolute; TOP: -3px">
				<input id="txtLevelOfAdvice" maxlength="50" style="WIDTH: 100px" class="msgTxt">
			</span>
		</span>
	</div>


	<div id="divAppDetails2" style="LEFT: 310px; WIDTH: 297px; POSITION: absolute; TOP: 809px; HEIGHT: 310px" class="msgGroup">

		<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
			Present Valuation
			<span style="LEFT: 140px; POSITION: absolute; TOP: 0px">
				<label style="LEFT: 11px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
				<input id="txtPresentValue" maxlength="10" style="LEFT: 20px; WIDTH: 100px; POSITION: absolute; TOP: -3px" class="msgTxt">
			</span>
		</span>
		
		<span style="LEFT: 4px; POSITION: absolute; TOP: 35px" class="msgLabel">
			Post Works Valuation
			<span style="LEFT: 140px; POSITION: absolute; TOP: 0px">
				<label style="LEFT: 11px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
				<input id="txtPostValue" maxlength="10" style="LEFT: 20px; WIDTH: 100px; POSITION: absolute; TOP: -3px" class="msgTxt">
			</span>
		</span>

		<span style="LEFT: 4px; POSITION: absolute; TOP: 60px" class="msgLabel">
			Retention Amount
			<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<label style="LEFT: 11px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
				<input id="txtRetentionAmount" maxlength="12" style="LEFT: 20px; WIDTH: 100px; POSITION: absolute; TOP: -3px" class="msgCurrency">
			</span>
		</span>
		
		<span style="LEFT: 4px; POSITION: absolute; TOP: 85px" class="msgLabel">
			Property val to be used in LTV
			<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<label style="LEFT: 10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
				<input id="txtPropertyValInLTV" maxlength="10" style="LEFT: 20px; WIDTH: 100px; POSITION: absolute; TOP: -3px" class="msgTxt">
			</span>
		</span>


		<span style="LEFT: 4px; POSITION: absolute; TOP: 110px" class="msgLabel">
			Application Status
			<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<input id="txtApplicationStatus" maxlength="20" style="LEFT: 20px; WIDTH: 80px; POSITION: absolute; TOP: -3px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 250px; POSITION: absolute; TOP: 102px">
			<input id="btnReview" value="Review" type="button" style="LEFT: 1px; VISIBILITY: hidden; WIDTH: 60px; TOP: 0px" class="msgButton">
		</span><!--Span style="LEFT: 4px; POSITION: absolute; TOP: 135px" class="msgLabel" style="VISIBILITY: hidden;">
			Application Priority
			<span style="LEFT: 140px; POSITION: absolute; TOP: -3px; VISIBILITY: hidden">
				<select id="cboApplicationPriority" style="LEFT: 20px; WIDTH: 100px; POSITION: absolute; TOP: -3px" class="msgTxt"></select>
			</span>
		</span>-->
		<span style="LEFT: 4px; VISIBILITY: hidden; POSITION: absolute; TOP: 160px" class="msgLabel">
			Regulation Indicator
			<span style="LEFT: 140px; VISIBILITY: hidden; POSITION: absolute; TOP: -3px">
				<input id="txtRegulationIndicator" maxlength="50" style="LEFT: 20px; WIDTH: 100px; POSITION: absolute; TOP: -3px" class="msgTxt">
			</span>
		</span>
		<span id="lblCreditCheckOptOutIndicator" name="lblCreditCheckOptOutIndicator" style="LEFT: 5px; VISIBILITY: hidden; POSITION: absolute; TOP: 185px" class="msgLabel">
			Opted out?
		</span>
		<span style="LEFT: 159px; VISIBILITY: hidden; POSITION: absolute; TOP: 180px" class="msgLabel">
			<input type="radio" value="1" id="idCreditCheckOptOutYes" name="CreditCheckOptOutIndicator" style="LEFT: 1px; TOP: 0px"> Yes <input type="radio" value="0" id="idCreditCheckOptOutNo" name="CreditCheckOptOutIndicator"> No
		</span>
	</div>


	<div id="divCreditDecisionDetail" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 1054px; HEIGHT: 235px" class="msgGroup">
		<span style="LEFT: 4px; POSITION: absolute; TOP: -13px" class="msgLabelHead" id=SPAN1>
			Credit Decision Detail
		</span>

		<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
			Experian 
Ref
			<span style="LEFT: 95px; POSITION: absolute; TOP: 0px">
				<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgTxt"></label>
				<input id="txtExperianRef" maxlength="10" style="LEFT: 1px; WIDTH: 100px; POSITION: absolute; TOP: -3px; HEIGHT: 20px" class="msgTxt" size=23>
			</span>
		</span>

		<span style="LEFT: 223px; POSITION: absolute; TOP: 10px" class="msgLabel">
			Eligibility Indicator
			<span style="LEFT: 95px; POSITION: absolute; TOP: 0px">
				<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgTxt"></label>
				<input id="txtEligibilityIndicator" maxlength="10" style="LEFT: 4px; WIDTH: 100px; POSITION: absolute; TOP: -3px; HEIGHT: 20px" class="msgTxt" size=18>
			</span>
		</span>

		<span style="LEFT: 459px; POSITION: absolute; TOP: 10px" class="msgLabel">
			Delphi 
Total
			<span style="LEFT: 95px; POSITION: absolute; TOP: 0px">
				<label style="LEFT: -29px; POSITION: absolute; TOP: 0px" class="msgTxt" id=LABEL1></label>
				<input id="txtDelphiTotal" maxlength="3" style="LEFT: -26px; WIDTH: 60px; POSITION: absolute; TOP: -3px; HEIGHT: 20px" class="msgTxt" size=10>
			</span>
		</span>
	</div>

	<div id="divThirdPartySummary" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 1106px; HEIGHT: 235px" class="msgGroup">
		<span style="LEFT: 4px; POSITION: absolute; TOP: 9px" class="msgLabelHead">
			Third Party Summary
		</span>
		<span id="spnTblThirdParty" style="LEFT: 4px; POSITION: absolute; TOP: 25px">
			<table id="tblThirdParty" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">
				<td width="15%" class="TableHead">Type</td>
				<td width="20%" class="TableHead">Contact Name/ Account Number</td>
				<td width="20%" class="TableHead">Company Name</td>
				<td width="28%" class="TableHead">Address</td>
				<td width="17%" class="TableHead">Repayment Bank Account</td>
			</tr>
			<tr id="row01">
				<td class="TableTopLeft">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row02">
				<td class="TableBottomLeft">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomRight">&nbsp;</td>
			</tr></table>
		</span>
		
		<span style="LEFT: 4px; POSITION: absolute; TOP: 95px" class="msgLabel">
			Contact Details
			<span style="LEFT: 90px; POSITION: absolute; TOP: -8px">
				<input id="btnContactThirdParty" type ="button" style="WIDTH: 26px; HEIGHT: 26px" class="msgDDButton">
			</span>
		</span>
	</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 1284px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP040attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var m_sReadOnly = "";
var m_sUnderReview = "";
var m_sStageName = "";
var m_sStageId = "";
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_AppSummaryXML = null;

var m_iThirdPartyTableLength = 2;
var m_iOccupancyTableLength = 3;
var m_iEmploymentTableLength = 3;
var m_iApplicantsTableLength = 4;
var m_iIntroducerTableLength = 3;
var m_iLoanTableLength = 4;

var scScreenFunctions;
var m_blnReadOnly = false;
var m_RequestArray = null;
var m_BaseNonPopupWindow = null;

var m_sCustomer1Number = "";
var m_sCustomer1Version = "";
var m_sCustomer1Name = "";
var m_sCustomer2Number = "";
var m_sCustomer2Version = "";
var m_sCustomer2Name = "";
var m_sCustomer3Number = "";
var m_sCustomer3Version = "";
var m_sCustomer3Name = "";
var m_sCustomer4Number = "";
var m_sCustomer4Version = "";
var m_sCustomer4Name = "";
var m_sCustomer5Number = "";
var m_sCustomer5Version = "";
var m_sCustomer5Name = "";
var m_sCurrentCustomerNumber = "";
var m_sCurrentCustomerVersionNumber = "";
var m_sCurrentCustomerName = "";
var m_sIdUnitName = "";

<% /* PSC 16/03/2007 EP2_1726 - Start */ %>
var m_XMLIntroducerSearchType = null;
var m_XMLListingStatus = null;
var m_XMLFirmStatus = null;
<% /* PSC 16/03/2007 EP2_1726 - End */ %>

function window.onload()
{
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	//scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	//FW030SetTitles("Application Summary","AP040",scScreenFunctions);
	
	RetrieveData();
	
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();	
	//idCurrency
	scScreenFunctions.SetContextParameter(window,"idCurrency", null);
	scScreenFunctions.SetThisCurrency(frmScreen,null)
	//scScreenFunctions.SetCurrency(window,frmScreen);

	SetMasks();
	Validation_Init();
	
	PopulateScreen();
	
	//m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP040P");
	
}

function RetrieveData()
{
	/*TEST
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","00005010");	
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber","1");
	END*/

	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters	= sArguments[4];
	
	m_BaseNonPopupWindow = sArguments[5];
	
	
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();	
	m_RequestArray = sParameters[0] // array
	m_sMetaAction = sParameters[1];
	m_sReadOnly = 1;
	m_sUnderReview = sParameters[3];
	m_sStageName = sParameters[4];
	m_sStageId = sParameters[5];
	m_sApplicationNumber = sParameters[6];
	m_sApplicationFactFindNumber = sParameters[7];
	
	m_sCustomer1Number = sParameters[8];
	m_sCustomer1Version = sParameters[9];
	m_sCustomer2Number = sParameters[10];
	m_sCustomer2Version = sParameters[11];
	m_sCustomer3Number = sParameters[12];
	m_sCustomer3Version = sParameters[13];
	m_sCustomer4Number = sParameters[14];
	m_sCustomer4Version = sParameters[15];
	m_sCustomer5Number = sParameters[16];
	m_sCustomer5Version = sParameters[17];
	
	m_sCustomerName1 = sParameters[18];
	m_sCustomerName2 = sParameters[19];
	m_sCustomerName3 = sParameters[20];
	m_sCustomerName4 = sParameters[21];
	m_sCustomerName5 = sParameters[22];
}
function PopulateScreen()
{

	<% /* PSC 16/03/2007 EP2_1726 */ %>
	GetComboData();
	
	if(m_sCurrentCustomerNumber == "" &&
	m_sCurrentCustomerVersionNumber == "")
	{
		// set up to point to the first customer from context
		m_sCurrentCustomerNumber = m_sCustomer1Number;
		m_sCurrentCustomerVersionNumber = m_sCustomer1Version;
	}

	GetCustomer();
		
	// Set fields readonly
	scScreenFunctions.SetCollectionState(divApplicants,"R");
	scScreenFunctions.SetCollectionState(divAppDetails, "R");
	scScreenFunctions.SetCollectionState(divMortgageDetails, "R");
	scScreenFunctions.SetCollectionState(divLoanDetails, "R");
	scScreenFunctions.SetCollectionState(divAppDetails1, "R");
	scScreenFunctions.SetCollectionState(divAppDetails2, "R");
	scScreenFunctions.SetCollectionState(divCreditDecisionDetail, "R");
	scScreenFunctions.SetCollectionState(divThirdPartySummary, "R");
		
	if(GetFullApplicationSummary())
	{
		PopulateListBox();
		PopulateDetails();
	}
	
}
function GetFullApplicationSummary()
{
	m_AppSummaryXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_AppSummaryXML.CreateRequestTagFromArray(m_RequestArray, "");
	
	m_AppSummaryXML.XMLDocument.documentElement.setAttribute("CRUD_OP", "READ");
	m_AppSummaryXML.XMLDocument.documentElement.setAttribute("SCHEMA_NAME", "FullApplicationSummaryData");
	m_AppSummaryXML.XMLDocument.documentElement.setAttribute("ENTITY_REF", "APPLICATION");
	m_AppSummaryXML.XMLDocument.documentElement.setAttribute("COMBOTYPELOOKUP", "EX");
	m_AppSummaryXML.XMLDocument.documentElement.setAttribute("COMBOLOOKUP", "Y");
	m_AppSummaryXML.XMLDocument.documentElement.setAttribute("postProcRef", "FullApplicationSummaryData.xsl");
	var xn = m_AppSummaryXML.CreateActiveTag("APPLICATION");
	xn.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	xn.setAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	
	m_AppSummaryXML.RunASP(document,"omCRUDIf.asp");
	
	if(m_AppSummaryXML.IsResponseOK())
		return true;
	return false;
}
function PopulateListBox()
{
	var nCustRowPos = scScrollTableApplicants.getRowSelected();
	
	//get customer count
	m_AppSummaryXML.ActiveTag = null;
	m_AppSummaryXML.CreateTagList("CUSTOMER");
	var iNumberOfApplicants = m_AppSummaryXML.ActiveTagList.length;

	//get introducers count
	m_AppSummaryXML.ActiveTag = null;
	<% /* PSC 16/03/2007 EP2_1726 */ %>
	m_AppSummaryXML.CreateTagList("INTERMEDIARY");
	var iNumberOfIntroducers = m_AppSummaryXML.ActiveTagList.length;
		
	//for selected customer: get employment count
	m_AppSummaryXML.ActiveTag = null;
	m_AppSummaryXML.CreateTagList("CUSTOMER");
	if(nCustRowPos == -1)
	{
		m_AppSummaryXML.SelectTagListItem(0);
	}
	else
	{
		m_AppSummaryXML.SelectTagListItem(nCustRowPos-1);
	}
	m_AppSummaryXML.CreateTagList("CUSTOMEREMPLOYMENT");
	var iNumberOfEmployments = m_AppSummaryXML.ActiveTagList.length;

	//for selected customer: get address count
	m_AppSummaryXML.ActiveTag = null;
	m_AppSummaryXML.CreateTagList("CUSTOMER");
	if(nCustRowPos == -1)
	{
		m_AppSummaryXML.SelectTagListItem(0);
		m_iCurrentCustomer = 0;
	}
	else
	{
		m_AppSummaryXML.SelectTagListItem(nCustRowPos-1);
		m_iCurrentCustomer = nCustRowPos-1;
	}
	m_AppSummaryXML.CreateTagList("CUSTOMERADDRESS");
	var iNumberOfAddresses = m_AppSummaryXML.ActiveTagList.length;
		
	//get loan count
	m_AppSummaryXML.ActiveTag = null;
	m_AppSummaryXML.CreateTagList("MORTGAGEDETAILS");
	var iNumberOfLoans = m_AppSummaryXML.ActiveTagList.length;		
	
	//for selected customer: get third party count
	m_AppSummaryXML.ActiveTag = null;
	m_AppSummaryXML.CreateTagList("APPLICATIONTHIRDPARTYLIST");		
	if(nCustRowPos == -1)
	{
		m_AppSummaryXML.SelectTagListItem(0);
	}
	else
	{
		m_AppSummaryXML.SelectTagListItem(nCustRowPos-1);
	}
	m_AppSummaryXML.CreateTagList("APPLICATIONTHIRDPARTY");
	var iNumberOfThirdParties = m_AppSummaryXML.ActiveTagList.length;
	
	//for debugging purposes
	/*alert('Applicant count: ' + iNumberOfApplicants);
	alert('Introducer count: ' + iNumberOfIntroducers);
	alert('Employment count: ' + iNumberOfEmployments);
	alert('Address count: ' + iNumberOfAddresses);
	alert('Introducer count: ' + iNumberOfIntroducers);
	alert('Loan count: ' + iNumberOfLoans);
	alert('Third party count: ' + iNumberOfThirdParties);*/
	
	// initialise scroll tables
	scScrollTableApplicants.initialiseTable(tblApplicants, 0, "", ShowListApplicants, m_iApplicantsTableLength, iNumberOfApplicants);
	scScrollTableEmployment.initialiseTable(tblEmployment, 0, "", ShowListEmployment, m_iEmploymentTableLength, iNumberOfEmployments);
	scScrollTableOccupancy.initialiseTable(tblOccupancy, 0, "", ShowListOccupancy, m_iOccupancyTableLength, iNumberOfAddresses);
	scScrollTableIntroducer.initialiseTable(tblIntroducer, 0, "", ShowListIntroducer, m_iIntroducerTableLength, iNumberOfIntroducers);
	scScrollTableLoan.initialiseTable(tblLoan, 0, "", ShowListLoan, m_iLoanTableLength, iNumberOfLoans);
	scScrollTableThirdParty.initialiseTable(tblThirdParty, 0, "", ShowListTP, m_iThirdPartyTableLength ,iNumberOfThirdParties);
	
	//display table data
	ShowListApplicants(0, tblApplicants);
	if(iNumberOfApplicants > 0) 
	{
		if (nCustRowPos == -1) //if the applicant table is being populated for the first time
		{
		scScrollTableApplicants.setRowSelected(1);
		}
		else
		{
		scScrollTableApplicants.setRowSelected(nCustRowPos);
		}
	}
	
	ShowListEmployment(0);
	if(iNumberOfApplicants > 0) scScrollTableEmployment.setRowSelected(1);
	
	ShowListOccupancy(0);
	if(iNumberOfApplicants > 0) scScrollTableOccupancy.setRowSelected(1);
	
	ShowListIntroducer(0);
	if(iNumberOfApplicants > 0) scScrollTableIntroducer.setRowSelected(1);

	ShowListLoan(0);
	if(iNumberOfApplicants > 0) scScrollTableLoan.setRowSelected(1);
	
	ShowListTP(0);
	if(iNumberOfApplicants > 0) scScrollTableThirdParty.setRowSelected(1);
}


function ShowListApplicants(nStart)
{
	scScrollTableApplicants.clear();
	m_AppSummaryXML.ActiveTag = null;
	m_AppSummaryXML.CreateTagList("CUSTOMER");
	
	for (iCount = 0; iCount < m_AppSummaryXML.ActiveTagList.length && iCount < m_iApplicantsTableLength; iCount++)
	{
		m_AppSummaryXML.SelectTagListItem(iCount + nStart);
			
		if (m_AppSummaryXML.SelectTag(m_AppSummaryXML.ActiveTag, "CUSTOMERDETAILS") != null)
		{	
				
			var sName = m_AppSummaryXML.GetAttribute("CUSTOMERFIRSTFORENAME");
			var sSurname = m_AppSummaryXML.GetAttribute("CUSTOMERSURNAME");
			var sTitle = m_AppSummaryXML.GetAttribute("CUSTOMERTITLETEXT");
			var sAge = m_AppSummaryXML.GetAttribute("AGE");
			var sMaritalStatus = m_AppSummaryXML.GetAttribute("MARITALSTATUSTEXT");
			var sDOB = m_AppSummaryXML.GetAttribute("DATEOFBIRTH");
			var sGender = m_AppSummaryXML.GetAttribute("GENDERTEXT");
			var sRole = m_AppSummaryXML.GetAttribute("CUSTOMERROLETEXT")
		
			scScreenFunctions.SizeTextToField(tblApplicants.rows(iCount+1).cells(0),sSurname);
			scScreenFunctions.SizeTextToField(tblApplicants.rows(iCount+1).cells(1),sName );
			scScreenFunctions.SizeTextToField(tblApplicants.rows(iCount+1).cells(2),sTitle );
			scScreenFunctions.SizeTextToField(tblApplicants.rows(iCount+1).cells(3),sMaritalStatus);
			scScreenFunctions.SizeTextToField(tblApplicants.rows(iCount+1).cells(4),sDOB );
			scScreenFunctions.SizeTextToField(tblApplicants.rows(iCount+1).cells(5),sAge );
			scScreenFunctions.SizeTextToField(tblApplicants.rows(iCount+1).cells(6),sGender);
			scScreenFunctions.SizeTextToField(tblApplicants.rows(iCount+1).cells(7),sRole);
			tblApplicants.rows(iCount+1).setAttribute("TagListItem", (iCount + nStart));
		}
	}
}

function ShowListEmployment(nStart)
{
	scScrollTableEmployment.clear();
	
	m_AppSummaryXML.ActiveTag = null;
	m_AppSummaryXML.CreateTagList("CUSTOMER");
	m_AppSummaryXML.SelectTagListItem(m_iCurrentCustomer);	
	m_AppSummaryXML.CreateTagList("CUSTOMEREMPLOYMENT");
		
	for (iCount = 0; iCount < m_AppSummaryXML.ActiveTagList.length && iCount < m_iEmploymentTableLength; iCount++)
	{
		m_AppSummaryXML.SelectTagListItem(iCount + nStart);
		
		var sEmploymentStatus = m_AppSummaryXML.GetAttribute("EMPLOYMENTSTATUSTEXT");
		var sOccupation = m_AppSummaryXML.GetAttribute("OCCUPATIONTEXT");
		var sDateStarted = m_AppSummaryXML.GetAttribute("DATESTARTED");
		var sDateEnded = m_AppSummaryXML.GetAttribute("DATEENDED");
		var sTimeEmployed = m_AppSummaryXML.GetAttribute("TIMEEMPLOYED");
		var sGrossIncome = m_AppSummaryXML.GetAttribute("GROSSINCOME");
		var sMain = m_AppSummaryXML.GetAttribute("MAINSTATUSTEXT");
				
		scScreenFunctions.SizeTextToField(tblEmployment.rows(iCount+1).cells(0),sEmploymentStatus);
		scScreenFunctions.SizeTextToField(tblEmployment.rows(iCount+1).cells(1),sOccupation );
		scScreenFunctions.SizeTextToField(tblEmployment.rows(iCount+1).cells(2),sDateStarted );
		scScreenFunctions.SizeTextToField(tblEmployment.rows(iCount+1).cells(3),sDateEnded);
		scScreenFunctions.SizeTextToField(tblEmployment.rows(iCount+1).cells(4),sTimeEmployed );
		scScreenFunctions.SizeTextToField(tblEmployment.rows(iCount+1).cells(5),sGrossIncome );
		scScreenFunctions.SizeTextToField(tblEmployment.rows(iCount+1).cells(6),sMain);
		tblEmployment.rows(iCount+1).setAttribute("TagListItem", (iCount + nStart));
	}						

}

<% /* PSC 16/03/2007 EP2_1726 - Start */ %>
function ShowListIntroducer(nStart)
{
	scScrollTableIntroducer.clear();
	
	m_AppSummaryXML.ActiveTag = null;
	m_AppSummaryXML.CreateTagList("INTERMEDIARY");
	
	var introducerId;

	for (iCount = 0; iCount < m_AppSummaryXML.ActiveTagList.length && iCount < m_iIntroducerTableLength; iCount++)
	{
		m_AppSummaryXML.SelectTagListItem(iCount + nStart);
		
		var sPrincipalFirmID = m_AppSummaryXML.GetAttribute("PRINCIPALFIRMID");
		var sARFirmID = m_AppSummaryXML.GetAttribute("ARFIRMID");
		var sClubAssocID = m_AppSummaryXML.GetAttribute("CLUBNETWORKASSOCIATIONID");
		var sIntroducerID = m_AppSummaryXML.GetAttribute("INTRODUCERID");
		var fsaRef =  m_AppSummaryXML.GetAttribute("FSAREFNUMBER");
		var sName =  m_AppSummaryXML.GetAttribute("NAME");
		var sAddress = '';
		var sType = '';
		var listingStatus = '';
		var sStatus = '';
		var sId = '';
		
		var sValidateType =  m_AppSummaryXML.GetAttribute("LISTSTATUSVALIDATION");
		var objValueName = m_XMLListingStatus.selectSingleNode(".//LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE='"+ sValidateType +"']/VALUENAME");
		
		if (objValueName != null)
		{
			listingStatus = objValueName.text;
		}
		
		var sStatusValue =  m_AppSummaryXML.GetAttribute("STATUS");
		objValueName = m_XMLFirmStatus.selectSingleNode(".//LISTENTRY[VALUEID='"+ sStatusValue +"']/VALUENAME");
		
		if (objValueName != null)
		{
			sStatus = objValueName.text;
		}
		
		sValidateType = m_AppSummaryXML.GetAttribute("TYPEVALIDATION");
		objValueName = m_XMLIntroducerSearchType.selectSingleNode(".//LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE='"+ sValidateType +"']/VALUENAME");
		
		if (objValueName)
		{
			sType = objValueName.text;
		}
		
		if (sPrincipalFirmID && sPrincipalFirmID.length > 0)
		{
			tblIntroducer.rows(iCount+1).setAttribute("PrincipalFirmID", sPrincipalFirmID);
			sId = sPrincipalFirmID;
		}
			
		if (sARFirmID && sARFirmID.length > 0)
		{
			tblIntroducer.rows(iCount+1).setAttribute("ARFirmID", sARFirmID);
			sId = sARFirmID;
		}
		if (sClubAssocID && sClubAssocID.length > 0)
		{
			tblIntroducer.rows(iCount+1).setAttribute("ClubAssocID", sClubAssocID);
			sId = sClubAssocID;
		}
		
		if (sIntroducerID && sIntroducerID.length > 0)
		{	
			tblIntroducer.rows(iCount+1).setAttribute("IntroducerID", sIntroducerID);
			sId = sIntroducerID;
		}
		
		if (m_AppSummaryXML.GetAttribute("ADDRESSLINE1") != '')
			sAddress = m_AppSummaryXML.GetAttribute("ADDRESSLINE1") + " ";

		if (m_AppSummaryXML.GetAttribute("ADDRESSLINE2") != '')
			sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("ADDRESSLINE2") + " ";

		if (m_AppSummaryXML.GetAttribute("ADDRESSLINE3") != '')
			sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("ADDRESSLINE3") + " ";

		if (m_AppSummaryXML.GetAttribute("ADDRESSLINE4") != '')
			sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("ADDRESSLINE4") + " ";
							
		if (m_AppSummaryXML.GetAttribute("ADDRESSLINE5") != '')
			sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("ADDRESSLINE5") + " ";
		
		if (m_AppSummaryXML.GetAttribute("ADDRESSLINE6") != '')
			sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("ADDRESSLINE6") + " ";
									
		scScreenFunctions.SizeTextToField(tblIntroducer.rows(iCount+1).cells(0),sId);
		scScreenFunctions.SizeTextToField(tblIntroducer.rows(iCount+1).cells(1),fsaRef);
		scScreenFunctions.SizeTextToField(tblIntroducer.rows(iCount+1).cells(2),sName);
		scScreenFunctions.SizeTextToField(tblIntroducer.rows(iCount+1).cells(3),sAddress);
		scScreenFunctions.SizeTextToField(tblIntroducer.rows(iCount+1).cells(4),sType);
		scScreenFunctions.SizeTextToField(tblIntroducer.rows(iCount+1).cells(5),listingStatus);
		scScreenFunctions.SizeTextToField(tblIntroducer.rows(iCount+1).cells(6),sStatus);
	}						
}
<% /* PSC 16/03/2007 EP2_1726 - End */ %>

function ShowListOccupancy(nStart)
{
	scScrollTableOccupancy.clear();
	
	m_AppSummaryXML.ActiveTag = null;
	m_AppSummaryXML.CreateTagList("CUSTOMER");
	m_AppSummaryXML.SelectTagListItem(m_iCurrentCustomer);
	m_AppSummaryXML.CreateTagList("CUSTOMERADDRESS");
	
	for (iCount=0;iCount < m_AppSummaryXML.ActiveTagList.length && iCount < m_iOccupancyTableLength; iCount++)
	{
		m_AppSummaryXML.SelectTagListItem(iCount+nStart)
		
		if (m_AppSummaryXML != null)
		{
			
			var sAddress = '';
						
			if (m_AppSummaryXML.GetAttribute("BUILDINGORHOUSENUMBER") != '')
				sAddress = m_AppSummaryXML.GetAttribute("BUILDINGORHOUSENUMBER") + " ";
			
			if (m_AppSummaryXML.GetAttribute("BUILDINGORHOUSENAME") != '')
				sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("BUILDINGORHOUSENAME") + " ";

			if (m_AppSummaryXML.GetAttribute("FLATNUMBER") != '')
				sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("FLATNUMBER") + " ";
		
			if (m_AppSummaryXML.GetAttribute("STREET") != '')
				sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("STREET") + " ";
					
			if (m_AppSummaryXML.GetAttribute("DISTRICT") != '')
				sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("DISTRICT") + " ";
		
			if (m_AppSummaryXML.GetAttribute("TOWN") != '')
				sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("TOWN") + " ";
		
			if (m_AppSummaryXML.GetAttribute("COUNTY") != '')
				sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("COUNTY") + " ";
		
			if (m_AppSummaryXML.GetAttribute("COUNTRY_TEXT") != '')
				sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("COUNTRY_TEXT") + " ";
		
			if (m_AppSummaryXML.GetAttribute("POSTCODE") != '')
				sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("POSTCODE") + " ";
						
			var sNatureOfOccupancy = m_AppSummaryXML.GetAttribute("NATUREOFOCCUPANCYTEXT");
			var sDateMovedIn = m_AppSummaryXML.GetAttribute("DATEMOVEDIN");
			var sDateMovedOut = m_AppSummaryXML.GetAttribute("DATEMOVEDOUT");
			var sTimeAtAddress = m_AppSummaryXML.GetAttribute("TIMEATADDRESS");
		
			scScreenFunctions.SizeTextToField(tblOccupancy.rows(iCount+1).cells(0),sAddress);
			scScreenFunctions.SizeTextToField(tblOccupancy.rows(iCount+1).cells(1),sNatureOfOccupancy );
			scScreenFunctions.SizeTextToField(tblOccupancy.rows(iCount+1).cells(2),sDateMovedIn );
			scScreenFunctions.SizeTextToField(tblOccupancy.rows(iCount+1).cells(3),sDateMovedOut );
			scScreenFunctions.SizeTextToField(tblOccupancy.rows(iCount+1).cells(4),sTimeAtAddress);
			tblOccupancy.rows(iCount+1).setAttribute("TagListItem", (iCount + nStart));
		}		
	}						
}

function ShowListLoan(nStart)
{
		scScrollTableLoan.clear();		

		m_AppSummaryXML.ActiveTag = null;
		m_AppSummaryXML.CreateTagList("MORTGAGEDETAILS");	

		for (iCount = 0; iCount < m_AppSummaryXML.ActiveTagList.length && iCount < m_iLoanTableLength; iCount++)
		{
			m_AppSummaryXML.SelectTagListItem(iCount + nStart);
			
			if (m_AppSummaryXML.SelectTag(m_AppSummaryXML.ActiveTag, "LOANCOMPONENT") != null)
			{			
				var sMortgageProductCode = m_AppSummaryXML.GetAttribute("MORTGAGEPRODUCTCODE");
				var sProductType = m_AppSummaryXML.GetAttribute("PRODUCTTYPE");
				var sProductDescription = m_AppSummaryXML.GetAttribute("PRODUCTDESCRIPTION");
				var sResolvedRate = m_AppSummaryXML.GetAttribute("RESOLVEDRATE");
				var sAdjustedRate = m_AppSummaryXML.GetAttribute("ADJUSTEDRATE");
				var sLoanAmount = m_AppSummaryXML.GetAttribute("LOANAMOUNT");
				var sTerm
			
				if (m_AppSummaryXML.GetAttribute("TERMINYEARS") != "")
					sTerm = m_AppSummaryXML.GetAttribute("TERMINYEARS") + " yrs";
			
				if (m_AppSummaryXML.GetAttribute("TERMINMONTHS") != "")
					sTerm = sTerm + " " + m_AppSummaryXML.GetAttribute("TERMINMONTHS") + " mths";				
			
				var sRepaymentMethod = m_AppSummaryXML.GetAttribute("REPAYMENTMETHODTEXT");
				var sAPR = m_AppSummaryXML.GetAttribute("APR");
				var sNetMonthlyCost = m_AppSummaryXML.GetAttribute("NETMONTHLYCOST");
				var sFinalRateMonthlyCost = m_AppSummaryXML.GetAttribute("FINALRATEMONTHLYCOST");
			
				scScreenFunctions.SizeTextToField(tblLoan.rows(iCount+1).cells(0),sMortgageProductCode);
				scScreenFunctions.SizeTextToField(tblLoan.rows(iCount+1).cells(1),sProductType);
				scScreenFunctions.SizeTextToField(tblLoan.rows(iCount+1).cells(2),sProductDescription);
				scScreenFunctions.SizeTextToField(tblLoan.rows(iCount+1).cells(3),sResolvedRate);
				scScreenFunctions.SizeTextToField(tblLoan.rows(iCount+1).cells(4),sAdjustedRate);
				scScreenFunctions.SizeTextToField(tblLoan.rows(iCount+1).cells(5),sLoanAmount);
				scScreenFunctions.SizeTextToField(tblLoan.rows(iCount+1).cells(6),sTerm);
				scScreenFunctions.SizeTextToField(tblLoan.rows(iCount+1).cells(7),sRepaymentMethod);
				scScreenFunctions.SizeTextToField(tblLoan.rows(iCount+1).cells(8),sAPR);
				scScreenFunctions.SizeTextToField(tblLoan.rows(iCount+1).cells(9),sNetMonthlyCost);
				scScreenFunctions.SizeTextToField(tblLoan.rows(iCount+1).cells(10),sFinalRateMonthlyCost);
				tblLoan.rows(iCount+1).setAttribute("TagListItem", (iCount + nStart));
			}
		}				
}


function ShowListTP(nStart)
{
		scScrollTableThirdParty.clear();
		
		m_AppSummaryXML.ActiveTag = null;
		m_AppSummaryXML.CreateTagList("APPLICATIONTHIRDPARTYLIST");
			
		m_AppSummaryXML.SelectTagListItem(nStart);
		m_AppSummaryXML.CreateTagList("APPLICATIONTHIRDPARTY");
		//alert(nStart);		
		//alert(m_AppSummaryXML.xml);
		//alert(m_AppSummaryXML.XMLDocument.xml);
		//alert(m_AppSummaryXML.ActiveTagList.length);
		
		for (iCount = 0; iCount < m_AppSummaryXML.ActiveTagList.length && iCount < m_iThirdPartyTableLength; iCount++)
		{
		
			m_AppSummaryXML.SelectTagListItem(iCount + nStart);
			
			if (m_AppSummaryXML.SelectTag(m_AppSummaryXML.ActiveTag, "THIRDPARTY") != null)
			{
				
				var sTypeText = m_AppSummaryXML.GetTagText("THIRDPARTYTYPE_TEXT");
				var sType = m_AppSummaryXML.GetTagText("THIRDPARTYTYPE");			
				var sCompanyName = m_AppSummaryXML.GetTagText("COMPANYNAME");
				var sContactNameAccount = m_AppSummaryXML.GetTagText("ACCOUNTNUMBER");
				var sRepaymentBankAccount = m_AppSummaryXML.GetTagText("REPAYMENTBANKACCOUNTINDICATOR");
				var sAddress = '';
				var ThirdPartyXML = m_AppSummaryXML.ActiveTag;
					
				if (m_AppSummaryXML.SelectTag(m_AppSummaryXML.ActiveTag, "CONTACTDETAILS") != null)
				{
					if (sContactNameAccount.length == 0)
					{
						var sContactNameAccount = m_AppSummaryXML.GetTagText("CONTACTFORENAME") + " " + m_AppSummaryXML.GetTagText("CONTACTSURNAME");
					}
				}

				m_AppSummaryXML.ActiveTag = ThirdPartyXML;
					
				if (m_AppSummaryXML.SelectTag(m_AppSummaryXML.ActiveTag, "ADDRESS") != null)
				{
					
					if (m_AppSummaryXML.GetTagText("BUILDINGORHOUSENUMBER") != '')
						sAddress = m_AppSummaryXML.GetTagText("BUILDINGORHOUSENUMBER") + " ";

					if (m_AppSummaryXML.GetTagText("BUILDINGORHOUSENAME") != '')
						sAddress = sAddress + " " + m_AppSummaryXML.GetTagText("BUILDINGORHOUSENAME") + " ";

					if (m_AppSummaryXML.GetTagText("FLATNUMBER") != '')
						sAddress = sAddress + " " + m_AppSummaryXML.GetTagText("FLATNUMBER") + " ";
					
					if (m_AppSummaryXML.GetTagText("STREET") != '')
						sAddress = sAddress + " " + m_AppSummaryXML.GetTagText("STREET") + " ";
				
					if (m_AppSummaryXML.GetTagText("DISTRICT") != '')
						sAddress = sAddress + " " + m_AppSummaryXML.GetTagText("DISTRICT") + " ";
		
					if (m_AppSummaryXML.GetTagText("TOWN") != '')
						sAddress = sAddress + " " + m_AppSummaryXML.GetTagText("TOWN") + " ";
		
					if (m_AppSummaryXML.GetTagText("COUNTY") != '')
						sAddress = sAddress + " " + m_AppSummaryXML.GetTagText("COUNTY") + " ";
		
					if (m_AppSummaryXML.GetTagText("COUNTRY_TEXT") != '')
						sAddress = sAddress + " " + m_AppSummaryXML.GetTagText("COUNTRY_TEXT") + " ";
		
					if (m_AppSummaryXML.GetTagText("POSTCODE") != '')
						sAddress = sAddress + " " + m_AppSummaryXML.GetTagText("POSTCODE") + " ";
			
				}
				
				
				if (sRepaymentBankAccount == 1) 
					sRepaymentBankAccount = 'Yes';
				else if (sRepaymentBankAccount == 0 && sType == 3) //3=Bank
					sRepaymentBankAccount = 'No';
				else
					sRepaymentBankAccount = 'N/A';
	
				scScreenFunctions.SizeTextToField(tblThirdParty.rows(iCount+1).cells(0),sTypeText);
				scScreenFunctions.SizeTextToField(tblThirdParty.rows(iCount+1).cells(1),sContactNameAccount);
				scScreenFunctions.SizeTextToField(tblThirdParty.rows(iCount+1).cells(2),sCompanyName);
				scScreenFunctions.SizeTextToField(tblThirdParty.rows(iCount+1).cells(3),sAddress);
				scScreenFunctions.SizeTextToField(tblThirdParty.rows(iCount+1).cells(4),sRepaymentBankAccount);
				tblThirdParty.rows(iCount+1).setAttribute("TagListItem", (iCount + nStart));										
			}	
		}						
}

function PopulateDetails()
{
	var dblAnnualIncome
	var dblConfirmedAnnualIncome
	
	m_AppSummaryXML.ActiveTag = null;
	m_AppSummaryXML.CreateTagList("MORTGAGEPROPERTYADDRESS");
	
	if(m_AppSummaryXML.ActiveTagList.length > 0)
	{
		m_AppSummaryXML.SelectTagListItem(0);
		var sAddress = '';
		
		if (m_AppSummaryXML.GetAttribute("BUILDINGORHOUSENUMBER") != '')
			sAddress = m_AppSummaryXML.GetAttribute("BUILDINGORHOUSENUMBER") + " ";

		if (m_AppSummaryXML.GetAttribute("BUILDINGORHOUSENAME") != '')
			sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("BUILDINGORHOUSENAME") + " ";

		if (m_AppSummaryXML.GetAttribute("FLATNUMBER") != '')
			sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("FLATNUMBER") + " ";
		
		if (m_AppSummaryXML.GetAttribute("STREET") != '')
			sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("STREET") + " ";
				
		if (m_AppSummaryXML.GetAttribute("DISTRICT") != '')
			sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("DISTRICT") + " ";
		
		if (m_AppSummaryXML.GetAttribute("TOWN") != '')
			sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("TOWN") + " ";
		
		if (m_AppSummaryXML.GetAttribute("COUNTY") != '')
			sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("COUNTY") + " ";
		
		if (m_AppSummaryXML.GetAttribute("COUNTRY_TEXT") != '')
			sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("COUNTRY_TEXT") + " ";
		
		if (m_AppSummaryXML.GetAttribute("POSTCODE") != '')
			sAddress = sAddress + " " + m_AppSummaryXML.GetAttribute("POSTCODE") + " ";
		
		frmScreen.txtMortgagePropertyAddress.value = sAddress;
		
		<% /* MAR1040 Populate fields from IncomeSummary table */ %>
		//dblAnnualIncome = parseFloat(m_AppSummaryXML.GetAttribute("NETALLOWABLEANNUALINCOME"))
		//dblConfirmedAnnualIncome = parseFloat(m_AppSummaryXML.GetAttribute("NETCONFIRMEDALLOWABLEINCOME"))
		
		//if (! isNaN(dblAnnualIncome))
		//	frmScreen.txtAnnualIncome.value = dblAnnualIncome; 	
		//else
		//	frmScreen.txtAnnualIncome.value = 0;

		//if (! isNaN(dblConfirmedAnnualIncome))
		//	frmScreen.txtConfirmedAnnualIncome.value = dblConfirmedAnnualIncome; 	
		//else
		//	frmScreen.txtConfirmedAnnualIncome.value = 0;
			
	}
	if(m_AppSummaryXML.SelectTag(null, "APPLICATION") != null)
	{
		frmScreen.txtProductCategory.value = m_AppSummaryXML.GetAttribute("NATUREOFLOANTEXT");
		frmScreen.txtTypeOfBuyer.value = m_AppSummaryXML.GetAttribute("TYPEOFBUYERTEXT");
		frmScreen.txtDeclaredIncomeMultiple.value = m_AppSummaryXML.GetAttribute("DECLAREDINCOMEMULTIPLE");
		frmScreen.txtConfirmedIncomeMultiple.value = m_AppSummaryXML.GetAttribute("CONFIRMEDINCOMEMULTIPLE");
		frmScreen.txtDIPDate.value = m_AppSummaryXML.GetAttribute("DIPDATE");
		frmScreen.txtApplicationDate.value = m_AppSummaryXML.GetAttribute("APPLICATIONRECEIVEDDATE");
		frmScreen.txtIncomeStatus.value = m_AppSummaryXML.GetAttribute("APPLICATIONINCOMESTATUS");
		frmScreen.txtLevelOfAdvice.value = m_AppSummaryXML.GetAttribute("LEVELOFADVICE");
		frmScreen.txtPresentValue.value = m_AppSummaryXML.GetAttribute("PRESENTVALUATION");
		frmScreen.txtPostValue.value = m_AppSummaryXML.GetAttribute("PRESENTVALUATION");
		frmScreen.txtRetentionAmount.value = m_AppSummaryXML.GetAttribute("RETENTIONAMOUNT");
		frmScreen.txtPropertyValInLTV.value = m_AppSummaryXML.GetAttribute("VALUATION");
		frmScreen.txtRegulationIndicator.value = m_AppSummaryXML.GetAttribute("REGULATIONINDICATOR");
		
		scScreenFunctions.SetRadioGroupValue(frmScreen, "CreditCheckOptOutIndicator", m_AppSummaryXML.GetAttribute("OPTEDOUT")) ;
		if(m_AppSummaryXML.GetAttribute("OPTEDOUT") != "")
		{
			if(m_AppSummaryXML.GetAttribute("OPTEDOUT") == "1")
				frmScreen.idCreditCheckOptOutYes.checked = true;
			else
				frmScreen.idCreditCheckOptOutNo.checked = true;
		}
	}
	
	if(m_AppSummaryXML.SelectTag(null, "CREDITDETAILS") != null)
	{
		frmScreen.txtExperianRef.value = m_AppSummaryXML.GetAttribute("EXPERIANREFERENCE");
		frmScreen.txtEligibilityIndicator.value = m_AppSummaryXML.GetAttribute("ELIGIBILITYINDICATOR_TEXT");
		frmScreen.txtDelphiTotal.value = m_AppSummaryXML.GetAttribute("DELPHITOTAL");
		
	}
	
	if(m_AppSummaryXML.SelectTag(null, "INTRODUCERS") != null)
	{
		frmScreen.txtPackagerApplicationRef.value = m_AppSummaryXML.GetAttribute("INGESTIONAPPLICATIONNUMBER");
	}
	
	if(m_AppSummaryXML.SelectTag(null, "MORTGAGEDETAILS") != null)
	{
		frmScreen.txtTotalLoanAmount.value = m_AppSummaryXML.GetAttribute("TOTALLOANAMOUNT");
		frmScreen.txtAmountRequested.value = m_AppSummaryXML.GetAttribute("AMOUNTREQUESTED");
		frmScreen.txtPurchasePrice.value = m_AppSummaryXML.GetAttribute("PURCHASEPRICE");
		frmScreen.txtLTV.value = m_AppSummaryXML.GetAttribute("LTV");
		frmScreen.txtTotalCost.value = m_AppSummaryXML.GetAttribute("TOTALNETMONTHLYCOST");
	}
		
	scScreenFunctions.SetRadioGroupState(frmScreen, "CreditCheckOptOutIndicator", "R");
	
	var dtApplicationRecommendedDate = m_AppSummaryXML.GetAttribute("APPLICATIONRECOMMENDEDDATE");
	var dtApplicationApprovalDate = m_AppSummaryXML.GetAttribute("APPLICATIONAPPROVALDATE");
	
	if (m_sUnderReview == 1)
		frmScreen.txtApplicationStatus.value = 'Under Review'
	else
	{
		if (m_sStageId != 910 || m_sStageId != 920)
		{
			if (dtApplicationRecommendedDate!= null && dtApplicationApprovalDate == null)
				frmScreen.txtApplicationStatus.value = 'Recommended';
			else if (dtApplicationApprovalDate != null)
				frmScreen.txtApplicationStatus.value = 'Approved';
			else
				frmScreen.txtApplicationStatus.value = 'Active';
		}
		else
		{
			if (m_sStageId == 910)
				frmScreen.txtApplicationStatus.value = 'Cancelled';
			else if (m_sStageId == 920)
				frmScreen.txtApplicationStatus.value = 'Declined';
		}
	}
	
}


function spnTblApplicants.onclick()
{
	PopulateListBox();
	PopulateDetails();	
}

function frmScreen.btnReview.onclick()
{
	frmToAP011.submit();
	
}

<% /* PSC 16/03/2007 EP2_1726 - Start */ %>
function frmScreen.btnDetail.onclick()
{
	
	var ArrayArguments = new Array(2); 
	var iRowSelected = scScrollTableIntroducer.getRowSelectedIndex()  //SR EP2_115
	var IntroducerDetailXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	IntroducerDetailXML.CreateRequestTagFromArray(m_RequestArray, null);
	var xe = m_AppSummaryXML.XMLDocument.selectNodes("//INTERMEDIARY")[iRowSelected-1];
	IntroducerDetailXML.XMLDocument.documentElement.appendChild(xe.cloneNode(true));

	ArrayArguments[0] = IntroducerDetailXML.XMLDocument.xml;

	scScreenFunctions.DisplayPopup(window, document, "dc011.asp", ArrayArguments, 420, 570);

	return;	
}
<% /* PSC 16/03/2007 EP2_1726 - End */ %>

function btnSubmit.onclick()
{
	window.close();
}

function frmScreen.btnContact.onclick()
{
	var iRowSelected = scScrollTableApplicants.getOffset() + (scScrollTableApplicants.getRowSelected());
	var iTableRowIndex = (tblApplicants.rows(iRowSelected).getAttribute("TagListItem"))
	
	switch (iTableRowIndex)
	{
		case 0: //1st applicant
			m_sCurrentCustomerNumber = m_sCustomer1Number;
			m_sCurrentCustomerVersionNumber = m_sCustomer1Version;
			m_sCurrentCustomerName = m_sCustomerName1;
			break;
		case 1: //2nd applicant
			m_sCurrentCustomerNumber = m_sCustomer2Number;
			m_sCurrentCustomerVersionNumber = m_sCustomer2Version;
			m_sCurrentCustomerName = m_sCustomerName2;
			break;
	
		case 2: //3rd applicant
			m_sCurrentCustomerNumber = m_sCustomer3Number;
			m_sCurrentCustomerVersionNumber = m_sCustomer3Version;		
			m_sCurrentCustomerName = m_sCustomerName3;
			break;
				
		case 3: //4th applicant
			m_sCurrentCustomerNumber = m_sCustomer4Number;
			m_sCurrentCustomerVersionNumber = m_sCustomer4Version;
			m_sCurrentCustomerName = m_sCustomerName4;
			break;
			
		case 4: //5th applicant
			m_sCurrentCustomerNumber = m_sCustomer5Number;
			m_sCurrentCustomerVersionNumber = m_sCustomer5Version;	
			m_sCurrentCustomerName = m_sCustomerName5;
			break;
				
		default: // Error
			var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			XML.SetErrorResponse();
			XML = null;
	}
								
	GetCustomer();

	var sReturn = null;
	var ArrayArguments = m_RequestArray;
	
	ArrayArguments[0] = m_RequestArray[0];
	ArrayArguments[1] = m_RequestArray[1];
	ArrayArguments[2] = m_RequestArray[2];
	ArrayArguments[3] = m_RequestArray[3];
	ArrayArguments[4] = m_RequestArray[4];
	ArrayArguments[5] = m_RequestArray[5];
	ArrayArguments[4] = (m_blnReadOnly || m_sReadOnly);
	ArrayArguments[5] = CustomerXML.XMLDocument.xml;
	ArrayArguments[6] = m_sCurrentCustomerName;
	
	sReturn = scScreenFunctions.DisplayPopup(window, document, "DC032.asp", ArrayArguments, 630, 395);

	if(sReturn != null)
	{
		FlagChange(sReturn[0]);
		var sXML = sReturn[1];
		CustomerXML.LoadXML(sXML);
	}
}

function frmScreen.btnContactThirdParty.onclick()
{
// Display popup DC241 - Contact Details
// =====================================
var sReturn = null;
var ArrayArguments = new Array(2);

	/* Ensure the correct xml is passed to TP in the form: 
	<CONTACTDETAILS><CONTACTTELEPHONEDETAILS/></CONTACTDETAILS>*/ 
	
	var sContactXML = GetContactDetails();
	var nStart = scScrollTableThirdParty.getOffset() + (scScrollTableThirdParty.getRowSelected() - 1);
	
	//if at least 1 third party row exists, then attempt to display contact details
	if (nStart > -1)
	{
		ArrayArguments[0] = 1; //read-only
		ArrayArguments[1] = sContactXML;
		sReturn = scScreenFunctions.DisplayPopup(window, document, "dc241.asp", ArrayArguments, 450, 290);
	}
	//DC241: NOTE: SetWorkExtState() - resets all tel numbers if usage is not populated!
}

function GetCustomer()
{
	CustomerXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	with (CustomerXML)
	{
		CustomerXML.CreateRequestTagFromArray(m_RequestArray, "");
		CreateActiveTag("CUSTOMERVERSION");
		CreateTag("CUSTOMERNUMBER", m_sCurrentCustomerNumber);
		CreateTag("CUSTOMERVERSIONNUMBER", m_sCurrentCustomerVersionNumber);
		RunASP(document, "GetPersonalDetails.asp");
		IsResponseOK();
	}
}


function GetContactDetails()
{
	//alert(m_AppSummaryXML.xml);
	//alert(m_AppSummaryXML.XMLDocument.xml);
	//alert(m_AppSummaryXML.ActiveTagList.length);

	var nStart = scScrollTableThirdParty.getOffset() + (scScrollTableThirdParty.getRowSelected() - 1);
	var sContactXML = '';
	
	XMLTp = new ActiveXObject("microsoft.XMLDOM");
	XMLTp = m_AppSummaryXML.XMLDocument;
	
	//if at least 1 third party row exists, then attempt to lookup contact details
	if (nStart > -1)
	{
		ContactXML = XMLTp.selectSingleNode('RESPONSE/APPLICATIONTHIRDPARTYLIST/APPLICATIONTHIRDPARTY[' + nStart + ']/THIRDPARTY/CONTACTDETAILS');
		sContactXML = ContactXML.xml;
	}
	
	return sContactXML;
	//alert(m_ContactXML.XMLDocument.xml);
}

<% /* PSC 16/03/2007 EP2_1726 - Start */ %>
function GetComboData()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var sGroupList = new Array("ListingStatus", "IntroducerSearchType", "FirmStatus");
	
	if(XML.GetComboLists(document,sGroupList))
	{
		m_XMLIntroducerSearchType = XML.GetComboListXML("IntroducerSearchType");
		m_XMLListingStatus = XML.GetComboListXML("ListingStatus");
		m_XMLFirmStatus = XML.GetComboListXML("FirmStatus");
	}
}
<% /* PSC 16/03/2007 EP2_1726 - End */ %>				
-->
</script>

<DIV></DIV>
</body>
</html>


