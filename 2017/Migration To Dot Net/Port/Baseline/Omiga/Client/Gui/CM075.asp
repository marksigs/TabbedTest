<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      CM075.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   View Subquotes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		22/03/00	Created - not complete
JLD		13/04/00	functionality completed
AY		14/04/00	New top menu/scScreenFunctions change
JLD		20/04/00	SYS0648, SYS0638, SYS0644(part1) - minor tweaks
JLD		04/05/00	Display the TOTALLOANCOMPONENTAMOUNT in the loancomponent listbox
					instead of LOANAMOUNT to show any added one off costs
MS		07/06/00	SYS0810 standardise on Sub-Quote, Rate type description other 
					display issues addressed.
CL		05/03/01	SYS1920 Read only functionality added
LD		23/05/02	SYS4727 Use cached versions of frame functions

BG		28/06/02	SYS4767 MSMS/Core integration
	JLD		16/04/02	MSMS0029	include ResolvedRate	
BG		28/06/02	SYS4767 MSMS/Core integration - END
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

BMIDS History:

Prog	Date		Description
GD		16/08/02	BMIDS00312
HMA     23/06/04    BMIDS776  Change heading from 'Monthly Premium' to 'Premium' for Buildings and Contents.

MARS HISTORY:
Prog	Date		Description
JD		09/03/2006	MAR1061 addition of purchaseprice to display
JD		22/03/2006	MAR1061 bug fix
*/
%>
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
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<span style="LEFT: 310px; POSITION: absolute; TOP: 278px">
<OBJECT id=scScrollTable style="LEFT: 0px; WIDTH: 304px; TOP: 0px; HEIGHT: 24px" 
tabIndex=-1 data=scTableListScroll.asp type=text/x-scriptlet 
VIEWASTEXT></OBJECT>
</span>
<OBJECT id=scLoanTable 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
data=scTableListScroll.asp type=text/x-scriptlet VIEWASTEXT></OBJECT>

<% /* Specify Forms Here */ %>
<form id="frmToCM020" method="post" action="CM020.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCM040" method="post" action="CM040.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCM100" method="post" action="CM100.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div id="divDetails" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 247px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<font face="MS Sans Serif" size="1">
			<strong>Sub-Quote Details and Costs</strong>
		</font>
	</span>
	<span id="spnTables" style="LEFT: 0px; POSITION: absolute; TOP: 26px">
		<span id="spnDetailsMTable" style="LEFT: 4px; POSITION: absolute; TOP: 0px">
			<table id="tblMTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
				<tr id="rowTitles">	<td width="5%" class="TableHead">Sub-Quote Number</td>	
									<td width="15%" class="TableHead">Total Loan Amount</td>		
									<td width="15%" class="TableHead">Purchase Price</td>
									<td width="10%" class="TableHead">LTV %</td>
									<td width="10%" class="TableHead">Number of Components</td>
									<td width="20%" class="TableHead">Total Net Monthly Cost</td>
									<td width="20%" class="TableHead">Total Net Charges</td>		
									<td class="TableHead">Active Quote?</td></tr>
				<tr id="row01">		<td width="5%" class="TableTopLeft">&nbsp;</td>		<td width="15%" class="TableTopCenter">&nbsp;</td>		<td width="15%" class="TableTopCenter">&nbsp;</td> <td width="10%" class="TableTopCenter">&nbsp;</td> <td width="10%" class="TableTopCenter">&nbsp;</td> <td width="20%" class="TableTopCenter">&nbsp;</td> <td width="20%" class="TableTopCenter">&nbsp;</td>		<td class="TableTopRight">&nbsp;</td></tr>
				<tr id="row02">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="15%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td>  <td width="10%" class="TableCenter">&nbsp;</td> <td width="20%" class="TableCenter">&nbsp;</td> <td width="20%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row03">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="15%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td> <td width="20%" class="TableCenter">&nbsp;</td> <td width="20%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr> 
				<tr id="row04">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="15%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td> <td width="20%" class="TableCenter">&nbsp;</td> <td width="20%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row05">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="15%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td> <td width="20%" class="TableCenter">&nbsp;</td> <td width="20%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row06">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="15%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td> <td width="20%" class="TableCenter">&nbsp;</td> <td width="20%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row07">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="15%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td> <td width="20%" class="TableCenter">&nbsp;</td> <td width="20%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row08">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="15%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td> <td width="20%" class="TableCenter">&nbsp;</td> <td width="20%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row09">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="15%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td> <td width="20%" class="TableCenter">&nbsp;</td> <td width="20%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row10">		<td width="5%" class="TableBottomLeft">&nbsp;</td>	<td width="15%" class="TableBottomCenter">&nbsp;</td> <td width="15%" class="TableBottomCenter">&nbsp;</td> <td width="10%" class="TableBottomCenter">&nbsp;</td> <td width="10%" class="TableBottomCenter">&nbsp;</td> <td width="20%" class="TableBottomCenter">&nbsp;</td>	<td width="20%" class="TableBottomCenter">&nbsp;</td>	<td class="TableBottomRight">&nbsp;</td></tr>
			</table>
		</span>
		<span id="spnDetailsBCTable" style="LEFT: 4px; POSITION: absolute; TOP: 0px DISPLAY: none">
			<table id="tblBCTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
				<tr id="rowTitles">	<td width="25%" class="TableHead">Sub-Quote Number</td>	
									<!--<td width="30%" class="TableHead">Product Name</td>		-->
									<td width="30%" class="TableHead">Cover Type</td>
									<td width="20%" class="TableHead">Reference Number</td>
									<td width="20%" class="TableHead">Premium</td>
									<!--<td width="20%" class="TableHead">Premium Frequency</td>-->
									<td class="TableHead">Active Quote?</td></tr>
				<tr id="row01">		<td width="25%" class="TableTopLeft">&nbsp;</td>	<td width="30%" class="TableTopCenter">&nbsp;</td>		<td width="20%" class="TableTopCenter">&nbsp;</td>  <td width="20%" class="TableTopCenter">&nbsp;</td> 	<td class="TableTopRight">&nbsp;</td></tr>
				<tr id="row02">		<td width="25%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>   <td width="20%" class="TableCenter">&nbsp;</td> 		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row03">		<td width="25%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	 <td width="20%" class="TableCenter">&nbsp;</td> 		<td class="TableRight">&nbsp;</td></tr> 
				<tr id="row04">		<td width="25%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	 <td width="20%" class="TableCenter">&nbsp;</td> 		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row05">		<td width="25%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	 <td width="20%" class="TableCenter">&nbsp;</td> 		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row06">		<td width="25%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	 <td width="20%" class="TableCenter">&nbsp;</td> 		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row07">		<td width="25%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	 <td width="20%" class="TableCenter">&nbsp;</td> 		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row08">		<td width="25%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	 <td width="20%" class="TableCenter">&nbsp;</td> 		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row09">		<td width="25%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	 <td width="20%" class="TableCenter">&nbsp;</td> 		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row10">		<td width="25%" class="TableBottomLeft">&nbsp;</td>	<td width="30%" class="TableBottomCenter">&nbsp;</td>   <td width="20%" class="TableBottomCenter">&nbsp;</td>  <td width="20%" class="TableBottomCenter">&nbsp;</td>		<td class="TableBottomRight">&nbsp;</td></tr>
			</table>
		</span>
		<span id="spnDetailsPPTable" style="LEFT: 4px; POSITION: absolute; TOP: 0px DISPLAY: none">
			<table id="tblPPTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
				<tr id="rowTitles">	<td width="25%" class="TableHead">Sub-Quote Number</td>	
									<!--<td width="30%" class="TableHead">Product Name</td>		-->
									<td width="30%" class="TableHead">Cover Type</td>
									<td width="20%" class="TableHead">Reference Number</td>
									<td width="20%" class="TableHead">Monthly Premium</td>
									<!--<td width="20%" class="TableHead">Premium Frequency</td>-->
									<td class="TableHead">Active Quote?</td></tr>
				<tr id="row01">		<td width="25%" class="TableTopLeft">&nbsp;</td>	<td width="30%" class="TableTopCenter">&nbsp;</td>		<td width="20%" class="TableTopCenter">&nbsp;</td>  <td width="20%" class="TableTopCenter">&nbsp;</td> 	<td class="TableTopRight">&nbsp;</td></tr>
				<tr id="row02">		<td width="25%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>   <td width="20%" class="TableCenter">&nbsp;</td> 		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row03">		<td width="25%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	 <td width="20%" class="TableCenter">&nbsp;</td> 		<td class="TableRight">&nbsp;</td></tr> 
				<tr id="row04">		<td width="25%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	 <td width="20%" class="TableCenter">&nbsp;</td> 		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row05">		<td width="25%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	 <td width="20%" class="TableCenter">&nbsp;</td> 		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row06">		<td width="25%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	 <td width="20%" class="TableCenter">&nbsp;</td> 		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row07">		<td width="25%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	 <td width="20%" class="TableCenter">&nbsp;</td> 		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row08">		<td width="25%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	 <td width="20%" class="TableCenter">&nbsp;</td> 		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row09">		<td width="25%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	 <td width="20%" class="TableCenter">&nbsp;</td> 		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row10">		<td width="25%" class="TableBottomLeft">&nbsp;</td>	<td width="30%" class="TableBottomCenter">&nbsp;</td>   <td width="20%" class="TableBottomCenter">&nbsp;</td>  <td width="20%" class="TableBottomCenter">&nbsp;</td>		<td class="TableBottomRight">&nbsp;</td></tr>
			</table>
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 218px">
		<input id="btnReinstate" value="Reinstate" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
</div>	
<div id="divLoanComponents" style="LEFT: 10px; VISIBILITY: hidden; WIDTH: 604px; POSITION: absolute; TOP: 310px; HEIGHT: 240px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 5px" class="msgLabel">
		<font face="MS Sans Serif" size="1">
			<strong>Loan Components</strong>
		</font>
	</span>	
	<span id="spnLoanTable" style="LEFT: 4px; POSITION: absolute; TOP: 26px">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<!--BG		28/06/02	SYS4767 MSMS/Core integration
				MO		01/07/02	BMIDS000131 Moved from the "sub-quote Details and Costs" table
									into the right place
			-->
				<tr id="rowTitles">	<td width="10%" class="TableHead">Prod. Number</td>	
									<td width="25%" class="TableHead">Prod. Name</td>		
									<td width="15%" class="TableHead">Rate Type</td>
									<td width="8%" class="TableHead">Rate</td>
									<td width="8%" class="TableHead">Resolved Rate</td>
									<td width="10%" class="TableHead">Loan Amount</td>
									<td width="16%" class="TableHead">Term (Yrs & Mnths)</td>
									<td width="8%" class="TableHead">Repay. Type</td>		
									<td class="TableHead">Monthly Cost</td></tr>
				<tr id="row01">		<td width="10%" class="TableTopLeft">&nbsp;</td>		<td width="25%" class="TableTopCenter">&nbsp;</td>		<td width="15%" class="TableTopCenter">&nbsp;</td> <td width="8%" class="TableTopCenter">&nbsp;</td> <td width="8%" class="TableTopCenter">&nbsp;</td> <td width="10%" class="TableTopCenter">&nbsp;</td> <td width="16%" class="TableTopCenter">&nbsp;</td>	<td width="8%" class="TableTopCenter">&nbsp;</td>	<td class="TableTopRight">&nbsp;</td></tr>
				<tr id="row02">		<td width="10%" class="TableLeft">&nbsp;</td>		<td width="25%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>  <td width="8%" class="TableCenter">&nbsp;</td> <td width="8%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td> <td width="16%" class="TableCenter">&nbsp;</td> <td width="8%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row03">		<td width="10%" class="TableLeft">&nbsp;</td>		<td width="25%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>	<td width="8%" class="TableCenter">&nbsp;</td>  <td width="8%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td> <td width="16%" class="TableCenter">&nbsp;</td> <td width="8%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr> 
				<tr id="row04">		<td width="10%" class="TableLeft">&nbsp;</td>		<td width="25%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>	<td width="8%" class="TableCenter">&nbsp;</td>  <td width="8%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td> <td width="16%" class="TableCenter">&nbsp;</td> <td width="8%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row05">		<td width="10%" class="TableLeft">&nbsp;</td>		<td width="25%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>	<td width="8%" class="TableCenter">&nbsp;</td>  <td width="8%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td> <td width="16%" class="TableCenter">&nbsp;</td> <td width="8%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row06">		<td width="10%" class="TableLeft">&nbsp;</td>		<td width="25%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>	<td width="8%" class="TableCenter">&nbsp;</td>  <td width="8%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td> <td width="16%" class="TableCenter">&nbsp;</td> <td width="8%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row07">		<td width="10%" class="TableLeft">&nbsp;</td>		<td width="25%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>	<td width="8%" class="TableCenter">&nbsp;</td>  <td width="8%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td> <td width="16%" class="TableCenter">&nbsp;</td> <td width="8%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row08">		<td width="10%" class="TableLeft">&nbsp;</td>		<td width="25%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>	<td width="8%" class="TableCenter">&nbsp;</td>  <td width="8%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td> <td width="16%" class="TableCenter">&nbsp;</td> <td width="8%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row09">		<td width="10%" class="TableLeft">&nbsp;</td>		<td width="25%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>	<td width="8%" class="TableCenter">&nbsp;</td>  <td width="8%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td> <td width="16%" class="TableCenter">&nbsp;</td> <td width="8%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
				<tr id="row10">		<td width="10%" class="TableBottomLeft">&nbsp;</td>	<td width="25%" class="TableBottomCenter">&nbsp;</td> <td width="15%" class="TableBottomCenter">&nbsp;</td> <td width="8%" class="TableBottomCenter">&nbsp;</td> <td width="8%" class="TableBottomCenter">&nbsp;</td> <td width="10%" class="TableBottomCenter">&nbsp;</td>	<td width="16%" class="TableBottomCenter">&nbsp;</td> <td width="8%" class="TableBottomCenter">&nbsp;</td>	<td class="TableBottomRight">&nbsp;</td></tr>

			<!--<tr id="rowTitles">	<td width="10%" class="TableHead">Product Number</td>	
								<td width="30%" class="TableHead">Product Name</td>		
								<td width="15%" class="TableHead">Rate Type</td>
								<td width="10%" class="TableHead">Rate</td>
								<td width="10%" class="TableHead">Loan Amount</td>
								<td width="15%" class="TableHead">Term (Years and Months)</td>
								<td width="10%" class="TableHead">Repayment Type</td>		
								<td class="TableHead">Monthly Cost</td></tr>
			<tr id="row01">		<td width="10%" class="TableTopLeft">&nbsp;</td>		<td width="30%" class="TableTopCenter">&nbsp;</td>		<td width="15%" class="TableTopCenter">&nbsp;</td> <td width="10%" class="TableTopCenter">&nbsp;</td> <td width="10%" class="TableTopCenter">&nbsp;</td> <td width="15%" class="TableTopCenter">&nbsp;</td>	<td width="10%" class="TableTopCenter">&nbsp;</td>	<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		<td width="10%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>  <td width="10%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td> <td width="15%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		<td width="10%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td> <td width="15%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr> 
			<tr id="row04">		<td width="10%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td> <td width="15%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		<td width="10%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td> <td width="15%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		<td width="10%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td> <td width="15%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		<td width="10%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td> <td width="15%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		<td width="10%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td> <td width="15%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		<td width="10%" class="TableLeft">&nbsp;</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td> <td width="15%" class="TableCenter">&nbsp;</td> <td width="10%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		<td width="10%" class="TableBottomLeft">&nbsp;</td>	<td width="30%" class="TableBottomCenter">&nbsp;</td> <td width="15%" class="TableBottomCenter">&nbsp;</td> <td width="10%" class="TableBottomCenter">&nbsp;</td> <td width="10%" class="TableBottomCenter">&nbsp;</td>	<td width="15%" class="TableBottomCenter">&nbsp;</td> <td width="10%" class="TableBottomCenter">&nbsp;</td>	<td class="TableBottomRight">&nbsp;</td></tr>-->
			<!-- BG,MO, MSMS Core Integration and BMIDS000131 END -->
		</table>
	</span>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 555px; HEIGHT: 19px">
<!-- #include FILE="msgButtons.asp" -->
</div>


<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/CM075Attribs.asp" -->


<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sBusinessType = "";
var m_sApplicationNumber = "";
var m_sApplicationFFNumber = "";
var m_sApplicationMode = "";	
var m_sReadOnly = "";
var m_tblDetailsTable = null;
var subQuoteXML = null;
var m_iNumOfSubQuotes = 0;
var m_iTableLength = 10;
var m_iLoanTableLength = 15;
var m_sIdMortgageSubQuoteNumber = "";
var m_sIdBCSubQuoteNumber = "";
var m_sIdPPSubQuoteNumber = "";
var m_nCurrentActiveRow = -1;
var scScreenFunctions;
var m_blnReadOnly = false;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("View Sub-Quotes","CM075",scScreenFunctions);

	RetrieveContextData();
<%	//This function is contained in the field attributes file (remove if not required)
%>	//not required SetMasks();
	Validation_Init();
	Initialise();
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "CM075");
	AdjustMSGButtons();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
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
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sBusinessType = scScreenFunctions.GetContextParameter(window,"idBusinessType",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sApplicationFFNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);	
	m_sApplicationMode = scScreenFunctions.GetContextParameter(window,"idApplicationMode",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sIdMortgageSubQuoteNumber = scScreenFunctions.GetContextParameter(window,"idMortgageSubQuoteNumber",null);
	m_sIdBCSubQuoteNumber = scScreenFunctions.GetContextParameter(window,"idBCSubQuoteNumber",null);
	m_sIdPPSubQuoteNumber = scScreenFunctions.GetContextParameter(window,"idPPSubQuoteNumber",null);
	m_sIdQuotationNumber = scScreenFunctions.GetContextParameter(window,"idQuotationNumber",null);
}

function Initialise()
{
	if(m_sReadOnly == "1")
	{
		scScreenFunctions.SetCollectionState(spnTables, "R");
		frmScreen.btnReinstate.disabled = true;
	}
	if(m_sBusinessType == "Mortgage") // CHANGE TO 'M'
	{
		m_tblDetailsTable = tblMTable;
		scScreenFunctions.HideCollection(spnDetailsBCTable);
		scScreenFunctions.HideCollection(spnDetailsPPTable);
		scScreenFunctions.ShowCollection(spnDetailsMTable);
		scScreenFunctions.ShowCollection(divLoanComponents);
	}
	else if (m_sBusinessType == "BC")
	{
		m_tblDetailsTable = tblBCTable;
		scScreenFunctions.HideCollection(spnDetailsMTable);
		scScreenFunctions.HideCollection(spnDetailsPPTable);
		scScreenFunctions.ShowCollection(spnDetailsBCTable);		
	}
	else if (m_sBusinessType == "PP")
	{
		m_tblDetailsTable = tblPPTable;
		scScreenFunctions.HideCollection(spnDetailsBCTable);
		scScreenFunctions.HideCollection(spnDetailsMTable);
		scScreenFunctions.ShowCollection(spnDetailsPPTable);		
	}
	subQuoteXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	subQuoteXML.CreateRequestTag(window,null);
	subQuoteXML.CreateActiveTag("QUOTATION");
	subQuoteXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	subQuoteXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	subQuoteXML.CreateTag("BUSINESSTYPE", m_sBusinessType);
	// CHANGE TO 'M'
	if(m_sBusinessType == "Mortgage") subQuoteXML.CreateTag("ACTIVEQUOTENUMBER", m_sIdMortgageSubQuoteNumber);
	else if(m_sBusinessType == "BC") subQuoteXML.CreateTag("BCSUBQUOTENUMBER", m_sIdBCSubQuoteNumber);
	else if(m_sBusinessType == "PP") subQuoteXML.CreateTag("PPSUBQUOTENUMBER", m_sIdPPSubQuoteNumber);
	subQuoteXML.RunASP(document,"FindSubQuotes.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = subQuoteXML.CheckResponse(ErrorTypes);

	if (ErrorReturn[0] == true)PopulateDetailsList(0);
	else if(ErrorReturn[1] == ErrorTypes[0]) alert("No Sub-Quotes found for this application.");
	ErrorTypes = null;
}
function PopulateDetailsList(nStart)
{
<%  /* Populate the listbox with values from subQuoteXML */
%>	subQuoteXML.ActiveTag = null;
	if(m_sBusinessType == "Mortgage")	subQuoteXML.CreateTagList("MORTGAGESUBQUOTE");// CHANGE TO 'M'
	else if(m_sBusinessType == "BC")	subQuoteXML.CreateTagList("BUILDINGSANDCONTENTSSUBQUOTE");
	else subQuoteXML.CreateTagList("PAYMENTPROTECTIONSUBQUOTE");
	m_iNumOfSubQuotes = subQuoteXML.ActiveTagList.length;
	scScrollTable.initialiseTable(m_tblDetailsTable, 0, "", ShowList, m_iTableLength, subQuoteXML.ActiveTagList.length);
	ShowList(nStart);
	//if(m_iNumOfSubQuotes > 0) scScrollTable.setRowSelected(1);
}
function ShowList(nStart)
{
	scScrollTable.clear();
	subQuoteXML.ActiveTag = null;
	if(m_sBusinessType == "Mortgage")	subQuoteXML.CreateTagList("MORTGAGESUBQUOTE");// CHANGE TO 'M'
	else if(m_sBusinessType == "BC")	subQuoteXML.CreateTagList("BUILDINGSANDCONTENTSSUBQUOTE");
	else subQuoteXML.CreateTagList("PAYMENTPROTECTIONSUBQUOTE");
	var currentTagList = subQuoteXML.ActiveTagList;
	for (var iCount = 0; iCount < subQuoteXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		subQuoteXML.SelectTagListItem(iCount + nStart);
		
		if(m_sBusinessType == "Mortgage")// CHANGE TO 'M'
		{
			scScreenFunctions.SizeTextToField(tblMTable.rows(iCount+1).cells(0),subQuoteXML.GetTagText("MORTGAGESUBQUOTENUMBER"));
			scScreenFunctions.SizeTextToField(tblMTable.rows(iCount+1).cells(1),subQuoteXML.GetTagText("TOTALLOANAMOUNT"));
			scScreenFunctions.SizeTextToField(tblMTable.rows(iCount+1).cells(2),subQuoteXML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE")); //mar1061
			scScreenFunctions.SizeTextToField(tblMTable.rows(iCount+1).cells(3),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("LTV"),2));
			scScreenFunctions.SizeTextToField(tblMTable.rows(iCount+1).cells(5),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("TOTALNETMONTHLYCOST"), 2));
			scScreenFunctions.SizeTextToField(tblMTable.rows(iCount+1).cells(6),subQuoteXML.GetTagText("TOTALNETCHARGES"));
			if(subQuoteXML.GetTagText("MORTGAGESUBQUOTENUMBER") == m_sIdMortgageSubQuoteNumber)
			{
				scScreenFunctions.SizeTextToField(tblMTable.rows(iCount+1).cells(7),"Yes");
				m_nCurrentActiveRow = iCount + 1;
			}
<%          /* now work out the number of loan components, remember to reset the ActiveTagList afterwards */
%>			subQuoteXML.CreateTagList("LOANCOMPONENT");
			var sNumOfLoanComps = subQuoteXML.ActiveTagList.length;
			scScreenFunctions.SizeTextToField(tblMTable.rows(iCount+1).cells(4),sNumOfLoanComps);
			subQuoteXML.ActiveTagList = currentTagList;
		}
		else if(m_sBusinessType == "BC")
		{
			scScreenFunctions.SizeTextToField(tblBCTable.rows(iCount+1).cells(0),subQuoteXML.GetTagText("BCSUBQUOTENUMBER"));
			//scScreenFunctions.SizeTextToField(tblBCTable.rows(iCount+1).cells(1),subQuoteXML.GetTagText("PRODUCTNAME")); // from BUILDINGANDCONTENTSPRODUCT
			scScreenFunctions.SizeTextToField(tblBCTable.rows(iCount+1).cells(1),subQuoteXML.GetTagAttribute("COVERTYPE", "TEXT"));
			scScreenFunctions.SizeTextToField(tblBCTable.rows(iCount+1).cells(2),subQuoteXML.GetTagText("BCREFERENCENUMBER"));
			
			scScreenFunctions.SizeTextToField(tblBCTable.rows(iCount+1).cells(3), top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("TOTALBCMONTHLYCOST"), 2) );
			//scScreenFunctions.SizeTextToField(tblBCTable.rows(iCount+1).cells(4),subQuoteXML.GetTagAttribute("FREQUENCY", "TEXT"));
			if(subQuoteXML.GetTagText("BCSUBQUOTENUMBER") == m_sIdBCSubQuoteNumber)
			{
				scScreenFunctions.SizeTextToField(tblBCTable.rows(iCount+1).cells(4),"Yes");
				m_nCurrentActiveRow = iCount + 1;
			} else
			{
				scScreenFunctions.SizeTextToField(tblBCTable.rows(iCount+1).cells(4),"No");
			
			}
		}
		else  // business Type = PP
		{
			//scScreenFunctions.SizeTextToField(tblPPTable.rows(iCount+1).cells(0),subQuoteXML.GetTagText("PPSUBQUOTENUMBER"));
			//var strApplicants = "";
			//if(subQuoteXML.GetTagText("MORTGAGECOVERFORAPPLICANT1") != "" && subQuoteXML.GetTagText("MORTGAGECOVERFORAPPLICANT1") != "0")
				//strApplicants = scScreenFunctions.GetContextParameter(window,"idCustomerName1",null);
			//if(subQuoteXML.GetTagText("MORTGAGECOVERFORAPPLICANT2") != "" && subQuoteXML.GetTagText("MORTGAGECOVERFORAPPLICANT2") != "0")
			//{
				//if(strApplicants != "")strApplicants = strApplicants + ", ";
				//strApplicants = strApplicants + scScreenFunctions.GetContextParameter(window,"idCustomerName2",null);
			//}
			//scScreenFunctions.SizeTextToField(tblPPTable.rows(iCount+1).cells(1),strApplicants);
			//var strCoverType = "Undefined";
			//if(subQuoteXML.GetTagText("COVERTYPE") == "1")strCoverType = "ASU";
			//else if(subQuoteXML.GetTagText("COVERTYPE") == "2")strCoverType = "AS";
			//else if(subQuoteXML.GetTagText("COVERTYPE") == "3")strCoverType = "U";
			scScreenFunctions.SizeTextToField(tblPPTable.rows(iCount+1).cells(0),subQuoteXML.GetTagText("PPSUBQUOTENUMBER"));
			//scScreenFunctions.SizeTextToField(tblBCTable.rows(iCount+1).cells(1),subQuoteXML.GetTagText("PRODUCTNAME")); // from BUILDINGANDCONTENTSPRODUCT
			scScreenFunctions.SizeTextToField(tblPPTable.rows(iCount+1).cells(1),subQuoteXML.GetTagAttribute("COVERTYPE", "TEXT"));
			scScreenFunctions.SizeTextToField(tblPPTable.rows(iCount+1).cells(2),subQuoteXML.GetTagText("PPREFERENCENUMBER"));
			
			scScreenFunctions.SizeTextToField(tblPPTable.rows(iCount+1).cells(3), top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("TOTALPPMONTHLYCOST"), 2) );
			//scScreenFunctions.SizeTextToField(tblBCTable.rows(iCount+1).cells(4),subQuoteXML.GetTagAttribute("FREQUENCY", "TEXT"));
			//scScreenFunctions.SizeTextToField(tblPPTable.rows(iCount+1).cells(2),strCoverType);
			//scScreenFunctions.SizeTextToField(tblPPTable.rows(iCount+1).cells(3),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("TOTALPPMONTHLYCOST"), 2));
			if(subQuoteXML.GetTagText("PPSUBQUOTENUMBER") == m_sIdPPSubQuoteNumber)
			{
				scScreenFunctions.SizeTextToField(tblPPTable.rows(iCount+1).cells(4),"Yes");
				m_nCurrentActiveRow = iCount + 1;
			} else
			{
				scScreenFunctions.SizeTextToField(tblPPTable.rows(iCount+1).cells(4),"No");
			
			}
		}
	}	
}
function spnTables.onclick()
{
<% /* SelectMortgageSubQuote */
%>	if (scScrollTable.getRowSelected() != -1)
	{
		if(m_sBusinessType == "Mortgage")// CHANGE TO 'M'
		{
			PopulateLoanCompList(0);
		}
	}
}
function PopulateLoanCompList(nStart)
{
<%  /* Populate the listbox with loan components from the selected row in subQuoteXML */
%>	var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	subQuoteXML.ActiveTag = null;			
	subQuoteXML.CreateTagList("MORTGAGESUBQUOTE");
	subQuoteXML.SelectTagListItem(nRowSelected -1);
	subQuoteXML.CreateTagList("LOANCOMPONENT");
	scLoanTable.initialiseTable(tblTable, 0, "", ShowLoanList, m_iLoanTableLength, subQuoteXML.ActiveTagList.length);
	scLoanTable.DisableTable();
	ShowLoanList(nStart);
}
function ShowLoanList(nStart)
{
	var sRateType = "";
	scLoanTable.clear();
	for (var iCount = 0; iCount < subQuoteXML.ActiveTagList.length && iCount < m_iLoanTableLength; iCount++)
	{
		subQuoteXML.SelectTagListItem(iCount + nStart);
		
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),subQuoteXML.GetTagText("PRODUCTNAME"));
		sRateType = subQuoteXML.GetTagText("RATETYPE")
		if (sRateType == "F") sRateType="Fixed";
		else if(sRateType == "C") sRateType="Capped";
		else if(sRateType == "D") sRateType="Discount";
		else if(sRateType == "B") sRateType="Base";
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),sRateType);
		
		<%/*BG		28/06/02	SYS4767 MSMS/Core integration*/%>
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("RATE"), 2));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("RESOLVEDRATE"), 2));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(5),subQuoteXML.GetTagText("TOTALLOANCOMPONENTAMOUNT"));
		var strYearsMonths = subQuoteXML.GetTagText("TERMINYEARS") + " Yrs ";
		if( parseInt(subQuoteXML.GetTagText("TERMINMONTHS")) > 0) strYearsMonths += subQuoteXML.GetTagText("TERMINMONTHS") + " Mths";
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(6),strYearsMonths);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(7),GetAbbrevRepayMethod(subQuoteXML.GetTagText("REPAYMENTMETHOD")));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(8),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("NETMONTHLYCOST"), 2));
		
		<%/*scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("RATE"), 2));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4),subQuoteXML.GetTagText("TOTALLOANCOMPONENTAMOUNT"));
		var strYearsMonths = subQuoteXML.GetTagText("TERMINYEARS") + " Yrs ";
		if( parseInt(subQuoteXML.GetTagText("TERMINMONTHS")) > 0) strYearsMonths += subQuoteXML.GetTagText("TERMINMONTHS") + " Mths";
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(5),strYearsMonths);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(6),subQuoteXML.GetTagAttribute("REPAYMENTMETHOD", "TEXT"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(7),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("NETMONTHLYCOST"), 2));*/%>
		<%/*BG		28/06/02	SYS4767 MSMS/Core integration - END*/%>		
		
		sRateType = ""
	}	
}

<%/*BG		28/06/02	SYS4767 MSMS/Core integration */%>	
function GetAbbrevRepayMethod(sRepayMethod)
{
<%	/* to condense the information in the listbox, output an abbreviated string */
%>	var sAbbrev;
	switch(sRepayMethod)
	{
		case "1" : sAbbrev = "I/O"; break;
		case "2" : sAbbrev = "C&I"; break;
		case "3" : sAbbrev = "P&P"; break;
		default : sAbbrev = "??"; break;
	}
	return sAbbrev;
}
<%/*BG		28/06/02	SYS4767 MSMS/Core integration - END*/%>	


function frmScreen.btnReinstate.onclick()
{
	if (scScrollTable.getRowSelected() != -1)
	{
		var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
		subQuoteXML.ActiveTag = null;
		
		var sActiveSubQuoteNumber = "";
		var sChosenQuoteNumber = "";
		var sOutput = "";
		var bAlreadyActive = false;
		var bInvalidSubQuote = false;
		var ValidateXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		ValidateXML.CreateRequestTag(window,null);
		ValidateXML.CreateActiveTag("BASICQUOTATIONDETAILS");
		ValidateXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		ValidateXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
		if(m_sBusinessType == "Mortgage")// CHANGE TO 'M'
		{
			subQuoteXML.CreateTagList("MORTGAGESUBQUOTE");
			subQuoteXML.SelectTagListItem(nRowSelected -1);
			if(subQuoteXML.GetTagText("MORTGAGESUBQUOTENUMBER") == m_sIdMortgageSubQuoteNumber)
				bAlreadyActive = true;
			else
			{
				sActiveSubQuoteNumber = m_sIdMortgageSubQuoteNumber;
				sChosenQuoteNumber = subQuoteXML.GetTagText("MORTGAGESUBQUOTENUMBER");
				if(m_sApplicationMode == "Cost Modelling")
				{
					sOutput = "VALIDMORTGAGESUBQUOTE";
					ValidateXML.CreateTag("MORTGAGESUBQUOTENUMBER",subQuoteXML.GetTagText("MORTGAGESUBQUOTENUMBER"));
					// 					ValidateXML.RunASP(document, "AQValidateMortgageSubQuote.asp"); 
					// Added by automated update TW 09 Oct 2002 SYS5115
					switch (ScreenRules())
						{
						case 1: // Warning
						case 0: // OK
											ValidateXML.RunASP(document, "AQValidateMortgageSubQuote.asp"); 
							break;
						default: // Error
							ValidateXML.SetErrorResponse();
						}

				}
			}
		}
		else if(m_sBusinessType == "BC")
		{
			subQuoteXML.CreateTagList("BUILDINGSANDCONTENTSSUBQUOTE");
			subQuoteXML.SelectTagListItem(nRowSelected -1);
			if(subQuoteXML.GetTagText("BCSUBQUOTENUMBER") == m_sIdBCSubQuoteNumber)
				bAlreadyActive = true;
			else
			{
				sActiveSubQuoteNumber = m_sIdBCSubQuoteNumber;
				sChosenQuoteNumber = subQuoteXML.GetTagText("BCSUBQUOTENUMBER");
				sOutput = "VALIDBCSUBQUOTE";
				ValidateXML.CreateActiveTag("BUILDINGSANDCONTENTSSUBQUOTE")
				ValidateXML.AddXMLBlock(subQuoteXML.CreateFragment());
				// 				ValidateXML.RunASP(document, "ValidateBCSubQuote.asp");
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
					{
					case 1: // Warning
					case 0: // OK
									ValidateXML.RunASP(document, "ValidateBCSubQuote.asp");
						break;
					default: // Error
						ValidateXML.SetErrorResponse();
					}

			}
		}
		else  // PP
		{
			subQuoteXML.CreateTagList("PAYMENTPROTECTIONSUBQUOTE");
			subQuoteXML.SelectTagListItem(nRowSelected -1);
			if(subQuoteXML.GetTagText("PPSUBQUOTENUMBER") == m_sIdPPSubQuoteNumber)
				bAlreadyActive = true;
			else
			{
				sActiveSubQuoteNumber = m_sIdPPSubQuoteNumber;
				sChosenQuoteNumber = subQuoteXML.GetTagText("PPSUBQUOTENUMBER");
				sOutput = "VALIDPPSUBQUOTE";
				ValidateXML.CreateActiveTag("PAYMENTPROTECTIONSUBQUOTE")
				ValidateXML.AddXMLBlock(subQuoteXML.CreateFragment());
				// 				ValidateXML.RunASP(document, "ValidatePPSubQuote.asp");
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
					{
					case 1: // Warning
					case 0: // OK
									ValidateXML.RunASP(document, "ValidatePPSubQuote.asp");
						break;
					default: // Error
						ValidateXML.SetErrorResponse();
					}

			}		
		}
		
		if(bAlreadyActive == true) alert("You cannot reinstate the currently active Sub-Quote");
		else
		{
<%			/* If business type is mortgage and applictionmode is Quick Quote we do not need 
			   to validate the subquote */
%>			var bQuoteIsValid = false;
			if(m_sBusinessType == "Mortgage" && m_sApplicationMode == "Quick Quote")bQuoteIsValid = true;
			else
			{
				if(ValidateXML.IsResponseOK())
				{
					ValidateXML.SelectTag(null, "RESPONSE");
					if(ValidateXML.GetTagText(sOutput) == "0") alert("The selected Sub-Quote is now invalid and therefore cannot be reinstated");
					else bQuoteIsValid = true;
				}
			}
			if(bQuoteIsValid)
			{
<%				/* We have a valid subquote selected so reinstate it */
%>				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				XML.CreateRequestTag(window,null);
				XML.CreateActiveTag("QUOTATION");
				XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
				XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
				XML.CreateTag("QUOTATIONNUMBER", m_sIdQuotationNumber);
				XML.CreateTag("ACTIVESUBQUOTENUMBER", sActiveSubQuoteNumber);
				XML.CreateTag("SELECTEDSUBQUOTENUMBER", sChosenQuoteNumber);
				XML.CreateTag("BUSINESSTYPE", m_sBusinessType);
				// 				XML.RunASP(document, "ReinstateSubQuote.asp");
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
					{
					case 1: // Warning
					case 0: // OK
									XML.RunASP(document, "ReinstateSubQuote.asp");
						break;
					default: // Error
						XML.SetErrorResponse();
					}

				if(XML.IsResponseOK())
				{
					if(m_sBusinessType == "Mortgage")// CHANGE TO 'M'
					{
						m_sIdMortgageSubQuoteNumber = sChosenQuoteNumber;
						scScreenFunctions.SetContextParameter(window,"idMortgageSubQuoteNumber",m_sIdMortgageSubQuoteNumber);
						if (m_nCurrentActiveRow != -1) scScreenFunctions.SizeTextToField(tblMTable.rows(m_nCurrentActiveRow).cells(7),"");
						scScreenFunctions.SizeTextToField(tblMTable.rows(nRowSelected).cells(7),"Yes");
					}
					else if(m_sBusinessType == "BC")
					{
						m_sIdBCSubQuoteNumber = sChosenQuoteNumber;
						scScreenFunctions.SetContextParameter(window,"idBCSubQuoteNumber",m_sIdBCSubQuoteNumber);
						if (m_nCurrentActiveRow != -1)scScreenFunctions.SizeTextToField(tblBCTable.rows(m_nCurrentActiveRow).cells(4),"No");
						scScreenFunctions.SizeTextToField(tblBCTable.rows(nRowSelected).cells(4),"Yes");
					}
					else
					{
						m_sIdPPSubQuoteNumber = sChosenQuoteNumber;
						scScreenFunctions.SetContextParameter(window,"idPPSubQuoteNumber",m_sIdPPSubQuoteNumber);
						if (m_nCurrentActiveRow != -1)scScreenFunctions.SizeTextToField(tblPPTable.rows(m_nCurrentActiveRow).cells(4),"No");
						scScreenFunctions.SizeTextToField(tblPPTable.rows(nRowSelected).cells(4),"Yes");
					}
					m_nCurrentActiveRow = nRowSelected;
				}
				XML = null;
			}
		}
	}
	else alert("Please select a Sub-Quote for reinstatement");
}
function btnSubmit.onclick()
{
	if(frmScreen.onsubmit())
	{
		if(m_sMetaAction == "CM020")frmToCM020.submit();
		else if(m_sMetaAction == "CM040")frmToCM040.submit();
		else if(m_sMetaAction == "CM100")frmToCM100.submit();
	}
}

function AdjustMSGButtons()
{
	var iTotalHt = 0;
	if (spnDetailsMTable.style.visibility =='visible') iTotalHt=iTotalHt + 215;
	if (spnDetailsBCTable.style.visibility =='visible') iTotalHt=iTotalHt + 215;
	if (spnDetailsPPTable.style.visibility =='visible') iTotalHt=iTotalHt + 215;	
	if (divLoanComponents.style.visibility =='visible') iTotalHt=iTotalHt + 215;		
	
	
	msgButtons.style.top = iTotalHt + 100;	
	


}

-->
</script>
</body>
</html>
<% /* OMIGA BUILD VERSION 046.02.08.08.00 */ %>



