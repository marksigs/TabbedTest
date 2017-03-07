<html>

<comment>
Workfile:      OmigaLogon.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Omiga 4 logon screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date     Description
RF		10/11/99 Created
GD		20/12/00 SYS1742:File back in use. Changed name of form and content of form
				 - 13 basic variables.
MV		05/03/01	SYS2001 Commented AccessAuditGUID
RF		13/11/01 SYS2927 Added AdminSystemState.
HMA     02/02/04 BMIDS678  Added context parameter CreditCheckAccess
PSC     17/10/05 MAR57 Added context parameter LaunchXML
HMA     25/11/05 MAR324 Added context parameter AllowExitFromWrap
AW      30/10/06 EP1240 Added context parameter IsUserQualityChecker, IsUserFulfillApproved
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
</comment>
<%@  Language=JavaScript %>
<body>
<FORM id=frmDetails>
<P>

	Debug<input id="txtDebug" name="Debug"> </P>
<P>

	User ID<input id="txtUserId" name="UserId" value=DUMMYDATA> </P>
<P> 

	User Name<input id="txtUserName" name="UserName"> </P>
<P>

	Access Type<input id="txtAccessType" name="AccessType"> </P>
<P>

	User Competency<input id="txtUserCompetency" name="UserCompetency"> </P>
<P> 

	Qual List XML<input id="txtQualificationsListXml" name="QualificationsListXml"> </P>
<P>

	Role<input id="txtRole" name="Role"> </P>
<P> 

	Unit ID<input id="txtUnitId" name="UnitId"> </P>
<P>

	Unit Name<input id="txtUnitName" name="UnitName"> </P>
	
<P>
	Department ID<input id="txtDepartmentId" name="DepartmentId"> </P>
	
<P>
	Department Name*<input id="txtDepartmentName" name="DepartmentName"> </P>
	
<P>
	Dist Channel ID<input id="txtDistributionChannelId" name="DistributionChannelId"> </P>
	
<P>
	Dist Channel Name*<input id="txtDistributionChannelName" name="DistributionChannelName"> </P>
	
<P>
	Processing Indic<input id="txtProcessingIndicator" name="ProcessingIndicator"> </P>

<P>
	Machine ID<input id="txtMachineId" name="MachineId"> </P>

<% // <P> Access Audit GUID <input id = "txtAccessAuditGUID" name = "AccessAuditGUID" > </P> %>

<P>
	Admin System State<input id="txtAdminSystemState" name="AdminSystemState"> </P>

<% // Add Credit Check Access %>
<P>
	Credit Check Access<input id="txtCreditCheckAccess" name="CreditCheckAccess"> </P>
<% // PSC 17/10/2005 MAR57 - Add Launch XML %>
<P>
	Launch XML<input id="txtLaunchXML" name="LaunchXML"> </P>

<% // MAR324 - Add AllowExitFromWrap %>
<P>
	Allow Exit From Wrap<input id="txtAllowExitFromWrap" name="AllowExitFromWrap"> </P>
<P>
	Is user an authorised Quality Checker<input id="txtIsUserQualityChecker" name="IsUserQualityChecker"></P>

<P>
	Is user approved to fulfill documents<input id="txtIsUserFulfillApproved" name="IsUserFulfillApproved"></P>
</FORM>
</body>
</html>
