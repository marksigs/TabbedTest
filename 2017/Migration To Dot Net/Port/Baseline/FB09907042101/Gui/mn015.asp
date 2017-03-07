<%@ Language=JScript %>
<HTML>
<HEAD>
<%
/*
Workfile:      MN015.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Task Search
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
APS		28/02/01	SYS1990 Added in screen name
CL		05/03/01	SYS1995 Changed images
CL		05/03/01	SYS1920 Read only functionality added
CL		12/03/01	SYS2031 Button re-sizing	
LD		23/05/02	SYS4727 Use cached versions of frame functions
	
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ 
%>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
</HEAD>
<BODY>

<form id="frmTaskSummary" method="post" action="tm010.asp" STYLE="DISPLAY: none"></form>
<form id="frmTaskSearch" method="post" action="tm020.asp" STYLE="DISPLAY: none"></form>
<form id="frmToMN010" method="post" action="mn010.asp" STYLE="DISPLAY: none"></form>

<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="HEIGHT: 345px; LEFT: 15px; POSITION: absolute; TOP: 60px; WIDTH: 595px" class="msgGroup">
	<span id="spnTaskSummary" style="LEFT: 25px; POSITION: absolute; TOP: 45px" class="msgLabel">					
		<img alt border="0" id="imgTaskSummary1" src="images/task_summary_1.jpg" style="LEFT:170px; POSITION: absolute; TOP: 10px; Z-INDEX: 2" WIDTH="200" HEIGHT="100">				
		<img alt border="0" id="imgTaskSummary2" src="images/task_summary_2.jpg" style="LEFT: 170px; POSITION: absolute; TOP: 10px; Z-INDEX: 1" WIDTH="200" HEIGHT="100"> 
	</span>				
	<span id="spnTaskSearch" style="LEFT: 25px; POSITION: absolute; TOP: 185px" class="msgLabel">					
		<img alt border="0" id="imgTaskSearch1" src="images/task_search_1.jpg" style="LEFT:170px; POSITION: absolute; TOP: 10px; Z-INDEX: 2" WIDTH="200" HEIGHT="100">				
		<img alt border="0" id="imgTaskSearch2" src="images/task_search_2.jpg" style="LEFT: 170px; POSITION: absolute; TOP: 10px; Z-INDEX: 1" WIDTH="200" HEIGHT="100"> 
	</span>
</form>
</div>

<% /* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" -->
</div>
<!-- #include FILE="fw030.asp" -->

<!-- Specify Code Here -->
<script language="JScript">
<!--

var m_blnReadOnly = false;


function window.onload() 
{
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);
	
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Task Management Menu","MN015",scScreenFunctions);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "MN015")
}
function SwapImage(imgImage)
{
	imgImage.style.cursor="hand";
	imgImage.style.zIndex="3";
}
function RestoreImage(imgImage)
{
	imgImage.style.zIndex="1";
}
/* Old button methods 
function frmScreen.btnTaskSummary.onclick()
{
	GoTaskSummary();
}
function frmScreen.btnTaskSearch.onclick()
{
	GoTaskSearch();
}
*/
function frmScreen.imgTaskSummary1.onmouseover()
{
	SwapImage(frmScreen.imgTaskSummary2);
}
function frmScreen.imgTaskSummary2.onmouseout()
{
	RestoreImage(frmScreen.imgTaskSummary2);
}
function frmScreen.imgTaskSummary2.onclick()
{
	GoTaskSummary();
}
function frmScreen.imgTaskSearch1.onmouseover()
{
	SwapImage(frmScreen.imgTaskSearch2);
}
function frmScreen.imgTaskSearch2.onmouseout()
{
	RestoreImage(frmScreen.imgTaskSearch2);
}
function frmScreen.imgTaskSearch2.onclick()
{
	GoTaskSearch();
}

function GoTaskSummary()
{
	frmTaskSummary.submit();
}

function GoTaskSearch()
{
	frmTaskSearch.submit();
}
function btnSubmit.onclick()
{
	frmToMN010.submit();
}  
-->
</script>
</BODY>
</HTML>
